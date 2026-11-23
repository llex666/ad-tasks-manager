#Requires -Version 5.1
<#
.SYNOPSIS
    Интерактивная консоль (в стиле sconfig) для аудита и удаления заданий
    планировщика на всех компьютерах домена.

.DESCRIPTION
    - Берёт из AD все включённые Windows-компьютеры (не заходившие в домен
      дольше $StaleDays дней пропускаются).
    - Проверяет доступность (TCP 5985 WinRM / 135 RPC), создаёт CIM-сессии
      (сначала WinRM, при неудаче DCOM).
    - Собирает все задания планировщика параллельно.
    - Разделяет компьютеры на «Клиенты» и «Серверы и ВМ» (по OperatingSystem).
    - Позволяет удалить задание на всех / выбранных компьютерах.

    Запускать от имени доменного администратора, на машине с RSAT (модуль ActiveDirectory).
    Файл сохранён в UTF-8 с BOM, чтобы PowerShell 5.1 корректно читал кириллицу.

.EXAMPLE
    .\Manage-DomainScheduledTasks.ps1 -Domain corp.example.com -StaleDays 60
#>

[CmdletBinding()]
param(
    # DNS-имя домена AD (по умолчанию - домен текущего пользователя)
    [string]$Domain = $(if ($env:USERDNSDOMAIN) { $env:USERDNSDOMAIN } else { 'example.local' }),

    # Пропускать компьютеры, не заходившие в домен более N дней
    [int]$StaleDays = 90
)

# ----------------------------- НАСТРОЙКИ ------------------------------------
$LogBase   = if ($PSScriptRoot) { $PSScriptRoot } else { $env:TEMP }
$LogPath   = Join-Path $LogBase 'DomainTaskManager.log'
# ----------------------------------------------------------------------------

Import-Module ActiveDirectory -ErrorAction Stop

$script:Computers = @()
$script:CompMap   = @{}
$script:Sessions  = @{}
$script:Tasks     = @()
$script:ShowMS    = $false          # показывать системные задания \Microsoft\*
$script:LastScan  = $null
$script:ScanType  = 'All'      # All | Client | Server (что сканировать)
$script:ScanOU    = $null      # DN выбранной OU; $null = весь домен

# ------------------------------ Вспомогательное -----------------------------

function Write-Log {
    param([string]$Message)
    $line = '{0:yyyy-MM-dd HH:mm:ss} [{1}] {2}' -f (Get-Date), $env:USERNAME, $Message
    Add-Content -Path $LogPath -Value $line -Encoding UTF8
}

function Read-Key {
    try { [Console]::ReadKey($true).KeyChar.ToString() }
    catch { Read-Host }   # ISE и т.п. — без ReadKey
}

function Wait-Key {
    Write-Host ''
    Write-Host 'Нажмите любую клавишу для возврата в меню...' -ForegroundColor DarkGray
    [void](Read-Key)
}

function Limit-Text {
    param($Text, [int]$Max = 70)
    $s = [string]$Text
    if ($s.Length -gt $Max) { $s.Substring(0, $Max - 1) + '…' } else { $s }
}

function Write-Header {
    param([string]$Title)
    Clear-Host
    Write-Host ('=' * 78) -ForegroundColor Cyan
    Write-Host "  $Title" -ForegroundColor Cyan
    Write-Host ('=' * 78) -ForegroundColor Cyan
    Write-Host ''
}

function Test-PortAsync {
    param([string[]]$Names, [int]$Port, [int]$TimeoutMs = 3000)
    $clients = @{}; $tasks = @{}
    foreach ($n in $Names) {
        try {
            $c = New-Object System.Net.Sockets.TcpClient
            $clients[$n] = $c
            $tasks[$n]   = $c.ConnectAsync($n, $Port)
        } catch { }
    }
    try {
        [void][System.Threading.Tasks.Task]::WaitAll(
            [System.Threading.Tasks.Task[]]@($tasks.Values), $TimeoutMs)
    } catch { }
    $res = @{}
    foreach ($n in $Names) {
        $res[$n] = ($tasks.ContainsKey($n) -and
                    $tasks[$n].Status -eq 'RanToCompletion' -and
                    $clients[$n].Connected)
        if ($clients.ContainsKey($n)) { $clients[$n].Dispose() }
    }
    $res
}

function Get-VisibleTasks {
    param([string]$Type)
    $items = $script:Tasks
    if (-not $script:ShowMS) { $items = $items | Where-Object { $_.Path -notlike '\Microsoft\*' } }
    if ($Type)               { $items = $items | Where-Object { $_.Type -eq $Type } }
    @($items)
}

function Get-TaskGroups {
    param([string]$Filter)
    $items = Get-VisibleTasks
    if ($Filter) {
        $items = $items | Where-Object { $_.Task -like "*$Filter*" -or $_.Path -like "*$Filter*" }
    }
    @($items | Group-Object -Property Key | Sort-Object Name)
}

function Show-GroupTable {
    param($Groups)
    $i = 0
    $view = foreach ($g in $Groups) {
        $i++
        [pscustomobject][ordered]@{
            '#'       = $i
            'Задание' = $g.Name
            'Клиенты' = @($g.Group | Where-Object { $_.Type -eq 'Client' }).Count
            'Серверы' = @($g.Group | Where-Object { $_.Type -eq 'Server' }).Count
            'Всего'   = $g.Count
        }
    }
    $view | Format-Table -AutoSize | Out-Host
}

# ------------------------------- Сканирование -------------------------------

function Get-ScanTypeName {
    switch ($script:ScanType) {
        'Client' { 'только компьютеры пользователей' }
        'Server' { 'только серверы и ВМ' }
        default  { 'все компьютеры' }
    }
}

function Switch-ScanType {
    $script:ScanType = switch ($script:ScanType) { 'All' { 'Client' } 'Client' { 'Server' } default { 'All' } }
}

function Clear-Results {
    foreach ($s in $script:Sessions.Values) { try { Remove-CimSession $s -ErrorAction SilentlyContinue } catch { } }
    $script:Sessions = @{}; $script:Computers = @(); $script:Tasks = @(); $script:CompMap = @{}; $script:LastScan = $null
}

function Update-CompMap {
    $script:CompMap = @{}
    foreach ($c in $script:Computers) { $script:CompMap[$c.Name] = $c }
}

function Test-EscPressed {
    try {
        while ([Console]::KeyAvailable) {
            if ([Console]::ReadKey($true).Key -eq 'Escape') { return $true }
        }
    } catch { }
    $false
}

function Get-ScanTargets {
    $limit = (Get-Date).AddDays(-$StaleDays)
    $p = @{
        Server = $Domain; Filter = 'Enabled -eq $true'
        Properties = 'OperatingSystem', 'LastLogonDate'; ErrorAction = 'Stop'
    }
    if ($script:ScanOU) { $p.SearchBase = $script:ScanOU; $p.SearchScope = 'Subtree' }
    Get-ADComputer @p | Where-Object {
        $_.OperatingSystem -like 'Windows*' -and $_.LastLogonDate -gt $limit -and
        ($script:ScanType -eq 'All' -or
         ($script:ScanType -eq 'Server' -and $_.OperatingSystem -like '*Server*') -or
         ($script:ScanType -eq 'Client' -and $_.OperatingSystem -notlike '*Server*'))
    }
}

function Select-ScanOU {
    Write-Header 'Выбор OU для сканирования'
    Write-Host 'Загружаю список OU из AD...' -ForegroundColor Cyan
    try { $ous = @(Get-ADOrganizationalUnit -Server $Domain -Filter * -ErrorAction Stop | Sort-Object DistinguishedName) }
    catch { Write-Host "Ошибка AD: $($_.Exception.Message)" -ForegroundColor Red; Wait-Key; return }

    Write-Host ("Текущая область: {0}" -f $(if ($script:ScanOU) { $script:ScanOU } else { "весь домен $Domain" }))
    Write-Host 'Сканируется выбранная OU вместе со всеми вложенными.' -ForegroundColor DarkGray
    Write-Host ''
    $f = Read-Host 'Фильтр по части имени OU (Enter — показать все)'
    $list = if ($f) { @($ous | Where-Object { $_.DistinguishedName -like "*$f*" }) } else { $ous }
    if ($list.Count -eq 0) { Write-Host 'OU не найдены.' -ForegroundColor Yellow; Wait-Key; return }

    $i = 0
    $view = foreach ($o in $list) { $i++; [pscustomobject]@{ '#' = $i; 'OU' = $o.DistinguishedName } }
    $fmt = $view | Format-Table -AutoSize
    try { $fmt | Out-Host -Paging } catch { $fmt | Out-Host }
    Write-Host ''
    $sel = (Read-Host 'Номер OU, либо DN целиком, либо 0 — весь домен (Enter — без изменений)').Trim()
    if (-not $sel) { return }
    if ($sel -eq '0') { $script:ScanOU = $null; return }
    if ($sel -match '^\d+$') {
        $n = [int]$sel
        if ($n -ge 1 -and $n -le $list.Count) { $script:ScanOU = $list[$n - 1].DistinguishedName }
        else { Write-Host 'Неверный номер.' -ForegroundColor Yellow; Wait-Key }
        return
    }
    if ($sel -match '^(OU|CN|DC)=') {
        try { $null = Get-ADObject -Identity $sel -Server $Domain -ErrorAction Stop; $script:ScanOU = $sel }
        catch { Write-Host 'Такой объект в AD не найден.' -ForegroundColor Red; Wait-Key }
        return
    }
    Write-Host 'Неверный ввод.' -ForegroundColor Yellow; Wait-Key
}

function Invoke-Scan {
    Write-Header 'Сканирование'
    Write-Host ("Тип компьютеров : {0}" -f (Get-ScanTypeName))
    Write-Host ("Область (OU)    : {0}" -f $(if ($script:ScanOU) { $script:ScanOU } else { "весь домен $Domain" }))
    Write-Host 'Прервать опрос можно клавишей Esc.' -ForegroundColor DarkGray
    Write-Host ''
    $act = 'Сканирование планировщика заданий'
    $sw  = [Diagnostics.Stopwatch]::StartNew()

    # ---- Шаг 1: список компьютеров из AD
    Write-Progress -Id 1 -Activity $act -Status 'Шаг 1 из 4: получаю список компьютеров из AD...' -PercentComplete 0
    try { $ad = @(Get-ScanTargets | Sort-Object Name) }
    catch {
        Write-Progress -Id 1 -Activity $act -Completed
        Write-Host "Ошибка AD: $($_.Exception.Message)" -ForegroundColor Red
        Wait-Key; return
    }
    if ($ad.Count -eq 0) {
        Write-Progress -Id 1 -Activity $act -Completed
        Write-Host 'Подходящих компьютеров не найдено.' -ForegroundColor Yellow
        Wait-Key; return
    }
    Write-Host "Найдено компьютеров: $($ad.Count)" -ForegroundColor Cyan

    # прежние данные (и сессии) заменяются только для сканируемых компьютеров
    $newSet = @{}; foreach ($c in $ad) { $newSet[$c.Name] = $true }
    foreach ($n in @($newSet.Keys)) {
        if ($script:Sessions.ContainsKey($n)) {
            try { Remove-CimSession $script:Sessions[$n] -ErrorAction SilentlyContinue } catch { }
            $script:Sessions.Remove($n)
        }
    }
    $script:Computers = @($script:Computers | Where-Object { -not $newSet.ContainsKey($_.Name) })
    $script:Tasks     = @($script:Tasks     | Where-Object { -not $newSet.ContainsKey($_.Computer) })

    $new = @(foreach ($c in $ad) {
        [pscustomobject]@{
            Name      = $c.Name
            OS        = $c.OperatingSystem
            Type      = $(if ($c.OperatingSystem -like '*Server*') { 'Server' } else { 'Client' })
            Online    = $false
            Connected = $false
            TaskCount = 0
            Error     = ''
        }
    })
    $script:Computers = @($script:Computers) + $new
    Update-CompMap

    # ---- Шаг 2: проверка доступности (WinRM 5985 / RPC 135)
    $names = @($new | Select-Object -ExpandProperty Name)
    $winrm = @{}; $rpc = @{}
    $chunk = 100
    for ($i = 0; $i -lt $names.Count; $i += $chunk) {
        $end  = [math]::Min($i + $chunk, $names.Count)
        $part = @($names[$i..($end - 1)])
        $r1 = Test-PortAsync -Names $part -Port 5985
        $r2 = Test-PortAsync -Names $part -Port 135
        foreach ($k in $r1.Keys) { $winrm[$k] = $r1[$k] }
        foreach ($k in $r2.Keys) { $rpc[$k]   = $r2[$k] }
        Write-Progress -Id 1 -Activity $act `
            -Status ("Шаг 2 из 4: проверка доступности: {0} из {1}" -f $end, $names.Count) `
            -PercentComplete ([int](100 * $end / $names.Count))
    }

    # ---- Шаг 3: подключение (CIM-сессии)
    $cancelled = $false
    $i = 0
    foreach ($c in $new) {
        $i++
        Write-Progress -Id 1 -Activity $act `
            -Status ("Шаг 3 из 4: подключение {0} из {1}: {2} | Esc — прервать" -f $i, $new.Count, $c.Name) `
            -PercentComplete ([int](100 * $i / $new.Count))
        if (Test-EscPressed) { $cancelled = $true; break }
        $c.Online = [bool]($winrm[$c.Name] -or $rpc[$c.Name])
        if (-not $c.Online) { $c.Error = 'Недоступен (выключен или порты 5985/135 закрыты)'; continue }

        $s = $null
        if ($winrm[$c.Name]) {
            try { $s = New-CimSession -ComputerName $c.Name -OperationTimeoutSec 60 -ErrorAction Stop }
            catch { $c.Error = (($_.Exception.Message) -split "`n")[0] }
        }
        if (-not $s -and $rpc[$c.Name]) {
            try {
                $opt = New-CimSessionOption -Protocol Dcom
                $s = New-CimSession -ComputerName $c.Name -SessionOption $opt -OperationTimeoutSec 60 -ErrorAction Stop
            } catch { $c.Error = (($_.Exception.Message) -split "`n")[0] }
        }
        if ($s) { $script:Sessions[$c.Name] = $s; $c.Connected = $true; $c.Error = '' }
    }
    if ($cancelled) {
        foreach ($c in $new) { if (-not $c.Connected -and -not $c.Error) { $c.Error = 'Опрос прерван' } }
    }

    # ---- Шаг 4: чтение заданий пачками
    $newSess = @($new | Where-Object { $script:Sessions.ContainsKey($_.Name) } | ForEach-Object { $script:Sessions[$_.Name] })
    $raw = @(); $errs = $null; $readSet = @{}
    $batch = 8
    for ($i = 0; $i -lt $newSess.Count -and -not $cancelled; $i += $batch) {
        if (Test-EscPressed) { $cancelled = $true; break }
        $end = [math]::Min($i + $batch, $newSess.Count)
        $b   = @($newSess[$i..($end - 1)])
        Write-Progress -Id 1 -Activity $act `
            -Status ("Шаг 4 из 4: чтение заданий: {0} из {1} | {2:mm\:ss} | Esc — прервать" -f $i, $newSess.Count, $sw.Elapsed) `
            -PercentComplete ([int](100 * $i / [math]::Max(1, $newSess.Count)))
        $raw += @(Get-ScheduledTask -CimSession $b -ErrorAction SilentlyContinue -ErrorVariable +errs)
        foreach ($s in $b) { $readSet[[string]$s.ComputerName] = $true }
    }
    if ($cancelled) {
        foreach ($c in $new) {
            if ($c.Connected -and -not $readSet.ContainsKey($c.Name)) { $c.Connected = $false; $c.Error = 'Опрос прерван' }
        }
    }

    $list = foreach ($t in $raw) {
        $comp = $script:CompMap[[string]$t.PSComputerName]
        if (-not $comp) { continue }
        $runAs = if ($t.Principal.UserId) { $t.Principal.UserId } else { $t.Principal.GroupId }
        $actn = (@($t.Actions) | ForEach-Object { ("$($_.Execute) $($_.Arguments)").Trim() }) -join ' ; '
        [pscustomobject]@{
            Computer = $comp.Name
            Type     = $comp.Type
            Path     = $t.TaskPath
            Task     = $t.TaskName
            State    = [string]$t.State
            RunAs    = $runAs
            Action   = $actn
            Key      = ($t.TaskPath + $t.TaskName)
        }
    }
    $script:Tasks = @($script:Tasks) + @($list)

    foreach ($g in ($list | Group-Object Computer)) {
        if ($script:CompMap.ContainsKey($g.Name)) { $script:CompMap[$g.Name].TaskCount = $g.Count }
    }
    foreach ($e in $errs) {
        $pc = $null
        try { $pc = [string]$e.OriginInfo.PSComputerName } catch { }
        if ($pc -and $script:CompMap.ContainsKey($pc) -and $script:CompMap[$pc].TaskCount -eq 0) {
            $script:CompMap[$pc].Connected = $false
            $script:CompMap[$pc].Error = Limit-Text ((($e.Exception.Message) -split "`n")[0]) 120
        }
    }

    Write-Progress -Id 1 -Activity $act -Completed
    $script:LastScan = Get-Date
    $okCount = @($new | Where-Object { $_.Connected }).Count
    $failed  = $new.Count - $okCount
    Write-Log ("Скан [{0}; {1}]: компьютеров {2}, опрошено {3}, ошибок/недоступно {4}, заданий {5}{6}" -f
        (Get-ScanTypeName), $(if ($script:ScanOU) { $script:ScanOU } else { 'весь домен' }),
        $new.Count, $okCount, $failed, @($list).Count, $(if ($cancelled) { ' (прервано)' } else { '' }))
    Write-Host ''
    Write-Host ("Готово за {0:mm\:ss}{1}. Компьютеров: {2}, опрошено: {3}, недоступно/ошибок: {4}, заданий: {5}." -f
        $sw.Elapsed, $(if ($cancelled) { ' (прервано пользователем)' } else { '' }), $new.Count, $okCount, $failed, @($list).Count) -ForegroundColor Green
    Write-Host 'Список проблемных компьютеров — в разделе «Недоступные».' -ForegroundColor DarkGray
    Wait-Key
}

# --------------------------------- Просмотр ---------------------------------

function Show-Tasks {
    param([string]$Type, [string]$Title)
    Write-Header $Title
    $comps = @($script:Computers | Where-Object { $_.Type -eq $Type })
    $ok    = @($comps | Where-Object { $_.Connected })
    Write-Host ("Компьютеров: {0}, опрошено: {1}, недоступно: {2}" -f $comps.Count, $ok.Count, ($comps.Count - $ok.Count))
    $items = @(Get-VisibleTasks $Type | Sort-Object Computer, Path, Task)
    Write-Host ("Заданий: {0}  (системные \Microsoft\*: {1})" -f $items.Count, $(if ($script:ShowMS) { 'показаны' } else { 'скрыты' }))
    Write-Host ''
    if ($items.Count -eq 0) { Write-Host 'Нет данных. Выполните сканирование (пункт 1).' -ForegroundColor Yellow; Wait-Key; return }

    $fmt = $items | Format-Table -GroupBy Computer -Property `
        @{n='Путь'; e={ $_.Path }},
        @{n='Задание'; e={ $_.Task }},
        @{n='Состояние'; e={ $_.State }},
        @{n='Запуск от'; e={ $_.RunAs }},
        @{n='Действие'; e={ Limit-Text $_.Action 60 }} -AutoSize -Wrap
    try { $fmt | Out-Host -Paging } catch { $fmt | Out-Host }
    Wait-Key
}

function Show-Summary {
    Write-Header 'Сводка: какие задания на скольких компьютерах'
    $f = Read-Host 'Фильтр по имени задания (Enter — все)'
    $groups = Get-TaskGroups $f
    if ($groups.Count -eq 0) { Write-Host 'Ничего не найдено.' -ForegroundColor Yellow; Wait-Key; return }
    Show-GroupTable $groups
    Wait-Key
}

function Show-Unreachable {
    Write-Header 'Компьютеры, с которых не удалось получить задания'
    $bad = @($script:Computers | Where-Object { -not $_.Connected })
    if ($bad.Count -eq 0) { Write-Host 'Все компьютеры опрошены.' -ForegroundColor Green }
    else {
        $bad | Sort-Object Type, Name |
            Format-Table Name, Type, @{n='ОС';e={ Limit-Text $_.OS 28 }}, Online, @{n='Причина';e={ Limit-Text $_.Error 60 }} -AutoSize |
            Out-Host
    }
    Wait-Key
}

# --------------------------------- Удаление ---------------------------------

function Remove-TaskRecords {
    param($Records)
    $Records = @($Records)
    if ($Records.Count -eq 0) { Write-Host 'Нечего удалять.' -ForegroundColor Yellow; return }

    Write-Host ''
    Write-Host 'Будут удалены задания:' -ForegroundColor Yellow
    $Records | Sort-Object Computer, Path, Task |
        Format-Table Computer, Type, Path, Task -AutoSize | Out-Host
    if ($Records | Where-Object { $_.Path -like '\Microsoft\*' }) {
        Write-Host 'ВНИМАНИЕ: среди заданий есть системные (\Microsoft\*) — удаление может повредить Windows!' -ForegroundColor Red
    }
    $ans = Read-Host ("Удалить {0} заданий на {1} компьютерах? Введите YES для подтверждения" -f
                      $Records.Count, @($Records | Select-Object -ExpandProperty Computer -Unique).Count)
    if ($ans -cne 'YES') { Write-Host 'Отменено.' -ForegroundColor Yellow; return }

    $removed = @(); $fail = 0
    foreach ($r in $Records) {
        $sess = $script:Sessions[$r.Computer]
        try {
            if (-not $sess) { throw 'нет CIM-сессии' }
            Unregister-ScheduledTask -TaskName $r.Task -TaskPath $r.Path -CimSession $sess -Confirm:$false -ErrorAction Stop
            Write-Host ("  [OK]  {0}: {1}{2}" -f $r.Computer, $r.Path, $r.Task) -ForegroundColor Green
            Write-Log   ("УДАЛЕНО {0}: {1}{2}" -f $r.Computer, $r.Path, $r.Task)
            $removed += $r
        } catch {
            $fail++
            $msg = (($_.Exception.Message) -split "`n")[0]
            Write-Host ("  [ERR] {0}: {1}{2} — {3}" -f $r.Computer, $r.Path, $r.Task, $msg) -ForegroundColor Red
            Write-Log   ("ОШИБКА {0}: {1}{2} — {3}" -f $r.Computer, $r.Path, $r.Task, $msg)
        }
    }
    if ($removed.Count -gt 0) {
        $script:Tasks = @($script:Tasks | Where-Object { $removed -notcontains $_ })
        foreach ($g in ($removed | Group-Object Computer)) {
            $script:CompMap[$g.Name].TaskCount = [math]::Max(0, $script:CompMap[$g.Name].TaskCount - $g.Count)
        }
    }
    Write-Host ''
    Write-Host ("Удалено: {0}, ошибок: {1}. Лог: {2}" -f $removed.Count, $fail, $LogPath) -ForegroundColor Cyan
}

function Test-GridView {
    if (-not (Get-Command Out-GridView -ErrorAction SilentlyContinue)) {
        Write-Host 'Out-GridView недоступен (Server Core или нет компонента ISE/GraphicalTools).' -ForegroundColor Red
        return $false
    }
    $true
}

function Remove-ByName {
    Write-Header 'Удаление задания по имени'
    $f = Read-Host 'Фильтр по имени задания (Enter — все)'
    $groups = Get-TaskGroups $f
    if ($groups.Count -eq 0) { Write-Host 'Ничего не найдено.' -ForegroundColor Yellow; Wait-Key; return }
    Show-GroupTable $groups

    $sel = Read-Host 'Номера заданий через запятую (Enter — отмена)'
    if (-not $sel) { return }
    $idx = @($sel -split '[,\s]+' | Where-Object { $_ -match '^\d+$' } | ForEach-Object { [int]$_ })
    $records = @()
    foreach ($n in $idx) {
        if ($n -ge 1 -and $n -le $groups.Count) { $records += @($groups[$n - 1].Group) }
    }
    if ($records.Count -eq 0) { Write-Host 'Неверный выбор.' -ForegroundColor Yellow; Wait-Key; return }

    Write-Host ''
    Write-Host 'Область удаления:' -ForegroundColor Cyan
    Write-Host '  [A] все компьютеры, где есть задание'
    Write-Host '  [C] только компьютеры пользователей'
    Write-Host '  [S] только серверы и ВМ'
    Write-Host '  [G] выбрать компьютеры вручную (окно Out-GridView)'
    Write-Host '  [любая другая] отмена'
    $k = (Read-Key).ToUpper()
    switch ($k) {
        { $_ -in 'A','Ф' } { $target = $records }
        { $_ -in 'C','С' } { $target = @($records | Where-Object { $_.Type -eq 'Client' }) }
        { $_ -in 'S','Ы' } { $target = @($records | Where-Object { $_.Type -eq 'Server' }) }
        { $_ -in 'G','П' } {
            if (-not (Test-GridView)) { Wait-Key; return }
            $target = @($records | Sort-Object Computer | Out-GridView -PassThru `
                        -Title 'Выделите компьютеры (Ctrl/Shift), затем OK')
        }
        default { Write-Host 'Отменено.' -ForegroundColor Yellow; Wait-Key; return }
    }
    Remove-TaskRecords $target
    Wait-Key
}

function Remove-ByGrid {
    Write-Header 'Удаление выбранных строк (Computer + Task)'
    if (-not (Test-GridView)) { Wait-Key; return }
    $items = @(Get-VisibleTasks | Sort-Object Type, Computer, Path, Task)
    if ($items.Count -eq 0) { Write-Host 'Нет данных. Выполните сканирование.' -ForegroundColor Yellow; Wait-Key; return }
    Write-Host 'Откроется окно: отметьте строки (Ctrl/Shift), нажмите OK.' -ForegroundColor Cyan
    $picked = @($items | Out-GridView -PassThru -Title 'Выберите задания для удаления')
    Remove-TaskRecords $picked
    Wait-Key
}

# --------------------------------- Экспорт ----------------------------------

function Export-Data {
    Write-Header 'Экспорт в CSV'
    if ($script:Tasks.Count -eq 0) { Write-Host 'Нет данных.' -ForegroundColor Yellow; Wait-Key; return }
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $f1 = Join-Path $LogBase "Tasks_$stamp.csv"
    $f2 = Join-Path $LogBase "Computers_$stamp.csv"
    $script:Tasks     | Select-Object Computer, Type, Path, Task, State, RunAs, Action |
        Export-Csv $f1 -NoTypeInformation -Encoding UTF8
    $script:Computers | Select-Object Name, Type, OS, Online, Connected, TaskCount, Error |
        Export-Csv $f2 -NoTypeInformation -Encoding UTF8
    Write-Host "Сохранено:`n  $f1`n  $f2" -ForegroundColor Green
    Wait-Key
}

# ----------------------------------- Меню -----------------------------------

function Show-MainMenu {
    Write-Header "Планировщик заданий домена $Domain"
    $cl = @($script:Computers | Where-Object { $_.Type -eq 'Client' })
    $sv = @($script:Computers | Where-Object { $_.Type -eq 'Server' })
    $on = @($script:Computers | Where-Object { $_.Connected })
    if ($script:Computers.Count -eq 0) {
        Write-Host '  Данных пока нет. Задайте параметры (пункты 2 и 3) и запустите сканирование (пункт 1).' -ForegroundColor Yellow
    } else {
        Write-Host ("  Просканировано компьютеров: {0} (клиенты: {1}, серверы/ВМ: {2}), опрошено успешно: {3}" -f
                    $script:Computers.Count, $cl.Count, $sv.Count, $on.Count)
        Write-Host ("  Заданий: {0}   Системные \Microsoft\*: {1}" -f
                    $script:Tasks.Count, $(if ($script:ShowMS) { 'показаны' } else { 'скрыты' }))
        if ($script:LastScan) { Write-Host ("  Последнее сканирование: {0:dd.MM.yyyy HH:mm:ss}" -f $script:LastScan) }
    }
    Write-Host ''
    Write-Host '  Сканирование' -ForegroundColor Cyan
    Write-Host '   1) Запустить сканирование с текущими параметрами' -ForegroundColor Green
    Write-Host ("   2) Тип компьютеров: [{0}]  (нажмите, чтобы переключить)" -f (Get-ScanTypeName))
    Write-Host ("   3) Область (OU): [{0}]" -f $(if ($script:ScanOU) { $script:ScanOU } else { "весь домен $Domain" }))
    Write-Host '   C) Очистить результаты'
    Write-Host ''
    Write-Host '  Просмотр' -ForegroundColor Cyan
    Write-Host '   4) Задания: компьютеры пользователей'
    Write-Host '   5) Задания: серверы и виртуальные машины'
    Write-Host '   6) Сводка по заданиям (имя → на скольких ПК)'
    Write-Host '   U) Недоступные компьютеры'
    Write-Host ''
    Write-Host '  Изменение' -ForegroundColor Cyan
    Write-Host '   7) Удалить задание по имени (все / выборочно)' -ForegroundColor Yellow
    Write-Host '   8) Удалить выбранные строки (Out-GridView)' -ForegroundColor Yellow
    Write-Host ''
    Write-Host '  Прочее' -ForegroundColor Cyan
    Write-Host ("   M) Системные задания \Microsoft\*: {0}" -f $(if ($script:ShowMS) { 'показаны' } else { 'скрыты' }))
    Write-Host '   E) Экспорт в CSV'
    Write-Host '   0) Выход'
    Write-Host ''
    Write-Host '  Выберите пункт: ' -NoNewline
}

# Программа сначала показывает меню, сканирование запускается только по пункту 1.
while ($true) {
    Show-MainMenu
    $key = (Read-Key).ToUpper()
    switch ($key) {
        '1' { Invoke-Scan }
        '2' { Switch-ScanType }
        '3' { Select-ScanOU }
        '4' { Show-Tasks -Type 'Client' -Title 'Задания планировщика: компьютеры пользователей' }
        '5' { Show-Tasks -Type 'Server' -Title 'Задания планировщика: серверы и виртуальные машины' }
        '6' { Show-Summary }
        '7' { Remove-ByName }
        '8' { Remove-ByGrid }
        { $_ -in 'U','Г' } { Show-Unreachable }
        { $_ -in 'M','Ь' } { $script:ShowMS = -not $script:ShowMS }
        { $_ -in 'E','У' } { Export-Data }
        { $_ -in 'C','С' } { Clear-Results }
        '0' {
            foreach ($s in $script:Sessions.Values) { try { Remove-CimSession $s -ErrorAction SilentlyContinue } catch { } }
            Clear-Host
            return
        }
    }
}
