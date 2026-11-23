# AD Scheduled Tasks Manager

An interactive PowerShell console (`sconfig`-style menu) to **inventory and delete Windows scheduled tasks** on all computers of an Active Directory domain: user workstations, servers and virtual machines.

> [Русская версия](README.md)

## Features

- Collects **all scheduled tasks** from every domain computer, or only from a selected OU.
- Separate tables for **user computers** and **servers / VMs** (classified by AD `OperatingSystem`).
- Choose what to scan: all computers, user computers only, or servers only. Scan **individual OUs** (including child OUs); results accumulate between scans.
- **Progress bar** while polling, `Esc` to abort (data already collected is kept).
- Summary view: "task → how many computers have it" (clients and servers separately).
- **Delete a task by name** from all or selected computers (all / clients only / servers only / pick manually).
- Delete individual rows (computer + task) picked in an `Out-GridView` window.
- System tasks under `\Microsoft\*` are hidden by default and trigger a warning if you try to delete them.
- `YES` confirmation for deletions, detailed log, CSV export.
- List of unreachable computers with the reason.

## Requirements

- Windows PowerShell 5.1 (developed and used with it; PowerShell 7 has not been tested).
- **ActiveDirectory** module (RSAT) on the machine that runs the script.
- An account with administrator rights on the target computers (e.g. a domain admin).
- Network access to targets via **WinRM (TCP 5985)** or **RPC/DCOM (TCP 135)**: WinRM is tried first, DCOM is the fallback.
- Target OS with the `ScheduledTasks` module (Windows 8 / Server 2012 or newer).
- `Out-GridView` is needed for manual row selection (not available on Server Core).

## Quick start

```powershell
git clone https://github.com/YOUR-USERNAME/ad-tasks-manager.git
cd ad-tasks-manager

# run (the domain is taken from $env:USERDNSDOMAIN)
powershell -ExecutionPolicy Bypass -File .\Manage-DomainScheduledTasks.ps1

# or with parameters
.\Manage-DomainScheduledTasks.ps1 -Domain corp.example.com -StaleDays 60
```

The program **shows the menu first**; nothing is scanned until you choose item `1`.

## Parameters

| Parameter     | Default                    | Description                                                   |
|---------------|----------------------------|---------------------------------------------------------------|
| `-Domain`     | `$env:USERDNSDOMAIN`       | DNS name of the AD domain                                     |
| `-StaleDays`  | `90`                       | Skip computers that have not logged on to the domain for N days |

## Menu

The menu is displayed in Russian (see the Russian README for a sample). Items: scan with current settings (1), switch computer type (2), choose OU (3), clear results (C), clients view (4), servers view (5), summary (6), unreachable computers (U), delete by name (7), delete picked rows (8), toggle system tasks (M), CSV export (E), exit (0). Hotkeys also work with the Russian keyboard layout.

## How it works

1. Enabled Windows computers are read from AD (`Get-ADComputer`), filtered by last logon (`-StaleDays`), type and OU.
2. Ports 5985 and 135 are probed in parallel.
3. CIM sessions are created (WinRM, DCOM fallback).
4. Tasks are read with `Get-ScheduledTask -CimSession` in batches of 8 sessions (with a progress bar).
5. Deletion uses `Unregister-ScheduledTask` in the corresponding CIM session.

## Safety and limitations

- **Deletion is irreversible.** Try it on a pilot OU first. The script lists what will be removed and requires typing `YES`.
- Tasks are read together with their run-as accounts and command lines, so the log (`DomainTaskManager.log`) and CSV exports may contain sensitive data. Do not publish them (they are already in `.gitignore`).
- Log and CSV files are written next to the script (or to the temp folder if the code was not run from a file).
- Powered-off or unreachable computers are listed under "Unreachable".
- Provided as is, without warranty (see the license).

## Related projects

- [AD Startup Manager](https://github.com/YOUR-USERNAME/ad-startup-manager): autostart entries.
- [AD Software Manager](https://github.com/YOUR-USERNAME/ad-software-manager): installed software.

## License

[MIT](LICENSE)
