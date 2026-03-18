# CLAUDE.md - SystemUpdatePro

## Overview
Enterprise/MSP-grade headless update automation. Auto-detects OEM (Dell/Lenovo/HP) and runs appropriate update tools alongside Windows Update and winget. v4.0.

## Tech Stack
- PowerShell 5.1, CLI/headless (no GUI)
- OEM tools: Dell Command Update CLI, LSUClient PS module, HP Image Assistant
- PSWindowsUpdate module, winget

## Key Details
- ~1,577 lines, single-file
- BitLocker-aware BIOS handling (suspends before BIOS updates)
- Battery safety check (blocks BIOS on battery)
- Lock file for concurrent execution prevention
- Post-reboot continuation via scheduled task
- Event Log integration for RMM visibility
- Log rotation, exponential backoff retry
- Logs to `C:\ProgramData\SystemUpdatePro\Logs`

## Build/Run
```powershell
# Run as Administrator
.\SystemUpdatePro.ps1
.\SystemUpdatePro.ps1 -SkipOEM -SkipWinget
.\SystemUpdatePro.ps1 -IncludeBIOS -Force
.\SystemUpdatePro.ps1 -RepairWindowsUpdate
```

## Key Parameters
`-SkipOEM`, `-SkipWindows`, `-SkipWinget`, `-IncludeBIOS`, `-BypassWSUS`, `-RepairWindowsUpdate`, `-CleanupAfter`, `-ContinueAfterReboot`, `-MaxRetries`, `-Reboot`, `-Force`

## Version
4.0
