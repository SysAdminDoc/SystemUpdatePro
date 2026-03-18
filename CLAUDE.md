# CLAUDE.md - SystemUpdatePro

## Overview
Enterprise/MSP-grade headless update automation. Auto-detects OEM (Dell/Lenovo/HP) and runs appropriate update tools alongside Windows Update and winget. v4.1.0.

## Tech Stack
- PowerShell 5.1, CLI/headless (no GUI)
- OEM tools: Dell Command Update CLI, LSUClient PS module, HP Image Assistant
- PSWindowsUpdate module, winget

## Key Details
- ~1,900 lines, single-file
- BitLocker-aware BIOS handling (suspends before BIOS updates)
- Battery safety check (blocks BIOS on battery)
- Lock file for concurrent execution prevention
- Post-reboot continuation via scheduled task
- Event Log integration for RMM visibility
- Log rotation, exponential backoff retry
- Logs to `C:\ProgramData\SystemUpdatePro\Logs`
- DryRun mode for safe preview of available updates
- HTML summary reports (Catppuccin Mocha dark theme, auto-opens in browser)
- Webhook notifications (Slack, Teams, generic JSON POST)
- Driver backup via Export-WindowsDriver before OEM updates
- Update history tracking (JSON log at `C:\ProgramData\SystemUpdatePro\update_history.json`)

## Build/Run
```powershell
# Run as Administrator
.\SystemUpdatePro.ps1
.\SystemUpdatePro.ps1 -SkipOEM -SkipWinget
.\SystemUpdatePro.ps1 -IncludeBIOS -Force
.\SystemUpdatePro.ps1 -RepairWindowsUpdate
.\SystemUpdatePro.ps1 -DryRun
.\SystemUpdatePro.ps1 -BackupDrivers
.\SystemUpdatePro.ps1 -ShowHistory -HistoryCount 20
.\SystemUpdatePro.ps1 -WebhookUrl "https://hooks.slack.com/services/..."
```

## Key Parameters
`-SkipOEM`, `-SkipWindows`, `-SkipWinget`, `-IncludeBIOS`, `-BypassWSUS`, `-RepairWindowsUpdate`, `-CleanupAfter`, `-ContinueAfterReboot`, `-DryRun`, `-BackupDrivers`, `-ShowHistory`, `-WebhookUrl`, `-HistoryCount`, `-MaxRetries`, `-Reboot`, `-Force`

## File Locations
| Path | Purpose |
|------|---------|
| `C:\ProgramData\SystemUpdatePro\Logs\` | Log files + HTML reports |
| `C:\ProgramData\SystemUpdatePro\update.lock` | Lock file |
| `C:\ProgramData\SystemUpdatePro\state.json` | State file (post-reboot) |
| `C:\ProgramData\SystemUpdatePro\update_history.json` | Update history log |
| `C:\ProgramData\SystemUpdatePro\DriverBackups\` | Driver backup snapshots |
| `C:\ProgramData\SystemUpdatePro\HPIA\` | HP Image Assistant |

## Version History
- 4.1.0 - DryRun mode, HTML reports, webhook notifications, driver backup, update history tracking
- 4.0.0 - Initial release with multi-OEM support, self-healing, BitLocker awareness

## Version
4.1.0
