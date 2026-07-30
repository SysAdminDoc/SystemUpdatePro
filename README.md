<p align="center"><img src="icon.png" width="128" height="128" alt="SystemUpdatePro"></p>

# SystemUpdatePro

<p align="center">
  <img src="https://img.shields.io/badge/PowerShell-5.1+-blue?logo=powershell&logoColor=white" alt="PowerShell 5.1+">
  <img src="https://img.shields.io/badge/Windows-10%20|%2011%20|%20Server-0078D6?logo=windows&logoColor=white" alt="Windows">
  <img src="https://img.shields.io/badge/License-MIT-green" alt="License">
  <img src="https://img.shields.io/badge/Version-4.1.0-orange" alt="Version">
</p>

**Enterprise-grade, bulletproof system update utility for MSPs and IT professionals.**

SystemUpdatePro is a fully automated, self-healing PowerShell script that handles OEM driver/BIOS updates (Dell, Lenovo, HP), Windows Updates, and application updates via Winget---all without user interaction.

---

## Features

### Multi-OEM Support
| Manufacturer | Tool Used | Auto-Install |
|--------------|-----------|--------------|
| Dell / Alienware | Dell Command Update CLI | via Winget |
| Lenovo | LSUClient PowerShell Module | via PSGallery |
| HP | HP Image Assistant | Auto-download |
| Other | Windows Update + Winget only | N/A |

### Self-Healing Capabilities
- **Winget Auto-Install**: Automatically installs Winget on Windows 10 with all dependencies (VCLibs, UI.Xaml)
- **Windows Update Repair**: Resets WU components, re-registers 30+ DLLs, clears cache
- **Service Recovery**: Detects and repairs broken OEM services
- **Retry Logic**: Exponential backoff with configurable retry attempts

### Safety Features
- **Lock File**: Prevents concurrent execution with stale lock detection
- **Disk Space Check**: Blocks execution if insufficient space available
- **Battery Protection**: Blocks BIOS updates when on battery power
- **BitLocker Awareness**: Handles BitLocker suspension for BIOS updates (Dell auto-suspends; Lenovo/HP skip BIOS when encrypted)
- **Pending Reboot Detection**: Checks 5 different sources for pending reboots
- **DryRun Mode**: Preview all available updates without installing anything
- **Driver Backup**: Export current drivers before installing updates for rollback capability

### Enterprise Integration
- **Event Log**: Writes to Windows Application log for RMM/SIEM visibility
- **Exit Codes**: Granular exit codes for automation pipelines
- **WSUS Bypass**: Option to bypass WSUS and connect directly to Microsoft
- **Post-Reboot Continuation**: Versioned, bounded state machine resumes update stages with the original run settings
- **Log Rotation**: Automatic cleanup of old log files
- **HTML Reports**: Responsive operations-dashboard report with update channels, device profile, exceptions, and print styles
- **Webhook Notifications**: Send completion status to Slack, Teams, or any generic webhook
- **Update History**: Schema-versioned JSON history with stage/item outcomes, provider codes, and evidence-delivery status

---

## Requirements

- **OS**: Windows 10, Windows 11, or Windows Server 2016+
- **PowerShell**: 5.1 or higher
- **Privileges**: Administrator
- **Network**: Internet access required

---

## Installation

### Option 1: Direct Download
```powershell
# Download the script
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/SysAdminDoc/SystemUpdatePro/main/SystemUpdatePro.ps1" -OutFile "SystemUpdatePro.ps1"

# Run it
.\SystemUpdatePro.ps1
```

### Option 2: Clone Repository
```powershell
git clone https://github.com/SysAdminDoc/SystemUpdatePro.git
cd SystemUpdatePro
.\SystemUpdatePro.ps1
```

---

## Usage

### Basic Usage
```powershell
# Full update: OEM drivers + Windows Updates + Winget upgrades
.\SystemUpdatePro.ps1

# Include BIOS updates with auto-reboot
.\SystemUpdatePro.ps1 -IncludeBIOS -Reboot

# Windows Updates only
.\SystemUpdatePro.ps1 -SkipOEM -SkipWinget

# OEM updates only
.\SystemUpdatePro.ps1 -SkipWindows -SkipWinget
```

### Dry Run (Preview Mode)
```powershell
# See what updates are available without installing anything
.\SystemUpdatePro.ps1 -DryRun

# Dry run with BIOS check included
.\SystemUpdatePro.ps1 -DryRun -IncludeBIOS
```

### Driver Backup
```powershell
# Backup drivers before updating
.\SystemUpdatePro.ps1 -BackupDrivers

# Backup drivers + include BIOS updates
.\SystemUpdatePro.ps1 -BackupDrivers -IncludeBIOS -Reboot
```

### Webhook Notifications
```powershell
# Notify Slack on completion
.\SystemUpdatePro.ps1 -WebhookUrl "https://hooks.slack.com/services/T00/B00/xxx"

# Notify Microsoft Teams
.\SystemUpdatePro.ps1 -WebhookUrl "https://outlook.office.com/webhook/..."

# Generic webhook (any JSON-accepting endpoint)
.\SystemUpdatePro.ps1 -WebhookUrl "https://your-api.example.com/webhook"
```

### Update History
```powershell
# Show last 10 update runs
.\SystemUpdatePro.ps1 -ShowHistory

# Show last 25 update runs
.\SystemUpdatePro.ps1 -ShowHistory -HistoryCount 25
```

### Advanced Usage
```powershell
# Full provisioning workflow with post-reboot continuation
.\SystemUpdatePro.ps1 -IncludeBIOS -Reboot -ContinueAfterReboot -CleanupAfter

# Repair broken Windows Update then run updates
.\SystemUpdatePro.ps1 -RepairWindowsUpdate -BypassWSUS

# Irreversibly remove superseded component versions (prevents update uninstall)
.\SystemUpdatePro.ps1 -ResetComponentBase

# Force run despite warnings (low disk, pending reboot, battery)
.\SystemUpdatePro.ps1 -Force -IncludeBIOS

# Custom configuration
.\SystemUpdatePro.ps1 -MaxRetries 5 -MaxUpdatePasses 5 -MinDiskSpaceGB 20 -LogRetentionDays 60

# Kitchen sink: backup drivers, dry run with webhook
.\SystemUpdatePro.ps1 -DryRun -BackupDrivers -WebhookUrl "https://hooks.slack.com/services/..."
```

---

## Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-SkipOEM` | Switch | False | Skip OEM-specific driver/firmware updates |
| `-SkipWindows` | Switch | False | Skip Windows Update |
| `-SkipWinget` | Switch | False | Skip Winget upgrade all |
| `-IncludeBIOS` | Switch | False | Include BIOS updates (requires AC power) |
| `-BypassWSUS` | Switch | False | Bypass WSUS, connect directly to Microsoft |
| `-RepairWindowsUpdate` | Switch | False | Repair Windows Update components before updating |
| `-CleanupAfter` | Switch | False | Run standard DISM cleanup while retaining installed-update rollback |
| `-ResetComponentBase` | Switch | False | Run irreversible `/ResetBase`; installed Windows updates can no longer be uninstalled |
| `-ContinueAfterReboot` | Switch | False | Resume update stages after reboot with the original run settings (maximum 3 attempts) |
| `-DryRun` | Switch | False | Preview updates without installing |
| `-BackupDrivers` | Switch | False | Export current drivers before updating |
| `-ShowHistory` | Switch | False | Display previous update run history |
| `-WebhookUrl` | String | (none) | Webhook URL for completion notification |
| `-HistoryCount` | Int | 10 | Number of history entries to show |
| `-MaxRetries` | Int | 3 | Maximum retry attempts for failed operations |
| `-MaxUpdatePasses` | Int | 3 | Maximum Windows Update passes |
| `-MinDiskSpaceGB` | Int | 10 | Minimum free disk space required (GB) |
| `-LogPath` | String | C:\ProgramData\SystemUpdatePro\Logs | Log directory |
| `-LogRetentionDays` | Int | 30 | Days to keep old logs |
| `-Reboot` | Switch | False | Allow automatic reboot if required |
| `-Force` | Switch | False | Continue despite warnings |

`-CleanupAfter` never uses `/ResetBase`. The separate `-ResetComponentBase` switch is intentionally high risk and also requests cleanup, so it does not need to be combined with `-CleanupAfter`. Use `-DryRun -ResetComponentBase` to preview the exact DISM command and rollback impact. Temporary Disk Cleanup `StateFlags0100` registry values are restored after each run, and restoration or `cleanmgr` failures are reported as partial cleanup.

---

## Exit Codes

| Code | Description |
|------|-------------|
| 0 | Success, no reboot needed |
| 1 | Success, reboot required |
| 2 | Partial success (some updates failed) |
| 3 | Critical failure |
| 4 | Insufficient disk space |
| 5 | Pending reboot blocked execution |
| 6 | Already running (lock file exists) |
| 7 | Battery power (BIOS update blocked) |

---

## Event Log Integration

SystemUpdatePro writes to the Windows Application event log under source **"SystemUpdatePro"**:

| Event ID | Meaning |
|----------|---------|
| 1000 | Success, no reboot needed |
| 1001 | Success, reboot required |
| 1002 | Partial success |
| 1003 | Critical failure |
| 1004 | Insufficient disk space |
| 1005 | Pending reboot blocked |
| 1006 | Already running |
| 1007 | Battery power blocked |

### Query Events via PowerShell
```powershell
Get-EventLog -LogName Application -Source "SystemUpdatePro" -Newest 10
```

---

## File Locations

| Path | Purpose |
|------|---------|
| `C:\ProgramData\SystemUpdatePro\Logs\` | Log files and HTML reports |
| `C:\ProgramData\SystemUpdatePro\update.lock` | Lock file (prevents concurrent runs) |
| `C:\ProgramData\SystemUpdatePro\state.json` | Protected, versioned post-reboot continuation state |
| `C:\ProgramData\SystemUpdatePro\update_history.json` | Update history log (last 100 runs) |
| `C:\ProgramData\SystemUpdatePro\DriverBackups\` | Driver backup snapshots (last 3 kept) |
| `C:\ProgramData\SystemUpdatePro\HPIA\` | HP Image Assistant installation |

Continuation state is atomically replaced and restricted to SYSTEM, Administrators, and the creating identity. It preserves the run ID, effective parameters, result history, attempt count, and next stage cursor; task command lines contain only the script path. Invalid or broadly writable state is moved to `state.corrupt.<timestamp>.<id>.json` instead of being executed. A continuation can resume at most three times, and terminal success or failure removes its one-shot task and active state.

---

## HTML Reports

After each run, SystemUpdatePro generates a responsive, self-contained operations report with:
- A decisive run-status summary and at-a-glance update metrics
- OEM, Windows Update, and Winget channel breakdowns
- A compact device inventory profile
- Dedicated exceptions and follow-up guidance
- Audit-friendly run metadata and log location
- Responsive layouts for desktop and mobile plus print-optimized styles
- HTML-encoded machine and update data for safe rendering

Reports are saved to the log directory and automatically open in your browser (unless running as SYSTEM or non-interactively).

---

## Webhook Payload

When using `-WebhookUrl`, the following JSON payload is sent:

```json
{
  "schema_version": 1,
  "run_id": "c59c67f1-2f28-45a2-b8de-14872cc4973e",
  "started_at": "2026-07-29T19:00:00.0000000-04:00",
  "completed_at": "2026-07-29T19:03:00.0000000-04:00",
  "hostname": "PCNAME",
  "status": "success|partial|failed",
  "oem_updates": 3,
  "windows_updates": 5,
  "winget_updates": 12,
  "total_installed": 20,
  "total_available": 20,
  "total_failed": 0,
  "reboot_required": true,
  "exit_code": 1,
  "errors": [],
  "warnings": [],
  "runtime_seconds": 180,
  "stages": [
    {
      "Name": "WindowsUpdate",
      "Provider": "Windows Update",
      "Status": "Succeeded",
      "Attempted": 5,
      "Available": 5,
      "Installed": 5,
      "Failed": 0,
      "Skipped": 0,
      "ProviderExitCode": 2,
      "HResult": 0,
      "RebootRequired": true,
      "DurationSeconds": 94,
      "Items": [],
      "Evidence": []
    }
  ]
}
```

Slack and Teams webhooks are auto-detected by URL pattern and formatted appropriately.
The JSON history entry uses the same stage schema and additionally records whether report, Event Log, webhook, and history delivery were attempted and succeeded.

---

## RMM Deployment Examples

### NinjaOne / NinjaRMM
```powershell
# Script Variables: None required
# Run As: System
# Architecture: 64-bit

.\SystemUpdatePro.ps1 -SkipWinget
exit $LASTEXITCODE
```

### Datto RMM
```powershell
# Component Type: PowerShell
# Run As: System

$result = .\SystemUpdatePro.ps1 -SkipWinget 2>&1
Write-Host $result
exit $LASTEXITCODE
```

### ConnectWise Automate
```powershell
# Script Type: PowerShell
# Execute As: Admin

powershell.exe -ExecutionPolicy Bypass -File "C:\Temp\SystemUpdatePro.ps1" -SkipWinget
```

### PDQ Deploy
```
Steps:
1. PowerShell (Run As: Deploy User)
   Command: .\SystemUpdatePro.ps1
   Success Codes: 0,1
   Error Mode: Continue
```

---

## Scheduled Task Deployment

Deploy as a scheduled task for automatic updates:

```powershell
$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -File `"C:\Scripts\SystemUpdatePro.ps1`" -SkipWinget"
$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Saturday -At 2am
$principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -RunOnlyIfNetworkAvailable

Register-ScheduledTask -TaskName "SystemUpdatePro Weekly" -Action $action -Trigger $trigger -Principal $principal -Settings $settings
```

---

## How It Works

```
+-------------------------------------------------------------+
|                  SystemUpdatePro v4.1.0                      |
+-------------------------------------------------------------+
|                                                              |
|  1. PRE-FLIGHT CHECKS                                       |
|     +-- Admin privileges                                    |
|     +-- Lock file (prevent concurrent runs)                 |
|     +-- Internet connectivity                               |
|     +-- Disk space verification                              |
|     +-- Pending reboot detection                             |
|     +-- Battery status (for BIOS)                            |
|     +-- Metered connection warning                           |
|                                                              |
|  2. DRIVER BACKUP (if -BackupDrivers)                        |
|     +-- Export-WindowsDriver to backup directory             |
|     +-- Auto-cleanup old backups (keep 3)                    |
|                                                              |
|  3. WINDOWS UPDATE REPAIR (if -RepairWindowsUpdate)          |
|     +-- Stop WU services                                    |
|     +-- Clear SoftwareDistribution cache                    |
|     +-- Re-register 30+ DLLs                                |
|     +-- Reset Winsock                                        |
|     +-- Restart WU services                                  |
|                                                              |
|  4. OEM UPDATES (auto-detected)                              |
|     +-- Dell: Install DCU -> Apply updates                   |
|     +-- Lenovo: Install LSUClient -> Apply updates           |
|     +-- HP: Install HPIA -> Apply updates                    |
|                                                              |
|  5. WINDOWS UPDATES                                          |
|     +-- Install PSWindowsUpdate module                       |
|     +-- Multi-pass update (catches dependent updates)        |
|     +-- Fallback to WUA COM API if needed                    |
|                                                              |
|  6. WINGET UPGRADES                                          |
|     +-- Install Winget if missing (Win10 compatible)         |
|     +-- winget upgrade --all                                 |
|                                                              |
|  7. CLEANUP (if cleanup requested)                           |
|     +-- DISM component cleanup (rollback retained by default) |
|     +-- Optional explicit /ResetBase (irreversible)           |
|     +-- Disk Cleanup (update files, temp files)              |
|                                                              |
|  8. FINALIZATION                                             |
|     +-- Generate HTML report                                 |
|     +-- Write Event Log entry                                |
|     +-- Send webhook notification (if configured)            |
|     +-- Save schema-versioned history + delivery status       |
|     +-- Create continuation task (if -ContinueAfterReboot)   |
|     +-- Remove lock file                                     |
|     +-- Initiate reboot (if -Reboot and required)            |
|                                                              |
+-------------------------------------------------------------+
```

---

## Troubleshooting

### Script won't run - "Already running"
The lock file exists from a previous run. Check if another instance is running, or remove the stale lock:
```powershell
Remove-Item "C:\ProgramData\SystemUpdatePro\update.lock" -Force
```

### Dell Command Update fails with exit 3000
The Dell Client Management Service isn't running. The script will attempt auto-repair, but you can manually fix:
```powershell
Start-Service DellClientManagementService
```

### Windows Update stuck or failing
Use the repair option:
```powershell
.\SystemUpdatePro.ps1 -RepairWindowsUpdate -BypassWSUS
```

### BIOS update blocked
BIOS updates require:
- AC power (not battery)
- The `-IncludeBIOS` flag
- For Lenovo/HP with BitLocker: Manual BitLocker suspension (Dell auto-suspends)

### View detailed logs
```powershell
# Main log
Get-Content "C:\ProgramData\SystemUpdatePro\Logs\SystemUpdatePro_*.log" -Tail 100

# Full transcript
Get-Content "C:\ProgramData\SystemUpdatePro\Logs\SystemUpdatePro_Transcript_*.log"

# DCU log (Dell)
Get-Content "C:\ProgramData\SystemUpdatePro\Logs\DCU_*.log"

# View update history
.\SystemUpdatePro.ps1 -ShowHistory
```

---

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

---

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

## Acknowledgments

- [Dell Command Update](https://www.dell.com/support/kbdoc/en-us/000177325/dell-command-update) - Dell driver/BIOS management
- [LSUClient](https://github.com/jantari/LSUClient) - Lenovo System Update PowerShell module
- [HP Image Assistant](https://ftp.ext.hp.com/pub/caps-softpaq/cmit/HPIA.html) - HP driver/BIOS management
- [PSWindowsUpdate](https://www.powershellgallery.com/packages/PSWindowsUpdate) - Windows Update PowerShell module
