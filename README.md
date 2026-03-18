# SystemUpdatePro

<p align="center">
  <img src="https://img.shields.io/badge/PowerShell-5.1+-blue?logo=powershell&logoColor=white" alt="PowerShell 5.1+">
  <img src="https://img.shields.io/badge/Windows-10%20|%2011%20|%20Server-0078D6?logo=windows&logoColor=white" alt="Windows">
  <img src="https://img.shields.io/badge/License-MIT-green" alt="License">
  <img src="https://img.shields.io/badge/Version-4.0.0-orange" alt="Version">
</p>

**Enterprise-grade, bulletproof system update utility for MSPs and IT professionals.**

SystemUpdatePro is a fully automated, self-healing PowerShell script that handles OEM driver/BIOS updates (Dell, Lenovo, HP), Windows Updates, and application updates via Winget—all without user interaction.

---

## Features

### Multi-OEM Support
| Manufacturer | Tool Used | Auto-Install |
|--------------|-----------|--------------|
| Dell / Alienware | Dell Command Update CLI | ✅ via Winget |
| Lenovo | LSUClient PowerShell Module | ✅ via PSGallery |
| HP | HP Image Assistant | ✅ Auto-download |
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

### Enterprise Integration
- **Event Log**: Writes to Windows Application log for RMM/SIEM visibility
- **Exit Codes**: Granular exit codes for automation pipelines
- **WSUS Bypass**: Option to bypass WSUS and connect directly to Microsoft
- **Post-Reboot Continuation**: Scheduled task to resume updates after reboot
- **Log Rotation**: Automatic cleanup of old log files

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
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/YOUR_USERNAME/SystemUpdatePro/main/SystemUpdatePro.ps1" -OutFile "SystemUpdatePro.ps1"

# Run it
.\SystemUpdatePro.ps1
```

### Option 2: Clone Repository
```powershell
git clone https://github.com/YOUR_USERNAME/SystemUpdatePro.git
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

### Advanced Usage
```powershell
# Full provisioning workflow with post-reboot continuation
.\SystemUpdatePro.ps1 -IncludeBIOS -Reboot -ContinueAfterReboot -CleanupAfter

# Repair broken Windows Update then run updates
.\SystemUpdatePro.ps1 -RepairWindowsUpdate -BypassWSUS

# Force run despite warnings (low disk, pending reboot, battery)
.\SystemUpdatePro.ps1 -Force -IncludeBIOS

# Custom configuration
.\SystemUpdatePro.ps1 -MaxRetries 5 -MaxUpdatePasses 5 -MinDiskSpaceGB 20 -LogRetentionDays 60
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
| `-CleanupAfter` | Switch | False | Run DISM component cleanup after updates |
| `-ContinueAfterReboot` | Switch | False | Create scheduled task to continue after reboot |
| `-MaxRetries` | Int | 3 | Maximum retry attempts for failed operations |
| `-MaxUpdatePasses` | Int | 3 | Maximum Windows Update passes |
| `-MinDiskSpaceGB` | Int | 10 | Minimum free disk space required (GB) |
| `-LogPath` | String | C:\ProgramData\SystemUpdatePro\Logs | Log directory |
| `-LogRetentionDays` | Int | 30 | Days to keep old logs |
| `-Reboot` | Switch | False | Allow automatic reboot if required |
| `-Force` | Switch | False | Continue despite warnings |

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
| `C:\ProgramData\SystemUpdatePro\Logs\` | Log files |
| `C:\ProgramData\SystemUpdatePro\update.lock` | Lock file (prevents concurrent runs) |
| `C:\ProgramData\SystemUpdatePro\state.json` | State file (for post-reboot continuation) |
| `C:\ProgramData\SystemUpdatePro\HPIA\` | HP Image Assistant installation |

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
┌─────────────────────────────────────────────────────────────┐
│                    SystemUpdatePro v4.0                      │
├─────────────────────────────────────────────────────────────┤
│                                                              │
│  1. PRE-FLIGHT CHECKS                                        │
│     ├── Admin privileges                                     │
│     ├── Lock file (prevent concurrent runs)                  │
│     ├── Internet connectivity                                │
│     ├── Disk space verification                              │
│     ├── Pending reboot detection                             │
│     ├── Battery status (for BIOS)                            │
│     └── Metered connection warning                           │
│                                                              │
│  2. WINDOWS UPDATE REPAIR (if -RepairWindowsUpdate)          │
│     ├── Stop WU services                                     │
│     ├── Clear SoftwareDistribution cache                     │
│     ├── Re-register 30+ DLLs                                 │
│     ├── Reset Winsock                                        │
│     └── Restart WU services                                  │
│                                                              │
│  3. OEM UPDATES (auto-detected)                              │
│     ├── Dell: Install DCU → Apply updates                    │
│     ├── Lenovo: Install LSUClient → Apply updates            │
│     └── HP: Install HPIA → Apply updates                     │
│                                                              │
│  4. WINDOWS UPDATES                                          │
│     ├── Install PSWindowsUpdate module                       │
│     ├── Multi-pass update (catches dependent updates)        │
│     └── Fallback to WUA COM API if needed                    │
│                                                              │
│  5. WINGET UPGRADES                                          │
│     ├── Install Winget if missing (Win10 compatible)         │
│     └── winget upgrade --all                                 │
│                                                              │
│  6. CLEANUP (if -CleanupAfter)                               │
│     ├── DISM /StartComponentCleanup /ResetBase               │
│     └── Disk Cleanup (update files, temp files)              │
│                                                              │
│  7. FINALIZATION                                             │
│     ├── Write Event Log entry                                │
│     ├── Create continuation task (if -ContinueAfterReboot)   │
│     ├── Remove lock file                                     │
│     └── Initiate reboot (if -Reboot and required)            │
│                                                              │
└─────────────────────────────────────────────────────────────┘
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

---

<p align="center">
  Made with ☕ for MSPs who are tired of broken update tools
</p>
