<#
.SYNOPSIS
    SystemUpdatePro v4.0 - Enterprise Multi-OEM System Update Utility
.DESCRIPTION
    Bulletproof MSP-grade unattended update tool with self-healing capabilities.
    
    FEATURES:
    - Multi-OEM Support: Dell, Lenovo, HP (auto-detects manufacturer)
    - Windows Update with automatic service repair
    - Winget upgrade all (auto-installs winget on Windows 10)
    - Self-healing: repairs corrupted Windows Update components
    - BitLocker-aware BIOS handling
    - Battery safety: blocks BIOS updates on battery power
    - Disk space verification before updates
    - Post-reboot continuation via scheduled task
    - Event Log integration for RMM visibility
    - Concurrent execution prevention (lock file)
    - Automatic log rotation
    - WSUS bypass option for direct Microsoft updates
    - Comprehensive retry logic with exponential backoff
    
    OEM TOOLS:
    - Dell/Alienware: Dell Command Update CLI
    - Lenovo: LSUClient PowerShell module  
    - HP: HP Image Assistant
    
.PARAMETER SkipOEM
    Skip OEM-specific driver/firmware updates
.PARAMETER SkipWindows
    Skip Windows Update
.PARAMETER SkipWinget
    Skip Winget upgrade all
.PARAMETER IncludeBIOS
    Include BIOS updates (requires AC power, handles BitLocker)
.PARAMETER BypassWSUS
    Bypass WSUS and connect directly to Microsoft Update
.PARAMETER RepairWindowsUpdate
    Force Windows Update component repair before updating
.PARAMETER CleanupAfter
    Run DISM component cleanup after updates to reclaim space
.PARAMETER ContinueAfterReboot
    Create scheduled task to continue updates after reboot
.PARAMETER MaxRetries
    Maximum retry attempts for failed operations (default: 3)
.PARAMETER MaxUpdatePasses
    Maximum Windows Update passes (default: 3)
.PARAMETER MinDiskSpaceGB
    Minimum free disk space required in GB (default: 10)
.PARAMETER LogPath
    Custom log directory (default: C:\ProgramData\SystemUpdatePro\Logs)
.PARAMETER LogRetentionDays
    Days to keep old logs (default: 30)
.PARAMETER Reboot
    Allow automatic reboot if required
.PARAMETER Force
    Continue despite warnings (pending reboot, low disk, battery)
.EXAMPLE
    .\SystemUpdatePro.ps1
    # Standard update: OEM + Windows + Winget
.EXAMPLE
    .\SystemUpdatePro.ps1 -IncludeBIOS -Reboot -ContinueAfterReboot
    # Full update with BIOS, auto-reboot, and post-reboot continuation
.EXAMPLE
    .\SystemUpdatePro.ps1 -BypassWSUS -RepairWindowsUpdate
    # Repair WU components and bypass WSUS
.EXAMPLE
    .\SystemUpdatePro.ps1 -SkipOEM -CleanupAfter
    # Windows + Winget only, cleanup after
.NOTES
    Version: 4.0.0
    Requires: Administrator, PowerShell 5.1+, Internet
    
    EXIT CODES:
        0 = Success, no reboot needed
        1 = Success, reboot required
        2 = Partial success (some failed)
        3 = Critical failure
        4 = Insufficient disk space
        5 = Pending reboot blocked execution
        6 = Already running (lock file exists)
        7 = Battery power (BIOS update blocked)
#>

#Requires -Version 5.1

[CmdletBinding()]
param(
    [switch]$SkipOEM,
    [switch]$SkipWindows,
    [switch]$SkipWinget,
    [switch]$IncludeBIOS,
    [switch]$BypassWSUS,
    [switch]$RepairWindowsUpdate,
    [switch]$CleanupAfter,
    [switch]$ContinueAfterReboot,
    [int]$MaxRetries = 3,
    [int]$MaxUpdatePasses = 3,
    [int]$MinDiskSpaceGB = 10,
    [string]$LogPath = "C:\ProgramData\SystemUpdatePro\Logs",
    [int]$LogRetentionDays = 30,
    [switch]$Reboot,
    [switch]$Force
)

# ============================================================================
# SCRIPT CONFIGURATION
# ============================================================================

$script:Version = "4.0.0"
$script:ProductName = "SystemUpdatePro"
$script:EventLogSource = "SystemUpdatePro"
$script:LockFile = "C:\ProgramData\SystemUpdatePro\update.lock"
$script:StateFile = "C:\ProgramData\SystemUpdatePro\state.json"
$script:TaskName = "SystemUpdatePro_Continue"

$script:ExitCode = 0
$script:RebootRequired = $false
$script:UpdatesInstalled = 0
$script:UpdatesFailed = 0
$script:Warnings = [System.Collections.ArrayList]::new()
$script:Errors = [System.Collections.ArrayList]::new()

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

# ============================================================================
# INITIALIZATION
# ============================================================================

# Verify admin privileges
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "[X] This script requires administrator privileges." -ForegroundColor Red
    exit 3
}

# Create directories
$script:DataPath = "C:\ProgramData\SystemUpdatePro"
@($script:DataPath, $LogPath) | ForEach-Object {
    if (-not (Test-Path $_)) {
        New-Item -ItemType Directory -Path $_ -Force | Out-Null
    }
}

# Setup logging
$script:LogFile = Join-Path $LogPath "$($script:ProductName)_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$script:TranscriptFile = Join-Path $LogPath "$($script:ProductName)_Transcript_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

# Start transcript for deep debugging
try {
    Start-Transcript -Path $script:TranscriptFile -Force | Out-Null
} catch {}

# ============================================================================
# EVENT LOG SETUP
# ============================================================================

function Initialize-EventLog {
    try {
        if (-not [System.Diagnostics.EventLog]::SourceExists($script:EventLogSource)) {
            New-EventLog -LogName "Application" -Source $script:EventLogSource -ErrorAction Stop
        }
        return $true
    } catch {
        return $false
    }
}

function Write-EventLogEntry {
    param(
        [string]$Message,
        [ValidateSet("Information", "Warning", "Error")]
        [string]$EntryType = "Information",
        [int]$EventId = 1000
    )
    
    try {
        Write-EventLog -LogName "Application" -Source $script:EventLogSource -EntryType $EntryType -EventId $EventId -Message $Message -ErrorAction SilentlyContinue
    } catch {}
}

# ============================================================================
# LOGGING FUNCTIONS
# ============================================================================

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO", "SUCCESS", "WARNING", "ERROR", "DEBUG", "HEADER", "STEP")]
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    
    try {
        Add-Content -Path $script:LogFile -Value $logEntry -ErrorAction SilentlyContinue
    } catch {}
    
    $colors = @{
        "HEADER"  = "Cyan"
        "STEP"    = "Magenta"
        "SUCCESS" = "Green"
        "WARNING" = "Yellow"
        "ERROR"   = "Red"
        "DEBUG"   = "DarkGray"
        "INFO"    = "White"
    }
    
    $prefixes = @{
        "HEADER"  = ""
        "STEP"    = "[*] "
        "SUCCESS" = "[+] "
        "WARNING" = "[!] "
        "ERROR"   = "[X] "
        "DEBUG"   = "    "
        "INFO"    = "    "
    }
    
    Write-Host "$($prefixes[$Level])$Message" -ForegroundColor $colors[$Level]
    
    # Track warnings and errors
    if ($Level -eq "WARNING") { [void]$script:Warnings.Add($Message) }
    if ($Level -eq "ERROR") { [void]$script:Errors.Add($Message) }
}

function Write-Banner {
    $banner = @"

  ================================================================
    $($script:ProductName) v$($script:Version)
    Enterprise System Update Utility
  ================================================================

"@
    Write-Host $banner -ForegroundColor Cyan
}

# ============================================================================
# LOCK FILE MANAGEMENT
# ============================================================================

function Test-LockFile {
    if (Test-Path $script:LockFile) {
        $lockContent = Get-Content $script:LockFile -ErrorAction SilentlyContinue | ConvertFrom-Json -ErrorAction SilentlyContinue
        if ($lockContent) {
            # Check if the process is still running
            $process = Get-Process -Id $lockContent.PID -ErrorAction SilentlyContinue
            if ($process -and $process.ProcessName -eq "powershell") {
                # Check if lock is stale (older than 4 hours)
                $lockTime = [DateTime]::Parse($lockContent.StartTime)
                if ((Get-Date) - $lockTime -lt [TimeSpan]::FromHours(4)) {
                    return $true  # Lock is valid
                }
            }
        }
        # Stale lock - remove it
        Remove-Item $script:LockFile -Force -ErrorAction SilentlyContinue
    }
    return $false
}

function New-LockFile {
    $lockData = @{
        PID = $PID
        StartTime = (Get-Date).ToString("o")
        Computer = $env:COMPUTERNAME
    } | ConvertTo-Json
    
    Set-Content -Path $script:LockFile -Value $lockData -Force
}

function Remove-LockFile {
    Remove-Item $script:LockFile -Force -ErrorAction SilentlyContinue
}

# ============================================================================
# STATE MANAGEMENT (for post-reboot continuation)
# ============================================================================

function Save-State {
    param([hashtable]$State)
    
    $State.LastUpdate = (Get-Date).ToString("o")
    $State | ConvertTo-Json -Depth 5 | Set-Content -Path $script:StateFile -Force
}

function Get-State {
    if (Test-Path $script:StateFile) {
        try {
            return Get-Content $script:StateFile | ConvertFrom-Json -AsHashtable
        } catch {}
    }
    return @{}
}

function Clear-State {
    Remove-Item $script:StateFile -Force -ErrorAction SilentlyContinue
}

# ============================================================================
# LOG ROTATION
# ============================================================================

function Invoke-LogRotation {
    param([int]$RetentionDays = 30)
    
    $cutoffDate = (Get-Date).AddDays(-$RetentionDays)
    
    Get-ChildItem -Path $LogPath -Filter "*.log" -ErrorAction SilentlyContinue | 
        Where-Object { $_.LastWriteTime -lt $cutoffDate } | 
        Remove-Item -Force -ErrorAction SilentlyContinue
    
    $removed = @(Get-ChildItem -Path $LogPath -Filter "*.log" -ErrorAction SilentlyContinue | 
        Where-Object { $_.LastWriteTime -lt $cutoffDate }).Count
    
    if ($removed -gt 0) {
        Write-Log "Removed $removed old log files" "DEBUG"
    }
}

# ============================================================================
# SYSTEM CHECKS
# ============================================================================

function Test-InternetConnection {
    $endpoints = @(
        "https://www.microsoft.com",
        "https://download.microsoft.com", 
        "https://www.google.com"
    )
    
    foreach ($url in $endpoints) {
        try {
            $request = [System.Net.WebRequest]::Create($url)
            $request.Timeout = 10000
            $request.Method = "HEAD"
            $response = $request.GetResponse()
            $response.Close()
            return $true
        } catch { continue }
    }
    return $false
}

function Test-DiskSpace {
    param([int]$MinGB = 10)
    
    $systemDrive = $env:SystemDrive
    $disk = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='$systemDrive'" -ErrorAction SilentlyContinue
    
    if ($disk) {
        $freeGB = [math]::Round($disk.FreeSpace / 1GB, 2)
        return @{
            Sufficient = ($freeGB -ge $MinGB)
            FreeGB = $freeGB
            RequiredGB = $MinGB
        }
    }
    
    return @{ Sufficient = $true; FreeGB = 0; RequiredGB = $MinGB }
}

function Test-BatteryPower {
    try {
        $battery = Get-CimInstance -ClassName Win32_Battery -ErrorAction SilentlyContinue
        if ($battery) {
            # BatteryStatus: 1=Discharging, 2=AC Power
            $onBattery = $battery.BatteryStatus -eq 1
            $chargePercent = $battery.EstimatedChargeRemaining
            
            return @{
                HasBattery = $true
                OnBattery = $onBattery
                OnACPower = -not $onBattery
                ChargePercent = $chargePercent
            }
        }
    } catch {}
    
    return @{
        HasBattery = $false
        OnBattery = $false
        OnACPower = $true
        ChargePercent = 100
    }
}

function Test-PendingReboot {
    $reasons = [System.Collections.ArrayList]::new()
    
    # Component Based Servicing
    if (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending") {
        [void]$reasons.Add("Component Based Servicing")
    }
    
    # Windows Update
    if (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired") {
        [void]$reasons.Add("Windows Update")
    }
    
    # Pending File Rename
    try {
        $pfro = Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" -Name PendingFileRenameOperations -ErrorAction SilentlyContinue
        if ($pfro.PendingFileRenameOperations) {
            [void]$reasons.Add("Pending File Rename")
        }
    } catch {}
    
    # Computer Rename
    try {
        $active = (Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ActiveComputerName" -ErrorAction SilentlyContinue).ComputerName
        $pending = (Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName" -ErrorAction SilentlyContinue).ComputerName
        if ($active -ne $pending) {
            [void]$reasons.Add("Computer Rename")
        }
    } catch {}
    
    # SCCM Client
    try {
        $ccm = Invoke-CimMethod -Namespace "ROOT\ccm\ClientSDK" -ClassName "CCM_ClientUtilities" -MethodName "DetermineIfRebootPending" -ErrorAction SilentlyContinue
        if ($ccm -and ($ccm.RebootPending -or $ccm.IsHardRebootPending)) {
            [void]$reasons.Add("SCCM Client")
        }
    } catch {}
    
    return @{
        Pending = ($reasons.Count -gt 0)
        Reasons = $reasons
    }
}

function Test-BitLockerEnabled {
    try {
        $bl = Get-BitLockerVolume -MountPoint $env:SystemDrive -ErrorAction Stop
        return @{
            Enabled = ($bl.ProtectionStatus -eq "On")
            Status = $bl.ProtectionStatus
            Method = $bl.EncryptionMethod
        }
    } catch {
        return @{ Enabled = $false; Status = "Unknown"; Method = "N/A" }
    }
}

function Test-MeteredConnection {
    try {
        $cost = [Windows.Networking.Connectivity.NetworkInformation, Windows, ContentType=WindowsRuntime]::GetInternetConnectionProfile().GetConnectionCost()
        return ($cost.NetworkCostType -ne [Windows.Networking.Connectivity.NetworkCostType]::Unrestricted)
    } catch {
        return $false
    }
}

function Get-SystemInfo {
    $cs = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction SilentlyContinue
    $bios = Get-CimInstance -ClassName Win32_BIOS -ErrorAction SilentlyContinue
    $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction SilentlyContinue
    $cpu = Get-CimInstance -ClassName Win32_Processor -ErrorAction SilentlyContinue | Select-Object -First 1
    
    return @{
        Manufacturer = $cs.Manufacturer
        Model = $cs.Model
        SerialNumber = $bios.SerialNumber
        BIOSVersion = $bios.SMBIOSBIOSVersion
        BIOSDate = $bios.ReleaseDate
        OSName = $os.Caption
        OSVersion = $os.Version
        OSBuild = $os.BuildNumber
        TotalRAM = [math]::Round($cs.TotalPhysicalMemory / 1GB, 1)
        Processor = $cpu.Name
    }
}

# ============================================================================
# RETRY LOGIC
# ============================================================================

function Invoke-WithRetry {
    param(
        [scriptblock]$ScriptBlock,
        [string]$OperationName,
        [int]$MaxAttempts = $MaxRetries,
        [int]$InitialDelaySeconds = 5
    )
    
    $attempt = 0
    $delay = $InitialDelaySeconds
    $lastError = $null
    
    while ($attempt -lt $MaxAttempts) {
        $attempt++
        try {
            return & $ScriptBlock
        } catch {
            $lastError = $_
            if ($attempt -lt $MaxAttempts) {
                Write-Log "$OperationName failed (attempt $attempt/$MaxAttempts): $($_.Exception.Message)" "WARNING"
                Write-Log "Retrying in $delay seconds..." "DEBUG"
                Start-Sleep -Seconds $delay
                $delay = [math]::Min($delay * 2, 120)  # Max 2 minutes
            }
        }
    }
    
    throw "Operation '$OperationName' failed after $MaxAttempts attempts. Last error: $($lastError.Exception.Message)"
}

# ============================================================================
# SERVICE MANAGEMENT
# ============================================================================

function Set-ServiceState {
    param(
        [string]$ServiceName,
        [string]$DesiredState = "Running",
        [string]$StartupType = "Automatic",
        [int]$TimeoutSeconds = 60
    )
    
    try {
        $service = Get-Service -Name $ServiceName -ErrorAction Stop
        
        Set-Service -Name $ServiceName -StartupType $StartupType -ErrorAction SilentlyContinue
        
        if ($DesiredState -eq "Running" -and $service.Status -ne "Running") {
            Start-Service -Name $ServiceName -ErrorAction Stop
            
            $timeout = [DateTime]::Now.AddSeconds($TimeoutSeconds)
            do {
                Start-Sleep -Milliseconds 500
                $service = Get-Service -Name $ServiceName
            } while ($service.Status -ne "Running" -and [DateTime]::Now -lt $timeout)
        }
        
        return ($service.Status -eq $DesiredState)
    } catch {
        return $false
    }
}

function Repair-WindowsUpdateServices {
    Write-Log "Repairing Windows Update services..." "STEP"
    
    $services = @(
        @{ Name = "wuauserv"; DisplayName = "Windows Update" },
        @{ Name = "bits"; DisplayName = "Background Intelligent Transfer" },
        @{ Name = "cryptsvc"; DisplayName = "Cryptographic Services" },
        @{ Name = "msiserver"; DisplayName = "Windows Installer" },
        @{ Name = "TrustedInstaller"; DisplayName = "Windows Modules Installer" }
    )
    
    # Stop services
    Write-Log "Stopping Windows Update services..." "DEBUG"
    foreach ($svc in $services) {
        Stop-Service -Name $svc.Name -Force -ErrorAction SilentlyContinue
    }
    
    Start-Sleep -Seconds 3
    
    # Clear update cache
    Write-Log "Clearing Windows Update cache..." "DEBUG"
    $cachePaths = @(
        "$env:SystemRoot\SoftwareDistribution\Download\*",
        "$env:SystemRoot\System32\catroot2\*"
    )
    
    foreach ($path in $cachePaths) {
        Remove-Item -Path $path -Recurse -Force -ErrorAction SilentlyContinue
    }
    
    # Re-register DLLs
    Write-Log "Re-registering Windows Update DLLs..." "DEBUG"
    $dlls = @(
        "atl.dll", "urlmon.dll", "mshtml.dll", "shdocvw.dll", "browseui.dll",
        "jscript.dll", "vbscript.dll", "scrrun.dll", "msxml.dll", "msxml3.dll",
        "msxml6.dll", "actxprxy.dll", "softpub.dll", "wintrust.dll", "dssenh.dll",
        "rsaenh.dll", "gpkcsp.dll", "sccbase.dll", "slbcsp.dll", "cryptdlg.dll",
        "oleaut32.dll", "ole32.dll", "shell32.dll", "initpki.dll", "wuapi.dll",
        "wuaueng.dll", "wuaueng1.dll", "wucltui.dll", "wups.dll", "wups2.dll",
        "wuweb.dll", "qmgr.dll", "qmgrprxy.dll", "wucltux.dll", "muweb.dll", "wuwebv.dll"
    )
    
    foreach ($dll in $dlls) {
        $dllPath = Join-Path $env:SystemRoot "System32\$dll"
        if (Test-Path $dllPath) {
            & regsvr32.exe /s $dllPath 2>$null
        }
    }
    
    # Reset Winsock
    Write-Log "Resetting network components..." "DEBUG"
    & netsh winsock reset 2>$null
    & netsh winhttp reset proxy 2>$null
    
    # Start services
    Write-Log "Starting Windows Update services..." "DEBUG"
    $failed = @()
    foreach ($svc in $services) {
        if (-not (Set-ServiceState -ServiceName $svc.Name -DesiredState "Running" -StartupType "Automatic")) {
            $failed += $svc.DisplayName
        }
    }
    
    if ($failed.Count -eq 0) {
        Write-Log "Windows Update services repaired successfully" "SUCCESS"
        return $true
    } else {
        Write-Log "Failed to start: $($failed -join ', ')" "WARNING"
        return $false
    }
}

function Set-WSUSBypass {
    param([switch]$Enable)
    
    $wuPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"
    $auPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"
    
    if ($Enable) {
        Write-Log "Configuring WSUS bypass (direct to Microsoft)..." "DEBUG"
        
        # Backup current settings
        $script:WSUSBackup = @{}
        
        if (Test-Path $wuPath) {
            $script:WSUSBackup.WUServer = (Get-ItemProperty $wuPath -Name WUServer -ErrorAction SilentlyContinue).WUServer
            $script:WSUSBackup.WUStatusServer = (Get-ItemProperty $wuPath -Name WUStatusServer -ErrorAction SilentlyContinue).WUStatusServer
        }
        
        if (Test-Path $auPath) {
            $script:WSUSBackup.UseWUServer = (Get-ItemProperty $auPath -Name UseWUServer -ErrorAction SilentlyContinue).UseWUServer
        }
        
        # Disable WSUS
        if (Test-Path $auPath) {
            Set-ItemProperty -Path $auPath -Name UseWUServer -Value 0 -Type DWord -ErrorAction SilentlyContinue
        }
        
        # Restart Windows Update service
        Restart-Service wuauserv -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 3
        
        Write-Log "WSUS bypass enabled" "SUCCESS"
    } else {
        # Restore original settings
        if ($script:WSUSBackup -and $script:WSUSBackup.UseWUServer) {
            Write-Log "Restoring WSUS settings..." "DEBUG"
            
            if (Test-Path $auPath) {
                Set-ItemProperty -Path $auPath -Name UseWUServer -Value $script:WSUSBackup.UseWUServer -Type DWord -ErrorAction SilentlyContinue
            }
            
            Restart-Service wuauserv -Force -ErrorAction SilentlyContinue
        }
    }
}

# ============================================================================
# POST-REBOOT CONTINUATION
# ============================================================================

function Register-ContinuationTask {
    Write-Log "Creating post-reboot continuation task..." "DEBUG"
    
    try {
        # Remove existing task if present
        Unregister-ScheduledTask -TaskName $script:TaskName -Confirm:$false -ErrorAction SilentlyContinue
        
        # Build the command to resume
        $scriptPath = $MyInvocation.PSCommandPath
        if (-not $scriptPath) { $scriptPath = $PSCommandPath }
        
        $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""
        if ($SkipOEM) { $arguments += " -SkipOEM" }
        if ($SkipWinget) { $arguments += " -SkipWinget" }
        if ($IncludeBIOS) { $arguments += " -IncludeBIOS" }
        if ($CleanupAfter) { $arguments += " -CleanupAfter" }
        
        $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument $arguments
        $trigger = New-ScheduledTaskTrigger -AtStartup
        $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
        $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable
        
        Register-ScheduledTask -TaskName $script:TaskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force | Out-Null
        
        # Save state
        Save-State @{
            Phase = "PostReboot"
            SkipOEM = $SkipOEM.IsPresent
            SkipWindows = $false
            SkipWinget = $SkipWinget.IsPresent
            IncludeBIOS = $IncludeBIOS.IsPresent
            UpdatesInstalled = $script:UpdatesInstalled
        }
        
        Write-Log "Continuation task registered: $($script:TaskName)" "SUCCESS"
        return $true
    } catch {
        Write-Log "Failed to create continuation task: $($_.Exception.Message)" "WARNING"
        return $false
    }
}

function Unregister-ContinuationTask {
    Unregister-ScheduledTask -TaskName $script:TaskName -Confirm:$false -ErrorAction SilentlyContinue
    Clear-State
}

# ============================================================================
# CLEANUP
# ============================================================================

function Invoke-ComponentCleanup {
    Write-Log "Running DISM component cleanup..." "STEP"
    
    try {
        $dismArgs = "/Online /Cleanup-Image /StartComponentCleanup /ResetBase"
        $process = Start-Process -FilePath "dism.exe" -ArgumentList $dismArgs -Wait -NoNewWindow -PassThru
        
        if ($process.ExitCode -eq 0) {
            Write-Log "Component cleanup completed" "SUCCESS"
            
            # Also run disk cleanup for update files
            Write-Log "Cleaning Windows Update files..." "DEBUG"
            
            # Set StateFlags for cleanup
            $volCachePath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches"
            $cleanupItems = @(
                "Update Cleanup",
                "Windows Update Cleanup", 
                "Temporary Files",
                "System error memory dump files",
                "Delivery Optimization Files"
            )
            
            foreach ($item in $cleanupItems) {
                $itemPath = Join-Path $volCachePath $item
                if (Test-Path $itemPath) {
                    Set-ItemProperty -Path $itemPath -Name StateFlags0100 -Value 2 -Type DWord -ErrorAction SilentlyContinue
                }
            }
            
            # Run cleanmgr
            Start-Process -FilePath "cleanmgr.exe" -ArgumentList "/sagerun:100" -Wait -NoNewWindow -ErrorAction SilentlyContinue
            
            return $true
        } else {
            Write-Log "DISM cleanup returned code: $($process.ExitCode)" "WARNING"
            return $false
        }
    } catch {
        Write-Log "Cleanup error: $($_.Exception.Message)" "WARNING"
        return $false
    }
}

# ============================================================================
# WINGET MANAGEMENT
# ============================================================================

function Test-WingetInstalled {
    try {
        $winget = Get-Command winget -ErrorAction Stop
        $null = & winget --version 2>&1
        return $true
    } catch {
        return $false
    }
}

function Install-Winget {
    Write-Log "Installing Winget..." "STEP"
    
    # Method 1: App Installer registration
    try {
        Add-AppxPackage -RegisterByFamilyName -MainPackage Microsoft.DesktopAppInstaller_8wekyb3d8bbwe -ErrorAction Stop
        Start-Sleep -Seconds 3
        if (Test-WingetInstalled) {
            Write-Log "Winget installed via App Installer" "SUCCESS"
            return $true
        }
    } catch {}
    
    # Method 2: Full installation with dependencies
    Write-Log "Installing Winget with dependencies..." "DEBUG"
    
    $tempDir = Join-Path $env:TEMP "WingetInstall_$(Get-Random)"
    New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
    
    try {
        # VCLibs
        try {
            $vcLibsUrl = "https://aka.ms/Microsoft.VCLibs.x64.14.00.Desktop.appx"
            $vcLibsPath = Join-Path $tempDir "VCLibs.appx"
            Invoke-WebRequest -Uri $vcLibsUrl -OutFile $vcLibsPath -UseBasicParsing -ErrorAction Stop
            Add-AppxPackage -Path $vcLibsPath -ErrorAction SilentlyContinue
        } catch {}
        
        # UI.Xaml
        try {
            $xamlUrl = "https://www.nuget.org/api/v2/package/Microsoft.UI.Xaml/2.8.6"
            $xamlZip = Join-Path $tempDir "xaml.zip"
            Invoke-WebRequest -Uri $xamlUrl -OutFile $xamlZip -UseBasicParsing -ErrorAction Stop
            Expand-Archive -Path $xamlZip -DestinationPath (Join-Path $tempDir "xaml") -Force
            $xamlAppx = Get-ChildItem -Path (Join-Path $tempDir "xaml") -Filter "*x64*.appx" -Recurse | Select-Object -First 1
            if ($xamlAppx) {
                Add-AppxPackage -Path $xamlAppx.FullName -ErrorAction SilentlyContinue
            }
        } catch {}
        
        # Winget bundle
        $release = Invoke-RestMethod -Uri "https://api.github.com/repos/microsoft/winget-cli/releases/latest" -ErrorAction Stop
        $bundleUrl = ($release.assets | Where-Object { $_.name -match "\.msixbundle$" -and $_.name -notmatch "License" }).browser_download_url
        $licenseUrl = ($release.assets | Where-Object { $_.name -match "License.*\.xml$" }).browser_download_url
        
        $bundlePath = Join-Path $tempDir "winget.msixbundle"
        Invoke-WebRequest -Uri $bundleUrl -OutFile $bundlePath -UseBasicParsing -ErrorAction Stop
        
        $licensePath = $null
        if ($licenseUrl) {
            $licensePath = Join-Path $tempDir "License.xml"
            try {
                Invoke-WebRequest -Uri $licenseUrl -OutFile $licensePath -UseBasicParsing
            } catch { $licensePath = $null }
        }
        
        if ($licensePath -and (Test-Path $licensePath)) {
            Add-AppxProvisionedPackage -Online -PackagePath $bundlePath -LicensePath $licensePath -ErrorAction Stop | Out-Null
        } else {
            Add-AppxPackage -Path $bundlePath -ErrorAction Stop
        }
        
        Start-Sleep -Seconds 3
        
        if (Test-WingetInstalled) {
            Write-Log "Winget installed successfully" "SUCCESS"
            return $true
        }
    } catch {
        Write-Log "Winget installation error: $($_.Exception.Message)" "WARNING"
    } finally {
        Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
    
    Write-Log "Failed to install Winget" "WARNING"
    return $false
}

function Invoke-WingetUpgradeAll {
    $result = @{ Success = $false; UpdateCount = 0; Message = "" }
    
    Write-Log "========== WINGET UPGRADE ALL ==========" "HEADER"
    
    if (-not (Test-WingetInstalled)) {
        if (-not (Install-Winget)) {
            $result.Message = "Winget not available"
            Write-Log $result.Message "WARNING"
            return $result
        }
    }
    
    try {
        # Update sources
        & winget source update --disable-interactivity 2>&1 | Out-Null
        
        Write-Log "Running winget upgrade --all..." "INFO"
        
        $args = @(
            "upgrade", "--all", "--silent",
            "--accept-package-agreements", "--accept-source-agreements",
            "--disable-interactivity", "--include-unknown"
        )
        
        $process = Start-Process -FilePath "winget" -ArgumentList $args -Wait -NoNewWindow -PassThru
        
        $result.Success = ($process.ExitCode -in @(0, -1978335189))  # 0 or "no updates"
        $result.Message = "Winget upgrade completed (exit: $($process.ExitCode))"
        
        Write-Log $result.Message $(if ($result.Success) { "SUCCESS" } else { "WARNING" })
        
    } catch {
        $result.Message = "Winget error: $($_.Exception.Message)"
        Write-Log $result.Message "ERROR"
    }
    
    return $result
}

# ============================================================================
# POWERSHELL MODULE MANAGEMENT
# ============================================================================

function Install-PSModuleWithRetry {
    param(
        [string]$ModuleName,
        [switch]$AcceptLicense
    )
    
    $existing = Get-Module -ListAvailable -Name $ModuleName -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Log "$ModuleName v$($existing.Version) available" "DEBUG"
        return $true
    }
    
    Write-Log "Installing $ModuleName module..." "INFO"
    
    return Invoke-WithRetry -OperationName "Install $ModuleName" -ScriptBlock {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        
        $nuget = Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue
        if (-not $nuget -or $nuget.Version -lt [Version]"2.8.5.201") {
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope AllUsers | Out-Null
        }
        
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue
        
        $params = @{
            Name = $ModuleName
            Force = $true
            AllowClobber = $true
            SkipPublisherCheck = $true
            Scope = "AllUsers"
        }
        
        if ($AcceptLicense) { $params.AcceptLicense = $true }
        
        Install-Module @params -ErrorAction Stop
        
        $installed = Get-Module -ListAvailable -Name $ModuleName
        if (-not $installed) { throw "Verification failed" }
        
        Write-Log "$ModuleName installed" "SUCCESS"
        return $true
    }
}

# ============================================================================
# DELL COMMAND UPDATE
# ============================================================================

function Get-DCUPath {
    @(
        "${env:ProgramFiles}\Dell\CommandUpdate\dcu-cli.exe",
        "${env:ProgramFiles(x86)}\Dell\CommandUpdate\dcu-cli.exe"
    ) | Where-Object { Test-Path $_ } | Select-Object -First 1
}

function Install-DellCommandUpdate {
    Write-Log "Installing Dell Command Update..." "INFO"
    
    if (-not (Test-WingetInstalled)) {
        if (-not (Install-Winget)) { return $false }
    }
    
    return Invoke-WithRetry -OperationName "Install DCU" -ScriptBlock {
        & winget source update --disable-interactivity 2>&1 | Out-Null
        
        $args = @("install", "--id", "Dell.CommandUpdate", "--source", "winget",
                  "--accept-package-agreements", "--accept-source-agreements", "--silent")
        
        Start-Process -FilePath "winget" -ArgumentList $args -Wait -NoNewWindow
        Start-Sleep -Seconds 5
        
        if (-not (Get-DCUPath)) { throw "DCU not found after install" }
        
        Write-Log "Dell Command Update installed" "SUCCESS"
        return $true
    }
}

function Repair-DellServices {
    $serviceName = "DellClientManagementService"
    $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
    
    if (-not $service) {
        Write-Log "Dell service not found - reinstalling DCU" "WARNING"
        & winget uninstall --id Dell.CommandUpdate --silent 2>&1 | Out-Null
        return (Install-DellCommandUpdate)
    }
    
    if ($service.Status -ne "Running") {
        Set-Service -Name $serviceName -StartupType Automatic -ErrorAction SilentlyContinue
        Start-Service -Name $serviceName -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 3
        $service = Get-Service -Name $serviceName
    }
    
    return ($service.Status -eq "Running")
}

function Invoke-DellUpdate {
    param([switch]$IncludeBIOS)
    
    $result = @{ Success = $false; RebootRequired = $false; UpdateCount = 0; Message = "" }
    
    Write-Log "========== DELL COMMAND UPDATE ==========" "HEADER"
    
    $sysInfo = Get-SystemInfo
    Write-Log "Service Tag: $($sysInfo.SerialNumber)" "INFO"
    
    $dcuPath = Get-DCUPath
    if (-not $dcuPath) {
        if (-not (Install-DellCommandUpdate)) {
            $result.Message = "Failed to install DCU"
            Write-Log $result.Message "ERROR"
            return $result
        }
        $dcuPath = Get-DCUPath
    }
    
    if (-not (Repair-DellServices)) {
        Write-Log "Dell service issues - proceeding anyway" "WARNING"
    }
    
    # Disable bloat services
    Get-Service -ErrorAction SilentlyContinue | Where-Object {
        ($_.DisplayName -like "*Dell*" -or $_.Name -like "*DDV*" -or $_.Name -like "*SupportAssist*") -and
        $_.Name -ne "DellClientManagementService"
    } | ForEach-Object {
        Stop-Service -Name $_.Name -Force -ErrorAction SilentlyContinue
        Set-Service -Name $_.Name -StartupType Disabled -ErrorAction SilentlyContinue
    }
    
    # Build arguments
    $bitlocker = Test-BitLockerEnabled
    $dcuArgs = @("/applyUpdates", "-silent", "-updateSeverity=security,critical,recommended", "-reboot=disable")
    
    if (-not $IncludeBIOS -or ($bitlocker.Enabled -and -not $IncludeBIOS)) {
        $dcuArgs += "-updateType=driver,firmware,application"
    } else {
        $dcuArgs += "-autoSuspendBitLocker=enable"
    }
    
    $dcuLog = Join-Path $LogPath "DCU_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    $dcuArgs += "-outputLog=`"$dcuLog`""
    
    Write-Log "Applying Dell updates..." "INFO"
    
    $attempts = 0
    while ($attempts -lt $MaxRetries) {
        $attempts++
        
        $process = Start-Process -FilePath $dcuPath -ArgumentList ($dcuArgs -join " ") -Wait -NoNewWindow -PassThru
        
        switch ($process.ExitCode) {
            0   { $result.Success = $true; $result.Message = "Updates applied"; break }
            1   { $result.Success = $true; $result.RebootRequired = $true; $result.Message = "Updates applied - reboot required"; break }
            500 { $result.Success = $true; $result.Message = "No updates available"; break }
            3000 {
                if ($attempts -lt $MaxRetries -and (Repair-DellServices)) { continue }
                $result.Message = "Dell service not running"
                break
            }
            3003 {
                if ($attempts -lt $MaxRetries) { Start-Sleep -Seconds 30; continue }
                $result.Message = "DCU service busy"
                break
            }
            default { $result.Message = "DCU exit code: $($process.ExitCode)"; break }
        }
        break
    }
    
    Write-Log $result.Message $(if ($result.Success) { "SUCCESS" } else { "WARNING" })
    return $result
}

# ============================================================================
# LENOVO LSUClient
# ============================================================================

function Invoke-LenovoUpdate {
    param([switch]$IncludeBIOS)
    
    $result = @{ Success = $false; RebootRequired = $false; UpdateCount = 0; Message = "" }
    
    Write-Log "========== LENOVO SYSTEM UPDATE ==========" "HEADER"
    
    $sysInfo = Get-SystemInfo
    Write-Log "Serial: $($sysInfo.SerialNumber)" "INFO"
    
    if (-not (Install-PSModuleWithRetry -ModuleName "LSUClient")) {
        $result.Message = "Failed to install LSUClient"
        Write-Log $result.Message "ERROR"
        return $result
    }
    
    try {
        Import-Module LSUClient -Force -ErrorAction Stop
        
        Write-Log "Scanning for updates..." "INFO"
        $updates = Get-LSUpdate -ErrorAction Stop | Where-Object { $_.Installer.Unattended -eq $true }
        
        $bitlocker = Test-BitLockerEnabled
        if (-not $IncludeBIOS -or $bitlocker.Enabled) {
            $updates = $updates | Where-Object { $_.Category -notmatch "BIOS|UEFI" -and $_.Type -ne "BIOS" }
        }
        
        if (-not $updates -or $updates.Count -eq 0) {
            $result.Success = $true
            $result.Message = "No updates available"
            Write-Log $result.Message "SUCCESS"
            return $result
        }
        
        Write-Log "Installing $($updates.Count) updates..." "INFO"
        $installResults = $updates | Install-LSUpdate -ErrorAction SilentlyContinue
        
        foreach ($r in $installResults) {
            if ($r.Success -or $r.Result -eq "Installed") {
                $result.UpdateCount++
                Write-Log "  [+] $($r.Title)" "SUCCESS"
            } else {
                Write-Log "  [!] $($r.Title): $($r.FailureReason)" "WARNING"
            }
        }
        
        if (Test-Path "HKLM:\Software\LSUClient\BIOSUpdate") {
            $action = (Get-ItemProperty "HKLM:\Software\LSUClient\BIOSUpdate" -ErrorAction SilentlyContinue).ActionNeeded
            if ($action -in @("REBOOT", "SHUTDOWN")) { $result.RebootRequired = $true }
        }
        
        $result.Success = $true
        $result.Message = "Installed $($result.UpdateCount) updates"
        Write-Log $result.Message "SUCCESS"
        
    } catch {
        $result.Message = "Lenovo error: $($_.Exception.Message)"
        Write-Log $result.Message "ERROR"
    }
    
    Remove-Item (Join-Path $env:TEMP "LSUPackages") -Recurse -Force -ErrorAction SilentlyContinue
    return $result
}

# ============================================================================
# HP IMAGE ASSISTANT
# ============================================================================

function Get-HPIAPath {
    $searchPaths = @(
        "C:\ProgramData\SystemUpdatePro\HPIA",
        "C:\SWSetup\SP*",
        "${env:ProgramFiles}\HP\HPIA"
    )
    
    foreach ($path in $searchPaths) {
        $found = Get-ChildItem -Path $path -Filter "HPImageAssistant.exe" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($found) { return $found.FullName }
    }
    return $null
}

function Install-HPIA {
    Write-Log "Installing HP Image Assistant..." "INFO"
    
    return Invoke-WithRetry -OperationName "Install HPIA" -ScriptBlock {
        $hpiaDir = "C:\ProgramData\SystemUpdatePro\HPIA"
        New-Item -ItemType Directory -Path $hpiaDir -Force | Out-Null
        
        $hpiaUrl = "https://hpia.hpcloud.hp.com/downloads/hpia/hp-hpia-5.2.0.exe"
        $installer = Join-Path $env:TEMP "hp-hpia.exe"
        
        Invoke-WebRequest -Uri $hpiaUrl -OutFile $installer -UseBasicParsing -ErrorAction Stop
        Start-Process -FilePath $installer -ArgumentList "/s /e /f `"$hpiaDir`"" -Wait -NoNewWindow
        Start-Sleep -Seconds 5
        Remove-Item $installer -Force -ErrorAction SilentlyContinue
        
        $hpiaExe = Get-ChildItem -Path $hpiaDir -Filter "HPImageAssistant.exe" -Recurse | Select-Object -First 1
        if (-not $hpiaExe) { throw "HPIA not found after extraction" }
        
        Write-Log "HPIA installed" "SUCCESS"
        return $true
    }
}

function Invoke-HPUpdate {
    param([switch]$IncludeBIOS)
    
    $result = @{ Success = $false; RebootRequired = $false; UpdateCount = 0; Message = "" }
    
    Write-Log "========== HP IMAGE ASSISTANT ==========" "HEADER"
    
    $sysInfo = Get-SystemInfo
    Write-Log "Serial: $($sysInfo.SerialNumber)" "INFO"
    
    $hpiaPath = Get-HPIAPath
    if (-not $hpiaPath) {
        if (-not (Install-HPIA)) {
            $result.Message = "Failed to install HPIA"
            Write-Log $result.Message "ERROR"
            return $result
        }
        $hpiaPath = Get-HPIAPath
    }
    
    # Kill existing HPIA processes
    Get-Process -Name "HPImageAssistant*" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    
    $reportDir = Join-Path $LogPath "HPIA_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    $softpaqDir = Join-Path $env:TEMP "HPSoftpaqs"
    New-Item -ItemType Directory -Path $reportDir, $softpaqDir -Force | Out-Null
    
    $categories = @("Drivers", "Firmware")
    $bitlocker = Test-BitLockerEnabled
    if ($IncludeBIOS -and -not $bitlocker.Enabled) { $categories += "BIOS" }
    
    $hpiaArgs = @(
        "/Operation:Analyze", "/Action:Install", "/Selection:All",
        "/Category:$($categories -join ',')", "/Silent", "/Noninteractive",
        "/ReportFolder:`"$reportDir`"", "/SoftpaqDownloadFolder:`"$softpaqDir`""
    )
    
    Write-Log "Applying HP updates..." "INFO"
    
    $process = Start-Process -FilePath $hpiaPath -ArgumentList ($hpiaArgs -join " ") -Wait -NoNewWindow -PassThru
    
    Get-Process -Name "HPImageAssistant*" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    
    switch ($process.ExitCode) {
        0    { $result.Success = $true; $result.Message = "Updates applied" }
        256  { $result.Success = $true; $result.Message = "No updates needed" }
        257  { $result.Success = $true; $result.RebootRequired = $true; $result.Message = "Updates applied - reboot required" }
        3010 { $result.Success = $true; $result.RebootRequired = $true; $result.Message = "Updates applied - reboot required" }
        default { $result.Success = ($process.ExitCode -lt 256); $result.Message = "HPIA exit: $($process.ExitCode)" }
    }
    
    Write-Log $result.Message $(if ($result.Success) { "SUCCESS" } else { "WARNING" })
    
    Remove-Item $softpaqDir -Recurse -Force -ErrorAction SilentlyContinue
    return $result
}

# ============================================================================
# WINDOWS UPDATE
# ============================================================================

function Invoke-WindowsUpdateWUA {
    $result = @{ Success = $false; RebootRequired = $false; Installed = 0; Failed = 0; Message = "" }
    
    try {
        $session = New-Object -ComObject Microsoft.Update.Session
        $searcher = $session.CreateUpdateSearcher()
        
        $searchResult = $searcher.Search("IsInstalled=0 and Type='Software' and IsHidden=0")
        
        if ($searchResult.Updates.Count -eq 0) {
            $result.Success = $true
            $result.Message = "No updates available"
            return $result
        }
        
        $updatesToInstall = New-Object -ComObject Microsoft.Update.UpdateColl
        
        foreach ($update in $searchResult.Updates) {
            $dominated = ($update.Title -match "Feature Update|Upgrade to Windows|Preview")
            if (-not $dominated) {
                foreach ($cat in $update.Categories) {
                    if ($cat.Name -eq "Drivers") { $dominated = $true; break }
                }
            }
            if (-not $dominated) { $updatesToInstall.Add($update) | Out-Null }
        }
        
        if ($updatesToInstall.Count -eq 0) {
            $result.Success = $true
            $result.Message = "No applicable updates"
            return $result
        }
        
        $downloader = $session.CreateUpdateDownloader()
        $downloader.Updates = $updatesToInstall
        $downloader.Download() | Out-Null
        
        $installer = New-Object -ComObject Microsoft.Update.Installer
        $installer.Updates = $updatesToInstall
        $installResult = $installer.Install()
        
        for ($i = 0; $i -lt $updatesToInstall.Count; $i++) {
            if ($installResult.GetUpdateResult($i).ResultCode -eq 2) { $result.Installed++ }
            else { $result.Failed++ }
        }
        
        $result.Success = ($result.Installed -gt 0 -or $result.Failed -eq 0)
        $result.RebootRequired = $installResult.RebootRequired
        $result.Message = "Installed: $($result.Installed), Failed: $($result.Failed)"
        
    } catch {
        $result.Message = "WUA error: $($_.Exception.Message)"
    }
    
    return $result
}

function Invoke-WindowsUpdatePSWU {
    $result = @{ Success = $false; RebootRequired = $false; Installed = 0; Failed = 0; Message = "" }
    
    try {
        Import-Module PSWindowsUpdate -Force -ErrorAction Stop
        
        $updates = Get-WindowsUpdate -MicrosoftUpdate -NotCategory "Drivers","Feature Packs" -NotTitle "Preview" -ErrorAction SilentlyContinue
        
        if (-not $updates -or $updates.Count -eq 0) {
            $result.Success = $true
            $result.Message = "No updates available"
            return $result
        }
        
        $installResults = Install-WindowsUpdate -MicrosoftUpdate -AcceptAll -IgnoreReboot -NotCategory "Drivers","Feature Packs" -NotTitle "Preview" -Confirm:$false -ErrorAction SilentlyContinue
        
        if ($installResults) {
            foreach ($r in $installResults) {
                if ($r.Result -eq "Installed") { $result.Installed++ }
                else { $result.Failed++ }
            }
        } else {
            $result.Installed = $updates.Count
        }
        
        $result.Success = ($result.Installed -gt 0 -or $result.Failed -eq 0)
        $result.RebootRequired = (Get-WURebootStatus -Silent -ErrorAction SilentlyContinue)
        $result.Message = "Installed: $($result.Installed), Failed: $($result.Failed)"
        
    } catch {
        $result.Message = "PSWU error: $($_.Exception.Message)"
    }
    
    return $result
}

function Invoke-WindowsUpdate {
    param([int]$MaxPasses = 3)
    
    $result = @{ Success = $false; RebootRequired = $false; TotalInstalled = 0; TotalFailed = 0; Passes = 0; Message = "" }
    
    Write-Log "========== WINDOWS UPDATE ==========" "HEADER"
    
    $usePSWU = Install-PSModuleWithRetry -ModuleName "PSWindowsUpdate"
    
    for ($pass = 1; $pass -le $MaxPasses; $pass++) {
        Write-Log "Pass $pass of $MaxPasses" "INFO"
        $result.Passes = $pass
        
        $passResult = if ($usePSWU) { Invoke-WindowsUpdatePSWU } else { Invoke-WindowsUpdateWUA }
        
        $result.TotalInstalled += $passResult.Installed
        $result.TotalFailed += $passResult.Failed
        if ($passResult.RebootRequired) { $result.RebootRequired = $true }
        
        if ($passResult.Installed -eq 0) { break }
        
        if ($pass -lt $MaxPasses) { Start-Sleep -Seconds 5 }
    }
    
    $result.Success = ($result.TotalInstalled -gt 0 -or $result.TotalFailed -eq 0)
    $result.Message = "Installed: $($result.TotalInstalled), Failed: $($result.TotalFailed)"
    
    Write-Log "Windows Update: $($result.Message)" $(if ($result.Success) { "SUCCESS" } else { "WARNING" })
    return $result
}

# ============================================================================
# MAIN EXECUTION
# ============================================================================

$scriptStart = Get-Date

Write-Banner
Initialize-EventLog | Out-Null

# Check lock file
if (Test-LockFile) {
    Write-Log "Another instance is already running" "ERROR"
    Write-EventLogEntry -Message "Update blocked - another instance running" -EntryType Warning -EventId 1001
    exit 6
}

New-LockFile

# Cleanup old logs
Invoke-LogRotation -RetentionDays $LogRetentionDays

Write-Log "Log: $script:LogFile" "DEBUG"

# Check for post-reboot continuation
$state = Get-State
if ($state.Phase -eq "PostReboot") {
    Write-Log "Resuming after reboot..." "STEP"
    $script:UpdatesInstalled = $state.UpdatesInstalled
    Unregister-ContinuationTask
}

# Pre-flight checks
Write-Log "Running pre-flight checks..." "STEP"

$sysInfo = Get-SystemInfo
Write-Log "System: $($sysInfo.Manufacturer) $($sysInfo.Model)" "INFO"
Write-Log "OS: $($sysInfo.OSName) (Build $($sysInfo.OSBuild))" "DEBUG"

# Internet check
if (-not (Test-InternetConnection)) {
    Write-Log "No internet connection" "WARNING"
    [void]$script:Warnings.Add("No internet")
}

# Disk space check
$disk = Test-DiskSpace -MinGB $MinDiskSpaceGB
Write-Log "Free disk space: $($disk.FreeGB) GB (required: $($disk.RequiredGB) GB)" "DEBUG"
if (-not $disk.Sufficient) {
    Write-Log "Insufficient disk space" "ERROR"
    if (-not $Force) {
        Remove-LockFile
        exit 4
    }
}

# Pending reboot check
$reboot = Test-PendingReboot
if ($reboot.Pending) {
    Write-Log "Pending reboot: $($reboot.Reasons -join ', ')" "WARNING"
    if (-not $Force) {
        Write-Log "Use -Force to override or reboot first" "ERROR"
        Remove-LockFile
        exit 5
    }
}

# Battery check for BIOS updates
if ($IncludeBIOS) {
    $battery = Test-BatteryPower
    if ($battery.OnBattery) {
        Write-Log "Cannot update BIOS on battery power" "ERROR"
        if (-not $Force) {
            Remove-LockFile
            exit 7
        }
    }
}

# Metered connection warning
if (Test-MeteredConnection) {
    Write-Log "Metered connection detected - large downloads may incur charges" "WARNING"
}

# BitLocker status
$bitlocker = Test-BitLockerEnabled
if ($bitlocker.Enabled) {
    Write-Log "BitLocker: Active" "INFO"
}

Write-Host ""

# Repair Windows Update if requested
if ($RepairWindowsUpdate) {
    Repair-WindowsUpdateServices
    Write-Host ""
}

# WSUS bypass if requested
if ($BypassWSUS) {
    Set-WSUSBypass -Enable
}

try {
    # OEM Updates
    $oemResult = $null
    if (-not $SkipOEM) {
        $manufacturer = $sysInfo.Manufacturer.ToUpper()
        
        if ($manufacturer -match "DELL|ALIENWARE") {
            $oemResult = Invoke-DellUpdate -IncludeBIOS:$IncludeBIOS
        } elseif ($manufacturer -match "LENOVO") {
            $oemResult = Invoke-LenovoUpdate -IncludeBIOS:$IncludeBIOS
        } elseif ($manufacturer -match "HP|HEWLETT") {
            $oemResult = Invoke-HPUpdate -IncludeBIOS:$IncludeBIOS
        } else {
            Write-Log "========== OEM UPDATES ==========" "HEADER"
            Write-Log "Manufacturer '$($sysInfo.Manufacturer)' not supported" "INFO"
        }
        
        if ($oemResult) {
            if ($oemResult.RebootRequired) { $script:RebootRequired = $true }
            if (-not $oemResult.Success) { $script:ExitCode = 2 }
            $script:UpdatesInstalled += $oemResult.UpdateCount
        }
        Write-Host ""
    }
    
    # Windows Updates
    if (-not $SkipWindows) {
        $wuResult = Invoke-WindowsUpdate -MaxPasses $MaxUpdatePasses
        
        if ($wuResult.RebootRequired) { $script:RebootRequired = $true }
        if ($wuResult.TotalFailed -gt 0) { $script:ExitCode = 2 }
        $script:UpdatesInstalled += $wuResult.TotalInstalled
        $script:UpdatesFailed += $wuResult.TotalFailed
        Write-Host ""
    }
    
    # Winget Updates
    if (-not $SkipWinget) {
        $wingetResult = Invoke-WingetUpgradeAll
        if ($wingetResult.UpdateCount -gt 0) {
            $script:UpdatesInstalled += $wingetResult.UpdateCount
        }
        Write-Host ""
    }
    
    # Component cleanup
    if ($CleanupAfter) {
        Invoke-ComponentCleanup
        Write-Host ""
    }
    
} finally {
    # Restore WSUS if bypassed
    if ($BypassWSUS) {
        Set-WSUSBypass
    }
    
    # Remove lock file
    Remove-LockFile
}

# Summary
$duration = (Get-Date) - $scriptStart

Write-Log "================================================================" "HEADER"
Write-Log "  UPDATE COMPLETE" "HEADER"
Write-Log "================================================================" "HEADER"
Write-Host ""
Write-Log "System:          $($sysInfo.Manufacturer) $($sysInfo.Model)" "INFO"
Write-Log "Duration:        $([math]::Round($duration.TotalMinutes, 1)) minutes" "INFO"
Write-Log "Updates Applied: $script:UpdatesInstalled" "INFO"
if ($script:UpdatesFailed -gt 0) {
    Write-Log "Updates Failed:  $script:UpdatesFailed" "WARNING"
}
Write-Log "Reboot Required: $script:RebootRequired" $(if ($script:RebootRequired) { "WARNING" } else { "INFO" })
Write-Log "Log File:        $script:LogFile" "DEBUG"

# Final exit code
if ($script:ExitCode -eq 0 -and $script:RebootRequired) { $script:ExitCode = 1 }

Write-Log "Exit Code:       $script:ExitCode" "DEBUG"

# Event log entry
$eventMessage = @"
SystemUpdatePro completed
System: $($sysInfo.Manufacturer) $($sysInfo.Model)
Updates Applied: $script:UpdatesInstalled
Updates Failed: $script:UpdatesFailed
Reboot Required: $script:RebootRequired
Duration: $([math]::Round($duration.TotalMinutes, 1)) minutes
"@

$eventType = if ($script:ExitCode -eq 0) { "Information" } elseif ($script:ExitCode -le 2) { "Warning" } else { "Error" }
Write-EventLogEntry -Message $eventMessage -EntryType $eventType -EventId (1000 + $script:ExitCode)

Write-Host ""

# Handle reboot
if ($script:RebootRequired) {
    if ($ContinueAfterReboot) {
        Register-ContinuationTask
    }
    
    if ($Reboot) {
        Write-Log "Rebooting in 30 seconds..." "WARNING"
        shutdown.exe /r /t 30 /c "SystemUpdatePro - Reboot Required"
    }
}

# Stop transcript
try { Stop-Transcript | Out-Null } catch {}

exit $script:ExitCode
