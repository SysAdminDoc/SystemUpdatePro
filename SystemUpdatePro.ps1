<#
.SYNOPSIS
    SystemUpdatePro v4.1.0 - Enterprise Multi-OEM System Update Utility
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
    - DryRun mode for safe preview of available updates
    - HTML summary report generation
    - Webhook notifications (Slack, Teams, generic)
    - Driver backup before OEM updates
    - Update history tracking with JSON log

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
.PARAMETER ResetComponentBase
    Irreversibly remove superseded component versions; installed updates can no longer be uninstalled
.PARAMETER ContinueAfterReboot
    Create scheduled task to continue updates after reboot
.PARAMETER DryRun
    Preview available updates without installing anything
.PARAMETER BackupDrivers
    Export current drivers before installing OEM/driver updates
.PARAMETER ShowHistory
    Display update history from previous runs
.PARAMETER WebhookUrl
    URL to send completion notification (Slack, Teams, or generic webhook)
.PARAMETER HistoryCount
    Number of history entries to display with -ShowHistory (default: 10)
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
    .\SystemUpdatePro.ps1 -DryRun
    # Preview what updates are available without installing
.EXAMPLE
    .\SystemUpdatePro.ps1 -BackupDrivers -IncludeBIOS -Reboot
    # Backup drivers, update everything including BIOS, reboot
.EXAMPLE
    .\SystemUpdatePro.ps1 -WebhookUrl "https://hooks.slack.com/services/..."
    # Run updates and notify Slack on completion
.EXAMPLE
    .\SystemUpdatePro.ps1 -ShowHistory -HistoryCount 20
    # Show last 20 update runs
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
    Version: 4.1.0
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
    [switch]$ResetComponentBase,
    [switch]$ContinueAfterReboot,
    [switch]$DryRun,
    [switch]$BackupDrivers,
    [switch]$ShowHistory,
    [string]$WebhookUrl,
    [int]$HistoryCount = 10,
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

$script:Version = "4.1.0"
$script:ProductName = "SystemUpdatePro"
$script:EventLogSource = "SystemUpdatePro"
$script:ResultSchemaVersion = 1
$script:StateSchemaVersion = 2
$script:MaxContinuationAttempts = 3
$script:RunId = [guid]::NewGuid().ToString()
$script:RunStartedAt = Get-Date
$script:EntryScriptPath = [string]$PSCommandPath
$script:DataPath = "C:\ProgramData\SystemUpdatePro"
$script:LockFile = "C:\ProgramData\SystemUpdatePro\update.lock"
$script:StateFile = "C:\ProgramData\SystemUpdatePro\state.json"
$script:HistoryFile = "C:\ProgramData\SystemUpdatePro\update_history.json"
$script:TaskName = "SystemUpdatePro_Continue"

$script:ExitCode = 0
$script:RebootRequired = $false
$script:UpdatesInstalled = 0
$script:UpdatesFailed = 0
$script:Warnings = [System.Collections.ArrayList]::new()
$script:Errors = [System.Collections.ArrayList]::new()

# Tracking for HTML report and webhook
$script:OEMUpdates = [System.Collections.ArrayList]::new()
$script:WindowsUpdates = [System.Collections.ArrayList]::new()
$script:WingetUpdates = [System.Collections.ArrayList]::new()
$script:OEMUpdateCount = 0
$script:WindowsUpdateCount = 0
$script:WingetUpdateCount = 0
$script:StageResults = [System.Collections.ArrayList]::new()
$script:RunFinalized = $false
$script:TranscriptStarted = $false
$script:ContinuationAttempt = 0
$script:ContinuationActive = $false
$script:ContinuationRegistered = $false
$script:ContinuationState = $null
$script:ResumeStageCursor = ""

# Paths are declared without touching disk so read-only commands and tests can
# load the script contract before privileged initialization.
$script:LogFile = Join-Path $LogPath "$($script:ProductName)_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$script:TranscriptFile = Join-Path $LogPath "$($script:ProductName)_Transcript_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

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
        Write-EventLog -LogName "Application" -Source $script:EventLogSource -EntryType $EntryType -EventId $EventId -Message $Message -ErrorAction Stop
        return $true
    } catch {
        return $false
    }
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

    # Prefix dry run messages
    $displayMsg = "$($prefixes[$Level])$Message"
    if ($DryRun -and $Level -notin @("HEADER", "DEBUG")) {
        $displayMsg = "[DRY RUN] $displayMsg"
    }

    Write-Host $displayMsg -ForegroundColor $colors[$Level]

    # Track warnings and errors
    if ($Level -eq "WARNING") { [void]$script:Warnings.Add($Message) }
    if ($Level -eq "ERROR") { [void]$script:Errors.Add($Message) }
}

function Write-Banner {
    $mode = if ($DryRun) { " [DRY RUN MODE]" } else { "" }
    $banner = @"

  ================================================================
    $($script:ProductName) v$($script:Version)$mode
    Enterprise System Update Utility
  ================================================================

"@
    Write-Host $banner -ForegroundColor Cyan
}

# ============================================================================
# RESULT CONTRACT AND RUN INITIALIZATION
# ============================================================================

function Test-Administrator {
    try {
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch {
        return $false
    }
}

function Initialize-RunEnvironment {
    try {
        foreach ($path in @($script:DataPath, $LogPath)) {
            if (-not (Test-Path -LiteralPath $path)) {
                New-Item -ItemType Directory -Path $path -Force -ErrorAction Stop | Out-Null
            }
        }

        try {
            Start-Transcript -Path $script:TranscriptFile -Force -ErrorAction Stop | Out-Null
            $script:TranscriptStarted = $true
        } catch {
            $script:TranscriptStarted = $false
        }

        return $true
    } catch {
        [Console]::Error.WriteLine("[X] Failed to initialize run storage: $($_.Exception.Message)")
        return $false
    }
}

function New-UpdateItemResult {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseShouldProcessForStateChangingFunctions", "", Justification = "Creates an in-memory result object only.")]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,
        [string]$Id = "",
        [ValidateSet("Available", "Attempted", "Installed", "Succeeded", "Failed", "Skipped", "Blocked", "Warning", "Unknown")]
        [string]$Status = "Unknown",
        [AllowNull()][object]$ProviderCode = $null,
        [AllowNull()][object]$HResult = $null,
        [bool]$RebootRequired = $false,
        [string]$Message = "",
        [AllowEmptyCollection()][string[]]$Evidence = @(),
        [datetime]$StartedAt = (Get-Date),
        [int]$DurationSeconds = 0
    )

    $attempted = $Status -in @("Attempted", "Installed", "Succeeded", "Failed")
    return [PSCustomObject][ordered]@{
        Id               = $Id
        Name             = $Name
        Status           = $Status
        Attempted        = $attempted
        Available        = ($Status -eq "Available")
        Installed        = ($Status -eq "Installed")
        Failed           = ($Status -eq "Failed")
        Skipped          = ($Status -in @("Skipped", "Blocked"))
        ProviderExitCode = $ProviderCode
        ProviderCode     = $ProviderCode
        HResult          = $HResult
        RebootRequired   = $RebootRequired
        Message          = $Message
        Evidence         = @($Evidence)
        StartedAt        = $StartedAt.ToString("o")
        DurationSeconds  = [math]::Max(0, $DurationSeconds)
    }
}

function New-StageResult {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseShouldProcessForStateChangingFunctions", "", Justification = "Creates an in-memory result object only.")]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,
        [string]$Provider = "SystemUpdatePro",
        [ValidateSet("Succeeded", "Partial", "Failed", "Skipped")]
        [string]$Status = "Succeeded",
        [int]$Attempted = 0,
        [int]$Available = 0,
        [int]$Installed = 0,
        [int]$Failed = 0,
        [int]$Skipped = 0,
        [bool]$RebootRequired = $false,
        [AllowNull()][object]$ProviderCode = $null,
        [AllowNull()][object]$HResult = $null,
        [string]$Message = "",
        [AllowEmptyCollection()][object[]]$Items = @(),
        [AllowEmptyCollection()][string[]]$Evidence = @(),
        [datetime]$StartedAt = (Get-Date),
        [int]$DurationSeconds = 0
    )

    return [PSCustomObject][ordered]@{
        Name             = $Name
        Provider         = $Provider
        Status           = $Status
        Attempted        = [math]::Max(0, $Attempted)
        Available        = [math]::Max(0, $Available)
        Installed        = [math]::Max(0, $Installed)
        Failed           = [math]::Max(0, $Failed)
        Skipped          = [math]::Max(0, $Skipped)
        RebootRequired   = $RebootRequired
        ProviderExitCode = $ProviderCode
        ProviderCode     = $ProviderCode
        HResult          = $HResult
        Message          = $Message
        Items            = @($Items)
        Evidence         = @($Evidence)
        StartedAt        = $StartedAt.ToString("o")
        DurationSeconds  = [math]::Max(0, $DurationSeconds)
    }
}

function Get-ResultValue {
    param(
        [AllowNull()][object]$Result,
        [Parameter(Mandatory = $true)]
        [string[]]$Names,
        [AllowNull()][object]$Default = $null
    )

    if ($null -eq $Result) { return $Default }

    foreach ($name in $Names) {
        if ($Result -is [System.Collections.IDictionary] -and $Result.Contains($name) -and $null -ne $Result[$name]) {
            return $Result[$name]
        }

        $property = $Result.PSObject.Properties[$name]
        if ($property -and $null -ne $property.Value) {
            return $property.Value
        }
    }

    return $Default
}

function ConvertTo-StageResult {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,
        [Parameter(Mandatory = $true)]
        [string]$Provider,
        [AllowNull()][object]$Result,
        [AllowEmptyCollection()][string[]]$ItemNames = @(),
        [AllowEmptyCollection()][string[]]$Evidence = @(),
        [datetime]$StartedAt = (Get-Date)
    )

    if ($null -eq $Result) {
        return New-StageResult -Name $Name -Provider $Provider -Status "Skipped" -Skipped 1 `
            -Message "Stage skipped by run configuration" -StartedAt $StartedAt
    }

    $success = [bool](Get-ResultValue -Result $Result -Names @("Success") -Default $false)
    $installed = [int](Get-ResultValue -Result $Result -Names @("TotalInstalled", "Installed", "UpdateCount") -Default 0)
    $failed = [int](Get-ResultValue -Result $Result -Names @("TotalFailed", "Failed") -Default 0)
    $skipped = [int](Get-ResultValue -Result $Result -Names @("Skipped") -Default 0)
    $reportedStatus = [string](Get-ResultValue -Result $Result -Names @("Status") -Default "")
    $attempted = [int](Get-ResultValue -Result $Result -Names @("Attempted", "TotalAttempted") -Default 0)
    $available = [int](Get-ResultValue -Result $Result -Names @("Available") -Default 0)
    $message = [string](Get-ResultValue -Result $Result -Names @("Message") -Default "")
    $rebootRequired = [bool](Get-ResultValue -Result $Result -Names @("RebootRequired") -Default $false)
    $providerCode = Get-ResultValue -Result $Result -Names @("ProviderCode", "ExitCode") -Default $null
    $hresult = Get-ResultValue -Result $Result -Names @("HResult") -Default $null
    $resultEvidence = @(Get-ResultValue -Result $Result -Names @("Evidence") -Default @())
    $items = @(Get-ResultValue -Result $Result -Names @("Items") -Default @())

    if ($DryRun) {
        if ($available -eq 0) { $available = $installed }
        $installed = 0
    }

    if ($attempted -eq 0 -and -not $DryRun) {
        $attempted = $available + $installed + $failed + $skipped
    }

    if ($items.Count -eq 0 -and $ItemNames.Count -gt 0) {
        $itemStatus = if ($DryRun) {
            "Available"
        } elseif ($success) {
            "Installed"
        } else {
            "Failed"
        }

        $items = @($ItemNames | ForEach-Object {
            New-UpdateItemResult -Name $_ -Status $itemStatus -ProviderCode $providerCode -HResult $hresult `
                -RebootRequired $rebootRequired -Message $message -Evidence $resultEvidence
        })
    }

    if ($items.Count -eq 0 -and -not [string]::IsNullOrWhiteSpace($message) -and $message -notmatch "^No (updates|applicable updates)") {
        if (-not $DryRun) { $attempted = [math]::Max(1, $attempted) }
        if ($items.Count -eq 0) {
            $operationStatus = if ($success) { "Succeeded" } else { "Failed" }
            $items = @(New-UpdateItemResult -Name "$Provider operation" -Status $operationStatus `
                -ProviderCode $providerCode -HResult $hresult -RebootRequired $rebootRequired `
                -Message $message -Evidence $resultEvidence)
        }
    }

    if (-not $success -and $failed -eq 0) {
        $failed = 1
    }

    $status = if ($reportedStatus -in @("Succeeded", "Partial", "Failed", "Skipped")) {
        $reportedStatus
    } elseif (-not $success) {
        if ($installed -gt 0) { "Partial" } else { "Failed" }
    } elseif ($failed -gt 0) {
        if ($installed -gt 0) { "Partial" } else { "Failed" }
    } else {
        "Succeeded"
    }

    $duration = [math]::Max(0, [int]((Get-Date) - $StartedAt).TotalSeconds)
    return New-StageResult -Name $Name -Provider $Provider -Status $status -Attempted $attempted `
        -Available $available -Installed $installed -Failed $failed -Skipped $skipped `
        -RebootRequired $rebootRequired -ProviderCode $providerCode -HResult $hresult `
        -Message $message -Items $items -Evidence (@($Evidence) + $resultEvidence) `
        -StartedAt $StartedAt -DurationSeconds $duration
}

function Add-StageResult {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Stage
    )

    [void]$script:StageResults.Add($Stage)
    if ($Stage.RebootRequired) { $script:RebootRequired = $true }
    return $Stage
}

function Get-RunExitCode {
    param(
        [AllowEmptyCollection()][object[]]$Stages = @(),
        [int]$RequestedExitCode = 0
    )

    if ($RequestedExitCode -ge 3) { return $RequestedExitCode }

    $failedStages = @($Stages | Where-Object { $_.Status -eq "Failed" })
    $partialStages = @($Stages | Where-Object { $_.Status -eq "Partial" })
    if ($failedStages.Count -gt 0 -or $partialStages.Count -gt 0) {
        $successfulWork = @($Stages | Where-Object {
            $_.Name -in @("OEM", "WindowsUpdate", "Winget", "DriverBackup", "WindowsUpdateRepair", "Cleanup", "Continuation") -and
            ($_.Status -eq "Succeeded" -or $_.Installed -gt 0)
        }).Count
        if ($successfulWork -gt 0) { return 2 }
        return 3
    }

    if ($RequestedExitCode -eq 2) { return 2 }
    if (@($Stages | Where-Object { $_.RebootRequired }).Count -gt 0) { return 1 }
    return 0
}

function New-RunData {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseShouldProcessForStateChangingFunctions", "", Justification = "Creates an in-memory result object only.")]
    param(
        [datetime]$StartedAt,
        [datetime]$CompletedAt = (Get-Date),
        [int]$RequestedExitCode = 0
    )

    $stages = @($script:StageResults)
    $exitCode = Get-RunExitCode -Stages $stages -RequestedExitCode $RequestedExitCode
    $countProperty = if ($DryRun) { "Available" } else { "Installed" }
    $oemStage = @($stages | Where-Object { $_.Name -eq "OEM" } | Select-Object -Last 1)
    $windowsStage = @($stages | Where-Object { $_.Name -eq "WindowsUpdate" } | Select-Object -Last 1)
    $wingetStage = @($stages | Where-Object { $_.Name -eq "Winget" } | Select-Object -Last 1)

    $oemCount = if ($oemStage.Count) { [int]$oemStage[0].$countProperty } else { 0 }
    $windowsCount = if ($windowsStage.Count) { [int]$windowsStage[0].$countProperty } else { 0 }
    $wingetCount = if ($wingetStage.Count) { [int]$wingetStage[0].$countProperty } else { 0 }
    $totalInstalled = [int](($stages | Measure-Object -Property Installed -Sum).Sum)
    $totalAvailable = [int](($stages | Measure-Object -Property Available -Sum).Sum)
    $totalFailed = [int](($stages | Measure-Object -Property Failed -Sum).Sum)

    return @{
        SchemaVersion    = $script:ResultSchemaVersion
        RunId            = $script:RunId
        StartedAt        = $StartedAt.ToString("o")
        CompletedAt      = $CompletedAt.ToString("o")
        Status           = switch ($exitCode) { 0 { "Succeeded" }; 1 { "SucceededRebootRequired" }; 2 { "Partial" }; default { "Failed" } }
        OEMUpdates       = $oemCount
        WindowsUpdates   = $windowsCount
        WingetUpdates    = $wingetCount
        TotalInstalled   = $totalInstalled
        TotalAvailable   = $totalAvailable
        TotalFailed      = $totalFailed
        RebootRequired   = (@($stages | Where-Object { $_.RebootRequired }).Count -gt 0)
        ExitCode         = $exitCode
        Errors           = @($script:Errors)
        Warnings         = @($script:Warnings)
        DurationSeconds  = [math]::Max(0, [int]($CompletedAt - $StartedAt).TotalSeconds)
        Stages           = $stages
        EvidenceDelivery = @{}
    }
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

function ConvertTo-Hashtable {
    param([AllowNull()][object]$InputObject)

    if ($null -eq $InputObject) { return $null }

    if ($InputObject -is [System.Collections.IDictionary]) {
        $dictionary = [ordered]@{}
        foreach ($key in $InputObject.Keys) {
            $dictionary[[string]$key] = ConvertTo-Hashtable -InputObject $InputObject[$key]
        }
        return ,$dictionary
    }

    if ($InputObject -is [System.Management.Automation.PSCustomObject]) {
        $dictionary = [ordered]@{}
        foreach ($property in $InputObject.PSObject.Properties) {
            $dictionary[$property.Name] = ConvertTo-Hashtable -InputObject $property.Value
        }
        return ,$dictionary
    }

    if ($InputObject -is [System.Collections.IEnumerable] -and $InputObject -isnot [string]) {
        $items = @($InputObject | ForEach-Object { ConvertTo-Hashtable -InputObject $_ })
        return ,$items
    }

    return $InputObject
}

function Get-EffectiveRunParameter {
    return [ordered]@{
        SkipOEM            = $SkipOEM.IsPresent
        SkipWindows        = $SkipWindows.IsPresent
        SkipWinget         = $SkipWinget.IsPresent
        IncludeBIOS        = $IncludeBIOS.IsPresent
        BypassWSUS         = $BypassWSUS.IsPresent
        RepairWindowsUpdate = $RepairWindowsUpdate.IsPresent
        CleanupAfter       = $CleanupAfter.IsPresent
        ResetComponentBase = $ResetComponentBase.IsPresent
        ContinueAfterReboot = $ContinueAfterReboot.IsPresent
        DryRun             = $DryRun.IsPresent
        BackupDrivers      = $BackupDrivers.IsPresent
        ShowHistory        = $false
        WebhookUrl         = [string]$WebhookUrl
        HistoryCount       = [int]$HistoryCount
        MaxRetries         = [int]$MaxRetries
        MaxUpdatePasses    = [int]$MaxUpdatePasses
        MinDiskSpaceGB     = [int]$MinDiskSpaceGB
        LogPath            = [string]$LogPath
        LogRetentionDays   = [int]$LogRetentionDays
        Reboot             = $Reboot.IsPresent
        Force              = $Force.IsPresent
    }
}

function Get-ContinuationParameterName {
    return @(
        "SkipOEM", "SkipWindows", "SkipWinget", "IncludeBIOS", "BypassWSUS",
        "RepairWindowsUpdate", "CleanupAfter", "ResetComponentBase", "ContinueAfterReboot", "DryRun",
        "BackupDrivers", "ShowHistory", "WebhookUrl", "HistoryCount", "MaxRetries",
        "MaxUpdatePasses", "MinDiskSpaceGB", "LogPath", "LogRetentionDays", "Reboot", "Force"
    )
}

function Test-ContinuationState {
    param([AllowNull()][object]$State)

    $failure = {
        param([string]$Reason)
        return [PSCustomObject]@{ Valid = $false; Reason = $Reason }
    }

    if ($State -isnot [System.Collections.IDictionary]) {
        return & $failure "State root is not an object"
    }
    if ([int]$State.SchemaVersion -ne $script:StateSchemaVersion) {
        return & $failure "Unsupported state schema version"
    }
    if ([string]$State.Phase -notin @("Registering", "AwaitingReboot", "Running")) {
        return & $failure "Invalid continuation phase"
    }
    if ([string]$State.StageCursor -notin @("WindowsUpdate", "Winget", "Cleanup", "Complete")) {
        return & $failure "Invalid continuation stage cursor"
    }

    $parsedRunId = [guid]::Empty
    if (-not [guid]::TryParse([string]$State.RunId, [ref]$parsedRunId)) {
        return & $failure "Invalid continuation run ID"
    }

    $attemptCount = 0
    $maximumAttempts = 0
    if (-not [int]::TryParse([string]$State.AttemptCount, [ref]$attemptCount) -or
        -not [int]::TryParse([string]$State.MaxAttempts, [ref]$maximumAttempts) -or
        $attemptCount -lt 0 -or $maximumAttempts -lt 1 -or $maximumAttempts -gt 10 -or
        $attemptCount -gt $maximumAttempts) {
        return & $failure "Invalid continuation attempt bounds"
    }

    $createdAt = [datetime]::MinValue
    if (-not [datetime]::TryParse([string]$State.CreatedAt, [ref]$createdAt)) {
        return & $failure "Invalid continuation creation time"
    }

    if ($State.Parameters -isnot [System.Collections.IDictionary]) {
        return & $failure "Continuation parameters are missing"
    }
    foreach ($parameterName in Get-ContinuationParameterName) {
        if (-not $State.Parameters.Contains($parameterName)) {
            return & $failure "Continuation parameter '$parameterName' is missing"
        }
    }

    return [PSCustomObject]@{ Valid = $true; Reason = "" }
}

function Set-ContinuationStateAccess {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseShouldProcessForStateChangingFunctions", "", Justification = "Hardens the run-owned state file against non-administrator modification.")]
    param([string]$Path = $script:StateFile)

    try {
        $currentIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent().User
        $acl = New-Object System.Security.AccessControl.FileSecurity
        $acl.SetOwner($currentIdentity)
        $acl.SetAccessRuleProtection($true, $false)

        $identities = @(
            (New-Object System.Security.Principal.SecurityIdentifier("S-1-5-18")),
            (New-Object System.Security.Principal.SecurityIdentifier("S-1-5-32-544")),
            $currentIdentity
        ) | Select-Object -Unique

        foreach ($identity in $identities) {
            $rule = New-Object System.Security.AccessControl.FileSystemAccessRule(
                $identity,
                [System.Security.AccessControl.FileSystemRights]::FullControl,
                [System.Security.AccessControl.AccessControlType]::Allow
            )
            [void]$acl.AddAccessRule($rule)
        }

        $fileInfo = [System.IO.FileInfo](Get-Item -LiteralPath $Path -ErrorAction Stop)
        if ($PSVersionTable.PSEdition -eq "Core") {
            [System.IO.FileSystemAclExtensions]::SetAccessControl($fileInfo, $acl)
        } else {
            $fileInfo.SetAccessControl($acl)
        }
        return $true
    } catch {
        $script:LastStateAccessError = $_.Exception.Message
        return $false
    }
}

function Test-ContinuationStateAccess {
    param([string]$Path = $script:StateFile)

    try {
        $acl = Get-Acl -LiteralPath $Path -ErrorAction Stop
        if (-not $acl.AreAccessRulesProtected) {
            return [PSCustomObject]@{ Valid = $false; Reason = "State file inherits access rules" }
        }

        $broadWriterSids = @("S-1-1-0", "S-1-5-11", "S-1-5-32-545")
        $writeRights = [System.Security.AccessControl.FileSystemRights]::Write -bor
            [System.Security.AccessControl.FileSystemRights]::Modify -bor
            [System.Security.AccessControl.FileSystemRights]::FullControl
        foreach ($rule in $acl.Access) {
            if ($rule.AccessControlType -eq [System.Security.AccessControl.AccessControlType]::Allow -and
                $rule.IdentityReference.Translate([System.Security.Principal.SecurityIdentifier]).Value -in $broadWriterSids -and
                ($rule.FileSystemRights -band $writeRights)) {
                return [PSCustomObject]@{ Valid = $false; Reason = "State file is writable by a broad principal" }
            }
        }
        return [PSCustomObject]@{ Valid = $true; Reason = "" }
    } catch {
        return [PSCustomObject]@{ Valid = $false; Reason = "State ACL could not be verified: $($_.Exception.Message)" }
    }
}

function Save-State {
    param([System.Collections.IDictionary]$State)

    $temporaryPath = "$($script:StateFile).tmp.$PID.$([guid]::NewGuid().ToString('N'))"
    $backupPath = "$($script:StateFile).previous"

    try {
        $stateDirectory = Split-Path -Parent $script:StateFile
        if (-not (Test-Path -LiteralPath $stateDirectory)) {
            New-Item -ItemType Directory -Path $stateDirectory -Force -ErrorAction Stop | Out-Null
        }

        $State["LastUpdatedAt"] = (Get-Date).ToString("o")
        $json = $State | ConvertTo-Json -Depth 20
        $utf8 = New-Object System.Text.UTF8Encoding($false)
        [System.IO.File]::WriteAllText($temporaryPath, $json, $utf8)

        # Validate the complete temporary payload before replacing durable state.
        $validated = [System.IO.File]::ReadAllText($temporaryPath) | ConvertFrom-Json -ErrorAction Stop
        if ($null -eq $validated) { throw "State validation returned no data" }

        if (Test-Path -LiteralPath $script:StateFile) {
            [System.IO.File]::Replace($temporaryPath, $script:StateFile, $backupPath, $true)
            Remove-Item -LiteralPath $backupPath -Force -ErrorAction SilentlyContinue
        } else {
            [System.IO.File]::Move($temporaryPath, $script:StateFile)
        }
        if (-not (Set-ContinuationStateAccess -Path $script:StateFile)) {
            throw "State file access controls could not be hardened: $($script:LastStateAccessError)"
        }
        return $true
    } catch {
        $script:LastStateError = $_.Exception.Message
        Remove-Item -LiteralPath $temporaryPath -Force -ErrorAction SilentlyContinue
        try {
            Write-Log "Failed to save continuation state: $($_.Exception.Message)" "ERROR"
        } catch {
            [System.Diagnostics.Debug]::WriteLine($_.Exception.Message)
        }
        return $false
    }
}

function Move-StateToQuarantine {
    param([string]$Reason)

    if (-not (Test-Path -LiteralPath $script:StateFile)) { return "" }

    $directory = Split-Path -Parent $script:StateFile
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($script:StateFile)
    $quarantineName = "{0}.corrupt.{1}.{2}.json" -f $baseName, (Get-Date -Format "yyyyMMddTHHmmss"), ([guid]::NewGuid().ToString("N"))
    $quarantinePath = Join-Path $directory $quarantineName

    try {
        Move-Item -LiteralPath $script:StateFile -Destination $quarantinePath -ErrorAction Stop
        [void](Set-ContinuationStateAccess -Path $quarantinePath)
        try {
            Write-Log "Quarantined invalid continuation state: $Reason ($quarantinePath)" "WARNING"
        } catch {
            [System.Diagnostics.Debug]::WriteLine($_.Exception.Message)
        }
        return $quarantinePath
    } catch {
        try {
            Write-Log "Could not quarantine invalid continuation state: $($_.Exception.Message)" "ERROR"
        } catch {
            [System.Diagnostics.Debug]::WriteLine($_.Exception.Message)
        }
        return ""
    }
}

function Get-State {
    if (-not (Test-Path -LiteralPath $script:StateFile)) { return @{} }

    try {
        $accessValidation = Test-ContinuationStateAccess -Path $script:StateFile
        if (-not $accessValidation.Valid) { throw $accessValidation.Reason }
        $rawState = Get-Content -LiteralPath $script:StateFile -Raw -ErrorAction Stop
        $stateObject = $rawState | ConvertFrom-Json -ErrorAction Stop
        $state = ConvertTo-Hashtable -InputObject $stateObject
        $validation = Test-ContinuationState -State $state
        if (-not $validation.Valid) { throw $validation.Reason }
        return $state
    } catch {
        [void](Move-StateToQuarantine -Reason $_.Exception.Message)
        return @{}
    }
}

function Clear-State {
    try {
        Remove-Item -LiteralPath $script:StateFile -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath "$($script:StateFile).previous" -Force -ErrorAction SilentlyContinue
        return (-not (Test-Path -LiteralPath $script:StateFile))
    } catch {
        return $false
    }
}

function New-ContinuationState {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseShouldProcessForStateChangingFunctions", "", Justification = "Creates an in-memory continuation state object only.")]
    param(
        [string]$StageCursor = "WindowsUpdate",
        [string]$ScriptPath = [string]$PSCommandPath
    )

    return [ordered]@{
        SchemaVersion = $script:StateSchemaVersion
        RunId          = $script:RunId
        Phase          = "Registering"
        StageCursor    = $StageCursor
        AttemptCount   = $script:ContinuationAttempt
        MaxAttempts    = $script:MaxContinuationAttempts
        CreatedAt      = $script:RunStartedAt.ToString("o")
        LastUpdatedAt  = (Get-Date).ToString("o")
        ScriptPath     = $ScriptPath
        Parameters     = Get-EffectiveRunParameter
        StageResults   = @($script:StageResults)
        Errors         = @($script:Errors)
        Warnings       = @($script:Warnings)
    }
}

function Import-ContinuationState {
    param([System.Collections.IDictionary]$State)

    $validation = Test-ContinuationState -State $State
    if (-not $validation.Valid) {
        return [PSCustomObject]@{ Success = $false; Message = $validation.Reason }
    }

    $nextAttempt = [int]$State.AttemptCount + 1
    if ($nextAttempt -gt [int]$State.MaxAttempts) {
        return [PSCustomObject]@{ Success = $false; Message = "Continuation attempt limit reached" }
    }

    $switchNames = @(
        "SkipOEM", "SkipWindows", "SkipWinget", "IncludeBIOS", "BypassWSUS",
        "RepairWindowsUpdate", "CleanupAfter", "ResetComponentBase", "ContinueAfterReboot", "DryRun",
        "BackupDrivers", "ShowHistory", "Reboot", "Force"
    )
    $integerNames = @(
        "HistoryCount", "MaxRetries", "MaxUpdatePasses", "MinDiskSpaceGB", "LogRetentionDays"
    )

    foreach ($name in $switchNames) {
        Set-Variable -Name $name -Scope Script -Value ([switch][bool]$State.Parameters[$name])
    }
    foreach ($name in $integerNames) {
        Set-Variable -Name $name -Scope Script -Value ([int]$State.Parameters[$name])
    }
    Set-Variable -Name "WebhookUrl" -Scope Script -Value ([string]$State.Parameters.WebhookUrl)
    Set-Variable -Name "LogPath" -Scope Script -Value ([string]$State.Parameters.LogPath)

    $script:RunId = [string]$State.RunId
    $script:RunStartedAt = [datetime]::Parse([string]$State.CreatedAt)
    $script:ContinuationAttempt = $nextAttempt
    $script:ContinuationActive = $true
    $script:ResumeStageCursor = [string]$State.StageCursor
    $script:ContinuationState = $State
    $script:StageResults = [System.Collections.ArrayList]::new()
    foreach ($stageData in @($State.StageResults)) {
        $stage = [PSCustomObject]$stageData
        if ($stage.PSObject.Properties["RebootRequired"] -and [bool]$stage.RebootRequired) {
            $stage | Add-Member -NotePropertyName "RebootSatisfied" -NotePropertyValue $true -Force
            $stage.RebootRequired = $false
        }
        [void]$script:StageResults.Add($stage)
    }
    $script:Errors = [System.Collections.ArrayList]::new()
    foreach ($message in @($State.Errors)) { [void]$script:Errors.Add([string]$message) }
    $script:Warnings = [System.Collections.ArrayList]::new()
    foreach ($message in @($State.Warnings)) { [void]$script:Warnings.Add([string]$message) }

    $script:LogFile = Join-Path $LogPath "$($script:ProductName)_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    $script:TranscriptFile = Join-Path $LogPath "$($script:ProductName)_Transcript_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

    $State.Phase = "Running"
    $State.AttemptCount = $nextAttempt
    if (-not (Save-State -State $State)) {
        return [PSCustomObject]@{ Success = $false; Message = "Could not claim continuation state" }
    }

    return [PSCustomObject]@{ Success = $true; Message = "Continuation state restored" }
}

function Set-ContinuationCursor {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseShouldProcessForStateChangingFunctions", "", Justification = "Updates the run-owned continuation state snapshot.")]
    param(
        [ValidateSet("WindowsUpdate", "Winget", "Cleanup", "Complete")]
        [string]$StageCursor
    )

    if (-not $script:ContinuationActive -or $null -eq $script:ContinuationState) { return $true }

    $script:ContinuationState.StageCursor = $StageCursor
    $script:ContinuationState.StageResults = @($script:StageResults)
    $script:ContinuationState.Errors = @($script:Errors)
    $script:ContinuationState.Warnings = @($script:Warnings)
    $script:ContinuationState.Parameters = Get-EffectiveRunParameter
    return Save-State -State $script:ContinuationState
}

function Test-ShouldRunContinuationStage {
    param(
        [ValidateSet("WindowsUpdate", "Winget", "Cleanup")]
        [string]$Stage
    )

    if (-not $script:ContinuationActive) { return $true }
    $order = @("WindowsUpdate", "Winget", "Cleanup", "Complete")
    $cursorIndex = [array]::IndexOf($order, $script:ResumeStageCursor)
    $stageIndex = [array]::IndexOf($order, $Stage)
    return ($stageIndex -ge $cursorIndex)
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
# UPDATE HISTORY TRACKING
# ============================================================================

function Save-UpdateHistory {
    param(
        [hashtable]$RunData
    )

    try {
        $history = @()
        if (Test-Path $script:HistoryFile) {
            $existing = Get-Content $script:HistoryFile -Raw -ErrorAction SilentlyContinue | ConvertFrom-Json -ErrorAction SilentlyContinue
            if ($existing) {
                # ConvertFrom-Json returns array or single object
                if ($existing -is [array]) {
                    $history = @($existing)
                } else {
                    $history = @($existing)
                }
            }
        }

        # A reboot continuation retains the logical run ID. Replace the prior
        # segment so history contains one durable record per logical run.
        $history = @($history | Where-Object { [string]$_.run_id -ne [string]$RunData.RunId })

        $entry = [ordered]@{
            schema_version    = $RunData.SchemaVersion
            run_id            = $RunData.RunId
            timestamp         = $RunData.CompletedAt
            started_at        = $RunData.StartedAt
            hostname          = $env:COMPUTERNAME
            status            = $RunData.Status
            dry_run           = $DryRun.IsPresent
            oem_updates       = $RunData.OEMUpdates
            windows_updates   = $RunData.WindowsUpdates
            winget_updates    = $RunData.WingetUpdates
            total_installed   = $RunData.TotalInstalled
            total_available   = $RunData.TotalAvailable
            total_failed      = $RunData.TotalFailed
            reboot_required   = $RunData.RebootRequired
            exit_code         = $RunData.ExitCode
            errors            = @($RunData.Errors)
            warnings          = @($RunData.Warnings)
            duration_seconds  = $RunData.DurationSeconds
            stages            = @($RunData.Stages)
            evidence_delivery = $RunData.EvidenceDelivery
            parameters        = [ordered]@{
                SkipOEM           = $SkipOEM.IsPresent
                SkipWindows       = $SkipWindows.IsPresent
                SkipWinget        = $SkipWinget.IsPresent
                IncludeBIOS       = $IncludeBIOS.IsPresent
                BackupDrivers     = $BackupDrivers.IsPresent
                CleanupAfter      = $CleanupAfter.IsPresent
                ResetComponentBase = $ResetComponentBase.IsPresent
                ContinueAfterReboot = $ContinueAfterReboot.IsPresent
            }
        }

        # Prepend new entry, keep last 100 runs
        $history = @($entry) + @($history)
        if ($history.Count -gt 100) {
            $history = $history[0..99]
        }

        $history | ConvertTo-Json -Depth 12 | Set-Content -Path $script:HistoryFile -Force -Encoding UTF8 -ErrorAction Stop
        return $true
    } catch {
        return $false
    }
}

function Show-UpdateHistory {
    param([int]$Count = 10)

    if (-not (Test-Path $script:HistoryFile)) {
        Write-Host "No update history found." -ForegroundColor Yellow
        return
    }

    try {
        $history = Get-Content $script:HistoryFile -Raw | ConvertFrom-Json
        if (-not $history) {
            Write-Host "No update history found." -ForegroundColor Yellow
            return
        }

        $entries = @($history) | Select-Object -First $Count

        Write-Host ""
        Write-Host "  ================================================================" -ForegroundColor Cyan
        Write-Host "    SystemUpdatePro - Update History (Last $($entries.Count) Runs)" -ForegroundColor Cyan
        Write-Host "  ================================================================" -ForegroundColor Cyan
        Write-Host ""

        $tableData = foreach ($e in $entries) {
            $status = switch ($e.exit_code) {
                0 { "Success" }
                1 { "Success+Reboot" }
                2 { "Partial" }
                default { "Failed" }
            }
            $dryLabel = if ($e.dry_run) { " [DRY]" } else { "" }
            $dur = if ($e.duration_seconds) { "$([math]::Round($e.duration_seconds / 60, 1))m" } else { "N/A" }
            $ts = if ($e.timestamp) {
                try { ([DateTime]::Parse($e.timestamp)).ToString("yyyy-MM-dd HH:mm") } catch { $e.timestamp }
            } else { "Unknown" }

            [PSCustomObject]@{
                Date     = $ts
                Status   = "$status$dryLabel"
                OEM      = $e.oem_updates
                WinUpd   = $e.windows_updates
                Winget   = $e.winget_updates
                Failed   = $e.total_failed
                Duration = $dur
                Errors   = $(if ($e.errors) { $e.errors.Count } else { 0 })
            }
        }

        $tableData | Format-Table -AutoSize

    } catch {
        Write-Host "Error reading history: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# ============================================================================
# DRIVER BACKUP
# ============================================================================

function Invoke-DriverBackup {
    $backupRoot = Join-Path $script:DataPath "DriverBackups"
    $backupDir = Join-Path $backupRoot "Backup_$(Get-Date -Format 'yyyyMMdd_HHmmss')"

    Write-Log "========== DRIVER BACKUP ==========" "HEADER"
    Write-Log "Backing up current drivers to: $backupDir" "STEP"

    try {
        New-Item -ItemType Directory -Path $backupDir -Force | Out-Null

        if ($DryRun) {
            Write-Log "Would export drivers via Export-WindowsDriver to $backupDir" "INFO"
            # Count installed third-party drivers for dry-run info
            try {
                $driverCount = @(Get-WindowsDriver -Online -ErrorAction SilentlyContinue | Where-Object { $_.OriginalFileName -notmatch '\\windows\\' }).Count
                Write-Log "Found $driverCount third-party drivers that would be backed up" "INFO"
            } catch {
                Write-Log "Driver enumeration not available in this context" "DEBUG"
            }
            return $backupDir
        }

        Export-WindowsDriver -Online -Destination $backupDir -ErrorAction Stop | Out-Null

        $driverFiles = @(Get-ChildItem -Path $backupDir -Recurse -File -ErrorAction SilentlyContinue)
        $totalSizeMB = [math]::Round(($driverFiles | Measure-Object -Property Length -Sum).Sum / 1MB, 1)
        $driverFolders = @(Get-ChildItem -Path $backupDir -Directory -ErrorAction SilentlyContinue).Count

        Write-Log "Backed up $driverFolders drivers ($totalSizeMB MB) to $backupDir" "SUCCESS"

        # Clean up old backups (keep last 3)
        $allBackups = Get-ChildItem -Path $backupRoot -Directory -ErrorAction SilentlyContinue | Sort-Object CreationTime -Descending
        if ($allBackups.Count -gt 3) {
            $allBackups | Select-Object -Skip 3 | ForEach-Object {
                Remove-Item $_.FullName -Recurse -Force -ErrorAction SilentlyContinue
                Write-Log "Removed old backup: $($_.Name)" "DEBUG"
            }
        }

        return $backupDir
    } catch {
        Write-Log "Driver backup failed: $($_.Exception.Message)" "WARNING"
        return $null
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

    if ($DryRun) {
        Write-Log "Would stop WU services, clear cache, re-register DLLs, reset Winsock, restart services" "INFO"
        return $true
    }

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

        if (-not $DryRun) {
            # Disable WSUS
            if (Test-Path $auPath) {
                Set-ItemProperty -Path $auPath -Name UseWUServer -Value 0 -Type DWord -ErrorAction SilentlyContinue
            }

            # Restart Windows Update service
            Restart-Service wuauserv -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 3
        }

        Write-Log "WSUS bypass enabled" "SUCCESS"
    } else {
        # Restore original settings
        if ($script:WSUSBackup -and $script:WSUSBackup.UseWUServer) {
            Write-Log "Restoring WSUS settings..." "DEBUG"

            if (-not $DryRun) {
                if (Test-Path $auPath) {
                    Set-ItemProperty -Path $auPath -Name UseWUServer -Value $script:WSUSBackup.UseWUServer -Type DWord -ErrorAction SilentlyContinue
                }

                Restart-Service wuauserv -Force -ErrorAction SilentlyContinue
            }
        }
    }
}

# ============================================================================
# POST-REBOOT CONTINUATION
# ============================================================================

function Register-ContinuationTask {
    Write-Log "Creating post-reboot continuation task..." "DEBUG"

    if ($DryRun) {
        Write-Log "Would create scheduled task '$($script:TaskName)' for post-reboot continuation" "INFO"
        return $true
    }

    if ($script:ContinuationAttempt -ge $script:MaxContinuationAttempts) {
        Write-Log "Continuation attempt limit reached ($($script:MaxContinuationAttempts))" "ERROR"
        return $false
    }

    $scriptPath = [string]$script:EntryScriptPath
    if ([string]::IsNullOrWhiteSpace($scriptPath) -or -not (Test-Path -LiteralPath $scriptPath)) {
        Write-Log "Cannot register continuation because the script path is unavailable" "ERROR"
        return $false
    }

    $state = New-ContinuationState -StageCursor "WindowsUpdate" -ScriptPath $scriptPath
    if (-not (Save-State -State $state)) {
        Write-Log "Continuation state could not be committed; task was not registered" "ERROR"
        return $false
    }

    try {
        # Keep the task command line free of saved parameters and webhook URLs;
        # the validated state file is the sole resume contract.
        Unregister-ScheduledTask -TaskName $script:TaskName -Confirm:$false -ErrorAction SilentlyContinue

        $powershellPath = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"
        $arguments = "-NoProfile -NonInteractive -ExecutionPolicy Bypass -File `"$scriptPath`""
        $action = New-ScheduledTaskAction -Execute $powershellPath -Argument $arguments
        $trigger = New-ScheduledTaskTrigger -AtStartup
        $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
        $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
            -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Hours 4)

        Register-ScheduledTask -TaskName $script:TaskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force | Out-Null

        $state.Phase = "AwaitingReboot"
        if (-not (Save-State -State $state)) {
            throw "Scheduled task was created but final state commit failed"
        }

        $script:ContinuationState = $state
        $script:ContinuationRegistered = $true
        Write-Log "Continuation task registered: $($script:TaskName)" "SUCCESS"
        return $true
    } catch {
        Unregister-ScheduledTask -TaskName $script:TaskName -Confirm:$false -ErrorAction SilentlyContinue
        [void](Clear-State)
        $script:ContinuationRegistered = $false
        Write-Log "Failed to create continuation task: $($_.Exception.Message)" "WARNING"
        return $false
    }
}

function Test-ContinuationTask {
    try {
        return $null -ne (Get-ScheduledTask -TaskName $script:TaskName -ErrorAction SilentlyContinue)
    } catch {
        return $false
    }
}

function Unregister-ContinuationTask {
    param([switch]$PreserveState)

    $success = $true
    try {
        if (Test-ContinuationTask) {
            Unregister-ScheduledTask -TaskName $script:TaskName -Confirm:$false -ErrorAction Stop
        }
    } catch {
        $success = $false
        try {
            Write-Log "Failed to remove continuation task: $($_.Exception.Message)" "ERROR"
        } catch {
            [System.Diagnostics.Debug]::WriteLine($_.Exception.Message)
        }
    }

    if (-not $PreserveState -and -not (Clear-State)) {
        $success = $false
        try {
            Write-Log "Failed to clear continuation state" "ERROR"
        } catch {
            [System.Diagnostics.Debug]::WriteLine($_.Exception.Message)
        }
    }

    if ($success -and -not $PreserveState) {
        $script:ContinuationState = $null
        $script:ContinuationActive = $false
    }
    return $success
}

# ============================================================================
# CLEANUP
# ============================================================================

function Get-RegistryValueSnapshot {
    param(
        [string]$Path,
        [string]$Name
    )

    $key = Get-Item -LiteralPath $Path -ErrorAction Stop
    $valueNames = @($key.GetValueNames())
    $exists = $valueNames -contains $Name
    return [ordered]@{
        Path   = $Path
        Name   = $Name
        Exists = $exists
        Value  = $(if ($exists) { $key.GetValue($Name, $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames) } else { $null })
        Kind   = $(if ($exists) { $key.GetValueKind($Name).ToString() } else { "" })
    }
}

function Restore-RegistryValueSnapshot {
    param([System.Collections.IDictionary]$Snapshot)

    try {
        if ([bool]$Snapshot.Exists) {
            New-ItemProperty -LiteralPath $Snapshot.Path -Name $Snapshot.Name -Value $Snapshot.Value `
                -PropertyType $Snapshot.Kind -Force -ErrorAction Stop | Out-Null
        } else {
            Remove-ItemProperty -LiteralPath $Snapshot.Path -Name $Snapshot.Name -Force -ErrorAction SilentlyContinue
        }

        $restored = Get-RegistryValueSnapshot -Path $Snapshot.Path -Name $Snapshot.Name
        if ([bool]$restored.Exists -ne [bool]$Snapshot.Exists) { return $false }
        if ($Snapshot.Exists -and
            ([string]$restored.Kind -ne [string]$Snapshot.Kind -or [string]$restored.Value -ne [string]$Snapshot.Value)) {
            return $false
        }
        return $true
    } catch {
        return $false
    }
}

function Invoke-ComponentCleanup {
    $dismArguments = "/Online /Cleanup-Image /StartComponentCleanup"
    $rollbackImpact = "Installed update uninstallability is retained"
    if ($ResetComponentBase) {
        $dismArguments += " /ResetBase"
        $rollbackImpact = "IRREVERSIBLE: installed Windows updates can no longer be uninstalled"
    }

    $result = @{
        Success = $false; Status = "Failed"; RebootRequired = $false
        Attempted = 0; Installed = 0; Failed = 0; Skipped = 0
        ExitCode = $null; HResult = $null; Items = @()
        Evidence = @("dism.exe $dismArguments", "cleanmgr.exe /sagerun:100")
        ResetBase = $ResetComponentBase.IsPresent; RollbackImpact = $rollbackImpact; Message = ""
    }

    Write-Log "Running DISM component cleanup..." "STEP"
    if ($ResetComponentBase) {
        Write-Log $rollbackImpact "WARNING"
    } else {
        Write-Log $rollbackImpact "INFO"
    }

    if ($DryRun) {
        Write-Log "Would run DISM $dismArguments" "INFO"
        Write-Log "Would run Disk Cleanup and restore all temporary StateFlags0100 values" "INFO"
        $result.Success = $true
        $result.Status = "Succeeded"
        $result.Skipped = 2
        $result.Message = "Cleanup plan only; $rollbackImpact"
        $result.Items = @(
            (New-UpdateItemResult -Name "DISM component cleanup" -Status "Skipped" -Message $dismArguments),
            (New-UpdateItemResult -Name "Disk Cleanup" -Status "Skipped" -Message "StateFlags0100 values would be restored")
        )
        return $result
    }

    $dismSucceeded = $false
    try {
        $result.Attempted++
        $process = Start-Process -FilePath "dism.exe" -ArgumentList $dismArguments -Wait -NoNewWindow -PassThru -ErrorAction Stop
        $result.ExitCode = $process.ExitCode
        if ($process.ExitCode -ne 0) {
            $result.Failed++
            $result.Items += New-UpdateItemResult -Name "DISM component cleanup" -Status "Failed" `
                -ProviderCode $process.ExitCode -Message $rollbackImpact -Evidence @("dism.exe $dismArguments")
            $result.Message = "DISM cleanup returned code $($process.ExitCode); $rollbackImpact"
            Write-Log $result.Message "WARNING"
            return $result
        }

        $dismSucceeded = $true
        $result.Items += New-UpdateItemResult -Name "DISM component cleanup" -Status "Succeeded" `
            -ProviderCode $process.ExitCode -Message $rollbackImpact -Evidence @("dism.exe $dismArguments")
        Write-Log "Component cleanup completed; $rollbackImpact" "SUCCESS"
    } catch {
        $result.Failed++
        $result.HResult = $_.Exception.HResult
        $result.Items += New-UpdateItemResult -Name "DISM component cleanup" -Status "Failed" `
            -HResult $_.Exception.HResult -Message $_.Exception.Message -Evidence @("dism.exe $dismArguments")
        $result.Message = "DISM cleanup error: $($_.Exception.Message)"
        Write-Log $result.Message "WARNING"
        return $result
    }

    $volCachePath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches"
    $cleanupItems = @(
        "Update Cleanup",
        "Windows Update Cleanup",
        "Temporary Files",
        "System error memory dump files",
        "Delivery Optimization Files"
    )
    $snapshots = [System.Collections.ArrayList]::new()
    $flagFailures = [System.Collections.ArrayList]::new()
    $cleanMgrSucceeded = $false

    try {
        foreach ($item in $cleanupItems) {
            $itemPath = Join-Path $volCachePath $item
            if (-not (Test-Path -LiteralPath $itemPath)) { continue }

            try {
                $snapshot = Get-RegistryValueSnapshot -Path $itemPath -Name "StateFlags0100"
                [void]$snapshots.Add($snapshot)
                New-ItemProperty -LiteralPath $itemPath -Name "StateFlags0100" -Value 2 `
                    -PropertyType DWord -Force -ErrorAction Stop | Out-Null
            } catch {
                [void]$flagFailures.Add("$item`: $($_.Exception.Message)")
            }
        }

        $result.Attempted++
        $cleanMgr = Start-Process -FilePath "cleanmgr.exe" -ArgumentList "/sagerun:100" `
            -Wait -NoNewWindow -PassThru -ErrorAction Stop
        if ($cleanMgr.ExitCode -eq 0) {
            $cleanMgrSucceeded = $true
        } else {
            $result.ExitCode = $cleanMgr.ExitCode
            [void]$flagFailures.Add("cleanmgr exit code $($cleanMgr.ExitCode)")
        }
    } catch {
        $result.HResult = $_.Exception.HResult
        [void]$flagFailures.Add($_.Exception.Message)
    } finally {
        foreach ($snapshot in @($snapshots)) {
            if (-not (Restore-RegistryValueSnapshot -Snapshot $snapshot)) {
                [void]$flagFailures.Add("Could not restore $($snapshot.Path)\$($snapshot.Name)")
            }
        }
    }

    if ($cleanMgrSucceeded -and $flagFailures.Count -eq 0) {
        $result.Success = $true
        $result.Status = "Succeeded"
        $result.Items += New-UpdateItemResult -Name "Disk Cleanup" -Status "Succeeded" `
            -ProviderCode 0 -Message "Temporary registry flags restored" -Evidence @("cleanmgr.exe /sagerun:100")
        $result.Message = "Component and disk cleanup completed; $rollbackImpact"
    } else {
        $result.Failed += [math]::Max(1, $flagFailures.Count)
        $result.Status = $(if ($dismSucceeded) { "Partial" } else { "Failed" })
        $failureMessage = if ($flagFailures.Count) { $flagFailures -join "; " } else { "Disk Cleanup failed" }
        $result.Items += New-UpdateItemResult -Name "Disk Cleanup" -Status "Failed" `
            -ProviderCode $result.ExitCode -HResult $result.HResult -Message $failureMessage `
            -Evidence @("cleanmgr.exe /sagerun:100")
        $result.Message = "Component cleanup completed with Disk Cleanup failures: $failureMessage; $rollbackImpact"
        Write-Log $result.Message "WARNING"
    }

    return $result
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

    if ($DryRun) {
        Write-Log "Would install Winget and dependencies (VCLibs, UI.Xaml)" "INFO"
        return $true
    }

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
    $result = @{
        Success = $false; RebootRequired = $false
        UpdateCount = 0; Available = 0; Attempted = 0; Installed = 0; Failed = 0; Skipped = 0
        ExitCode = $null; HResult = $null; Items = @(); Evidence = @(); Message = ""
    }

    Write-Log "========== WINGET UPGRADE ALL ==========" "HEADER"

    if (-not (Test-WingetInstalled)) {
        if (-not (Install-Winget)) {
            $result.Message = "Winget not available"
            $result.Failed = 1
            Write-Log $result.Message "WARNING"
            return $result
        }
    }

    try {
        # Update sources
        & winget source update --disable-interactivity 2>&1 | Out-Null

        if ($DryRun) {
            Write-Log "Checking for available winget upgrades..." "INFO"
            $listOutput = & winget upgrade --include-unknown 2>&1
            $result.ExitCode = $LASTEXITCODE
            $upgradeLines = @($listOutput | Where-Object { $_ -match '\S' -and $_ -notmatch '^(-|Name |\\|The following)' -and $_ -notmatch 'upgrades available' })

            # Count available upgrades (rough parse)
            $availCount = 0
            foreach ($line in $listOutput) {
                if ($line -match '(\d+) upgrades available') {
                    $availCount = [int]$Matches[1]
                    break
                }
            }
            if ($availCount -eq 0 -and $upgradeLines.Count -gt 2) {
                $availCount = $upgradeLines.Count - 2  # Subtract header lines
            }

            $result.UpdateCount = [math]::Max(0, $availCount)
            $result.Available = $result.UpdateCount
            $result.Success = ($result.ExitCode -in @(0, -1978335189))
            if (-not $result.Success) { $result.Failed = 1 }
            $result.Message = "$($result.UpdateCount) winget upgrades available (dry run - not installed)"
            foreach ($line in @($upgradeLines | Select-Object -First $result.UpdateCount)) {
                [void]$script:WingetUpdates.Add(([string]$line).Trim())
            }
            Write-Log $result.Message "INFO"
            $script:WingetUpdateCount = $result.UpdateCount
            return $result
        }

        Write-Log "Running winget upgrade --all..." "INFO"

        $wingetArgs = @(
            "upgrade", "--all", "--silent",
            "--accept-package-agreements", "--accept-source-agreements",
            "--disable-interactivity", "--include-unknown"
        )

        $process = Start-Process -FilePath "winget" -ArgumentList $wingetArgs -Wait -NoNewWindow -PassThru

        $result.Attempted = 1
        $result.ExitCode = $process.ExitCode
        $result.Success = ($process.ExitCode -in @(0, -1978335189))  # 0 or "no updates"
        if (-not $result.Success) { $result.Failed = 1 }
        $result.Message = "Winget upgrade completed (exit: $($process.ExitCode))"

        Write-Log $result.Message $(if ($result.Success) { "SUCCESS" } else { "WARNING" })

    } catch {
        $result.Message = "Winget error: $($_.Exception.Message)"
        $result.Failed = 1
        $result.HResult = $_.Exception.HResult
        Write-Log $result.Message "ERROR"
    }

    $script:WingetUpdateCount = $result.UpdateCount
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

    if ($DryRun) {
        Write-Log "Would install $ModuleName from PSGallery" "INFO"
        return $true
    }

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

    if ($DryRun) {
        Write-Log "Would install Dell Command Update via winget" "INFO"
        return $true
    }

    if (-not (Test-WingetInstalled)) {
        if (-not (Install-Winget)) { return $false }
    }

    return Invoke-WithRetry -OperationName "Install DCU" -ScriptBlock {
        & winget source update --disable-interactivity 2>&1 | Out-Null

        $wingetArgs = @("install", "--id", "Dell.CommandUpdate", "--source", "winget",
                  "--accept-package-agreements", "--accept-source-agreements", "--silent")

        Start-Process -FilePath "winget" -ArgumentList $wingetArgs -Wait -NoNewWindow
        Start-Sleep -Seconds 5

        if (-not (Get-DCUPath)) { throw "DCU not found after install" }

        Write-Log "Dell Command Update installed" "SUCCESS"
        return $true
    }
}

function Repair-DellServices {
    if ($DryRun) { return $true }

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

    $result = @{
        Success = $false; RebootRequired = $false
        UpdateCount = 0; Available = 0; Attempted = 0; Installed = 0; Failed = 0; Skipped = 0
        ExitCode = $null; HResult = $null; Items = @(); Evidence = @(); Message = ""
    }

    try {

    Write-Log "========== DELL COMMAND UPDATE ==========" "HEADER"

    $sysInfo = Get-SystemInfo
    Write-Log "Service Tag: $($sysInfo.SerialNumber)" "INFO"

    $dcuPath = Get-DCUPath
    if (-not $dcuPath) {
        if (-not (Install-DellCommandUpdate)) {
            $result.Message = "Failed to install DCU"
            $result.Failed = 1
            Write-Log $result.Message "ERROR"
            return $result
        }
        $dcuPath = Get-DCUPath
    }

    if (-not (Repair-DellServices)) {
        Write-Log "Dell service issues - proceeding anyway" "WARNING"
    }

    if (-not $DryRun) {
        # Disable bloat services
        Get-Service -ErrorAction SilentlyContinue | Where-Object {
            ($_.DisplayName -like "*Dell*" -or $_.Name -like "*DDV*" -or $_.Name -like "*SupportAssist*") -and
            $_.Name -ne "DellClientManagementService"
        } | ForEach-Object {
            Stop-Service -Name $_.Name -Force -ErrorAction SilentlyContinue
            Set-Service -Name $_.Name -StartupType Disabled -ErrorAction SilentlyContinue
        }
    }

    # Build arguments
    $bitlocker = Test-BitLockerEnabled

    if ($DryRun) {
        # Scan only, do not install
        $dcuArgs = @("/scan", "-silent", "-updateSeverity=security,critical,recommended")
        if (-not $IncludeBIOS -or ($bitlocker.Enabled -and -not $IncludeBIOS)) {
            $dcuArgs += "-updateType=driver,firmware,application"
        }

        $dcuLog = Join-Path $LogPath "DCU_DryRun_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
        $dcuArgs += "-outputLog=`"$dcuLog`""

        Write-Log "Scanning for available Dell updates (dry run)..." "INFO"

        if ($dcuPath) {
            $process = Start-Process -FilePath $dcuPath -ArgumentList ($dcuArgs -join " ") -Wait -NoNewWindow -PassThru
            $result.Attempted = 1
            $result.ExitCode = $process.ExitCode
            $result.Success = ($process.ExitCode -in @(0, 500))
            if (-not $result.Success) { $result.Failed = 1 }
            $result.Message = "DCU scan completed (exit: $($process.ExitCode)) - no updates installed (dry run)"
        } else {
            $result.Success = $true
            $result.Message = "DCU not installed - would install and scan (dry run)"
        }

        Write-Log $result.Message $(if ($result.Success) { "INFO" } else { "WARNING" })
        $result.Evidence = @($dcuLog)
        $script:OEMUpdateCount = $result.UpdateCount
        return $result
    }

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
        $result.Attempted++
        $result.ExitCode = $process.ExitCode

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

    if (-not $result.Success) { $result.Failed = 1 }
    $result.Evidence = @($dcuLog)
    Write-Log $result.Message $(if ($result.Success) { "SUCCESS" } else { "WARNING" })
    $script:OEMUpdateCount = $result.UpdateCount
    return $result
    } catch {
        $result.Success = $false
        $result.Failed = [math]::Max(1, $result.Failed)
        $result.HResult = $_.Exception.HResult
        $result.Message = "Dell update error: $($_.Exception.Message)"
        Write-Log $result.Message "ERROR"
        return $result
    }
}

# ============================================================================
# LENOVO LSUClient
# ============================================================================

function Invoke-LenovoUpdate {
    param([switch]$IncludeBIOS)

    $result = @{
        Success = $false; RebootRequired = $false
        UpdateCount = 0; Available = 0; Attempted = 0; Installed = 0; Failed = 0; Skipped = 0
        ExitCode = $null; HResult = $null; Items = @(); Evidence = @(); Message = ""
    }

    Write-Log "========== LENOVO SYSTEM UPDATE ==========" "HEADER"

    $sysInfo = Get-SystemInfo
    Write-Log "Serial: $($sysInfo.SerialNumber)" "INFO"

    if (-not (Install-PSModuleWithRetry -ModuleName "LSUClient")) {
        $result.Message = "Failed to install LSUClient"
        $result.Failed = 1
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
            $script:OEMUpdateCount = 0
            return $result
        }

        if ($DryRun) {
            $result.UpdateCount = $updates.Count
            $result.Available = $updates.Count
            $result.Success = $true
            $result.Message = "$($updates.Count) updates available (dry run - not installed)"
            Write-Log "Available Lenovo updates:" "INFO"
            foreach ($u in $updates) {
                Write-Log "  -- $($u.Title) ($($u.Category))" "INFO"
                [void]$script:OEMUpdates.Add("$($u.Title) ($($u.Category))")
            }
            Write-Log $result.Message "INFO"
            $script:OEMUpdateCount = $result.UpdateCount
            return $result
        }

        Write-Log "Installing $($updates.Count) updates..." "INFO"
        $result.Attempted = $updates.Count
        $installResults = $updates | Install-LSUpdate -ErrorAction SilentlyContinue

        if (-not $installResults) {
            $result.Failed = $updates.Count
            foreach ($update in $updates) {
                $result.Items += New-UpdateItemResult -Name $update.Title -Id $update.ID -Status "Failed" `
                    -Message "LSUClient returned no installation result"
            }
        } else {
            foreach ($r in $installResults) {
                $itemName = if ($r.Title) { $r.Title } else { "Lenovo update" }
                if ($r.Success -or $r.Result -eq "Installed") {
                    $result.UpdateCount++
                    $result.Installed++
                    $result.Items += New-UpdateItemResult -Name $itemName -Id $r.ID -Status "Installed" `
                        -ProviderCode $r.Result -RebootRequired ([bool]$r.RebootRequired)
                    Write-Log "  [+] $itemName" "SUCCESS"
                } else {
                    $result.Failed++
                    $result.Items += New-UpdateItemResult -Name $itemName -Id $r.ID -Status "Failed" `
                        -ProviderCode $r.Result -Message ([string]$r.FailureReason)
                    Write-Log "  [!] $itemName`: $($r.FailureReason)" "WARNING"
                }
            }
            if ($result.Items.Count -lt $result.Attempted) {
                $missingResults = $result.Attempted - $result.Items.Count
                $result.Failed += $missingResults
                $result.Items += New-UpdateItemResult -Name "$missingResults update result(s) missing" -Status "Failed" `
                    -Message "LSUClient returned fewer results than requested"
            }
        }

        if (Test-Path "HKLM:\Software\LSUClient\BIOSUpdate") {
            $action = (Get-ItemProperty "HKLM:\Software\LSUClient\BIOSUpdate" -ErrorAction SilentlyContinue).ActionNeeded
            if ($action -in @("REBOOT", "SHUTDOWN")) { $result.RebootRequired = $true }
        }

        $result.Success = ($result.Failed -eq 0 -and $result.Items.Count -eq $result.Attempted)
        $result.Message = "Installed: $($result.Installed), Failed: $($result.Failed)"
        Write-Log $result.Message $(if ($result.Success) { "SUCCESS" } else { "WARNING" })

    } catch {
        $result.Message = "Lenovo error: $($_.Exception.Message)"
        $result.Failed = [math]::Max(1, $result.Failed)
        $result.HResult = $_.Exception.HResult
        Write-Log $result.Message "ERROR"
    }

    Remove-Item (Join-Path $env:TEMP "LSUPackages") -Recurse -Force -ErrorAction SilentlyContinue
    $script:OEMUpdateCount = $result.UpdateCount
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

    if ($DryRun) {
        Write-Log "Would download and install HP Image Assistant" "INFO"
        return $true
    }

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

    $result = @{
        Success = $false; RebootRequired = $false
        UpdateCount = 0; Available = 0; Attempted = 0; Installed = 0; Failed = 0; Skipped = 0
        ExitCode = $null; HResult = $null; Items = @(); Evidence = @(); Message = ""
    }

    try {

    Write-Log "========== HP IMAGE ASSISTANT ==========" "HEADER"

    $sysInfo = Get-SystemInfo
    Write-Log "Serial: $($sysInfo.SerialNumber)" "INFO"

    $hpiaPath = Get-HPIAPath
    if (-not $hpiaPath) {
        if (-not (Install-HPIA)) {
            $result.Message = "Failed to install HPIA"
            $result.Failed = 1
            Write-Log $result.Message "ERROR"
            return $result
        }
        $hpiaPath = Get-HPIAPath
    }

    if (-not $DryRun) {
        # Kill existing HPIA processes
        Get-Process -Name "HPImageAssistant*" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    }

    $reportDir = Join-Path $LogPath "HPIA_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    $softpaqDir = Join-Path $env:TEMP "HPSoftpaqs"
    New-Item -ItemType Directory -Path $reportDir, $softpaqDir -Force | Out-Null

    $categories = @("Drivers", "Firmware")
    $bitlocker = Test-BitLockerEnabled
    if ($IncludeBIOS -and -not $bitlocker.Enabled) { $categories += "BIOS" }

    if ($DryRun) {
        # Analyze only, do not install
        $hpiaArgs = @(
            "/Operation:Analyze", "/Action:List", "/Selection:All",
            "/Category:$($categories -join ',')", "/Silent", "/Noninteractive",
            "/ReportFolder:`"$reportDir`""
        )

        Write-Log "Scanning for available HP updates (dry run)..." "INFO"

        if ($hpiaPath) {
            $process = Start-Process -FilePath $hpiaPath -ArgumentList ($hpiaArgs -join " ") -Wait -NoNewWindow -PassThru
            $result.Attempted = 1
            $result.ExitCode = $process.ExitCode
            $result.Success = ($process.ExitCode -in @(0, 256, 257, 3010))
            if (-not $result.Success) { $result.Failed = 1 }
            $result.Message = "HPIA scan completed (exit: $($process.ExitCode)) - no updates installed (dry run)"
        } else {
            $result.Success = $true
            $result.Message = "HPIA not installed - would install and scan (dry run)"
        }

        Write-Log $result.Message $(if ($result.Success) { "INFO" } else { "WARNING" })
        $result.Evidence = @($reportDir)
        $script:OEMUpdateCount = $result.UpdateCount
        return $result
    }

    $hpiaArgs = @(
        "/Operation:Analyze", "/Action:Install", "/Selection:All",
        "/Category:$($categories -join ',')", "/Silent", "/Noninteractive",
        "/ReportFolder:`"$reportDir`"", "/SoftpaqDownloadFolder:`"$softpaqDir`""
    )

    Write-Log "Applying HP updates..." "INFO"

    $process = Start-Process -FilePath $hpiaPath -ArgumentList ($hpiaArgs -join " ") -Wait -NoNewWindow -PassThru
    $result.Attempted = 1
    $result.ExitCode = $process.ExitCode

    Get-Process -Name "HPImageAssistant*" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

    switch ($process.ExitCode) {
        0    { $result.Success = $true; $result.Message = "Updates applied" }
        256  { $result.Success = $true; $result.Message = "No updates needed" }
        257  { $result.Success = $true; $result.RebootRequired = $true; $result.Message = "Updates applied - reboot required" }
        3010 { $result.Success = $true; $result.RebootRequired = $true; $result.Message = "Updates applied - reboot required" }
        default { $result.Success = $false; $result.Message = "HPIA exit: $($process.ExitCode)" }
    }
    if (-not $result.Success) { $result.Failed = 1 }

    Write-Log $result.Message $(if ($result.Success) { "SUCCESS" } else { "WARNING" })

    Remove-Item $softpaqDir -Recurse -Force -ErrorAction SilentlyContinue
    $result.Evidence = @($reportDir)
    $script:OEMUpdateCount = $result.UpdateCount
    return $result
    } catch {
        $result.Success = $false
        $result.Failed = [math]::Max(1, $result.Failed)
        $result.HResult = $_.Exception.HResult
        $result.Message = "HP update error: $($_.Exception.Message)"
        Write-Log $result.Message "ERROR"
        return $result
    }
}

# ============================================================================
# WINDOWS UPDATE
# ============================================================================

function Invoke-WindowsUpdateWUA {
    $result = @{
        Success = $false; RebootRequired = $false
        Available = 0; Attempted = 0; Installed = 0; Failed = 0; Skipped = 0
        ExitCode = $null; HResult = $null; Items = @(); Evidence = @(); Message = ""
    }

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

        $result.Available = $updatesToInstall.Count

        if ($DryRun) {
            $result.Success = $true
            $result.Message = "$($updatesToInstall.Count) updates available (dry run - not installed)"
            for ($i = 0; $i -lt $updatesToInstall.Count; $i++) {
                $update = $updatesToInstall.Item($i)
                $uTitle = $update.Title
                Write-Log "  -- $uTitle" "INFO"
                [void]$script:WindowsUpdates.Add($uTitle)
                $result.Items += New-UpdateItemResult -Name $uTitle -Id $update.Identity.UpdateID -Status "Available"
            }
            return $result
        }

        $downloader = $session.CreateUpdateDownloader()
        $downloader.Updates = $updatesToInstall
        $downloadResult = $downloader.Download()
        if ($downloadResult.ResultCode -notin @(2, 3)) {
            $result.ExitCode = $downloadResult.ResultCode
            $result.HResult = $downloadResult.HResult
            $result.Failed = $updatesToInstall.Count
            $result.Message = "WUA download failed (result: $($downloadResult.ResultCode), HRESULT: $($downloadResult.HResult))"
            return $result
        }

        $installer = New-Object -ComObject Microsoft.Update.Installer
        $installer.Updates = $updatesToInstall
        $installResult = $installer.Install()
        $result.Attempted = $updatesToInstall.Count
        $result.ExitCode = $installResult.ResultCode
        $result.HResult = $installResult.HResult

        for ($i = 0; $i -lt $updatesToInstall.Count; $i++) {
            $update = $updatesToInstall.Item($i)
            $itemResult = $installResult.GetUpdateResult($i)
            $itemStatus = if ($itemResult.ResultCode -eq 2) { "Installed" } else { "Failed" }
            if ($itemStatus -eq "Installed") { $result.Installed++ } else { $result.Failed++ }
            [void]$script:WindowsUpdates.Add($update.Title)
            $result.Items += New-UpdateItemResult -Name $update.Title -Id $update.Identity.UpdateID `
                -Status $itemStatus -ProviderCode $itemResult.ResultCode -HResult $itemResult.HResult `
                -RebootRequired ([bool]$itemResult.RebootRequired)
        }

        $result.Success = ($result.Failed -eq 0 -and $installResult.ResultCode -in @(2, 3))
        $result.RebootRequired = $installResult.RebootRequired
        $result.Message = "Installed: $($result.Installed), Failed: $($result.Failed)"

    } catch {
        $result.Message = "WUA error: $($_.Exception.Message)"
        $result.HResult = $_.Exception.HResult
        $result.Failed = [math]::Max(1, $result.Failed)
    }

    return $result
}

function Invoke-WindowsUpdatePSWU {
    $result = @{
        Success = $false; RebootRequired = $false
        Available = 0; Attempted = 0; Installed = 0; Failed = 0; Skipped = 0
        ExitCode = $null; HResult = $null; Items = @(); Evidence = @(); Message = ""
    }

    try {
        Import-Module PSWindowsUpdate -Force -ErrorAction Stop

        $updates = @(Get-WindowsUpdate -MicrosoftUpdate -NotCategory "Drivers","Feature Packs" -NotTitle "Preview" -ErrorAction Stop)

        if (-not $updates -or $updates.Count -eq 0) {
            $result.Success = $true
            $result.Message = "No updates available"
            return $result
        }

        $result.Available = $updates.Count

        if ($DryRun) {
            $result.Success = $true
            $result.Message = "$($updates.Count) updates available (dry run - not installed)"
            foreach ($u in $updates) {
                $uTitle = if ($u.Title) { $u.Title } else { $u.KB }
                Write-Log "  -- $uTitle" "INFO"
                [void]$script:WindowsUpdates.Add($uTitle)
                $result.Items += New-UpdateItemResult -Name $uTitle -Id ([string]($u.KB -join ",")) -Status "Available"
            }
            return $result
        }

        $result.Attempted = $updates.Count
        $installResults = @(Install-WindowsUpdate -MicrosoftUpdate -AcceptAll -IgnoreReboot -NotCategory "Drivers","Feature Packs" -NotTitle "Preview" -Confirm:$false -ErrorAction Stop)

        if ($installResults.Count -gt 0) {
            foreach ($r in $installResults) {
                $itemName = if ($r.Title) { $r.Title } else { [string]($r.KB -join ",") }
                $itemStatus = if ($r.Result -eq "Installed") { "Installed" } else { "Failed" }
                if ($itemStatus -eq "Installed") { $result.Installed++ } else { $result.Failed++ }
                [void]$script:WindowsUpdates.Add($itemName)
                $result.Items += New-UpdateItemResult -Name $itemName -Id ([string]($r.KB -join ",")) `
                    -Status $itemStatus -ProviderCode $r.Result -HResult $r.HResult `
                    -RebootRequired ([bool]$r.RebootRequired)
            }
            if ($installResults.Count -lt $result.Attempted) {
                $missingResults = $result.Attempted - $installResults.Count
                $result.Failed += $missingResults
                $result.Items += New-UpdateItemResult -Name "$missingResults update result(s) missing" -Status "Failed" `
                    -Message "PSWindowsUpdate returned fewer results than requested"
            }
        } else {
            $result.Failed = $result.Attempted
            foreach ($u in $updates) {
                $itemName = if ($u.Title) { $u.Title } else { [string]($u.KB -join ",") }
                $result.Items += New-UpdateItemResult -Name $itemName -Id ([string]($u.KB -join ",")) `
                    -Status "Failed" -Message "PSWindowsUpdate returned no installation result"
            }
        }

        $result.Success = ($result.Failed -eq 0 -and $result.Items.Count -eq $result.Attempted)
        $result.RebootRequired = (Get-WURebootStatus -Silent -ErrorAction SilentlyContinue)
        $result.Message = "Installed: $($result.Installed), Failed: $($result.Failed)"

    } catch {
        $result.Message = "PSWU error: $($_.Exception.Message)"
        $result.HResult = $_.Exception.HResult
        $result.Failed = [math]::Max(1, $result.Failed)
    }

    return $result
}

function Invoke-WindowsUpdate {
    param([int]$MaxPasses = 3)

    $result = @{
        Success = $false; RebootRequired = $false
        Available = 0; Attempted = 0; TotalInstalled = 0; TotalFailed = 0; Skipped = 0
        ExitCode = $null; HResult = $null; Items = @(); Evidence = @(); Passes = 0; Message = ""
    }

    Write-Log "========== WINDOWS UPDATE ==========" "HEADER"

    $usePSWU = Install-PSModuleWithRetry -ModuleName "PSWindowsUpdate"
    $providerSucceeded = $true

    for ($pass = 1; $pass -le $MaxPasses; $pass++) {
        Write-Log "Pass $pass of $MaxPasses" "INFO"
        $result.Passes = $pass

        $passResult = if ($usePSWU) { Invoke-WindowsUpdatePSWU } else { Invoke-WindowsUpdateWUA }

        $result.Available += [int]$passResult.Available
        $result.Attempted += [int]$passResult.Attempted
        $result.TotalInstalled += $passResult.Installed
        $result.TotalFailed += $passResult.Failed
        $result.Items += @($passResult.Items)
        $result.Evidence += @($passResult.Evidence)
        $result.ExitCode = $passResult.ExitCode
        $result.HResult = $passResult.HResult
        if ($passResult.RebootRequired) { $result.RebootRequired = $true }

        if (-not $passResult.Success) {
            $providerSucceeded = $false
            $result.Message = "Pass $pass failed: $($passResult.Message)"
            break
        }

        if ($passResult.Installed -eq 0) { break }

        # In dry run, one pass is enough (we just list what is available)
        if ($DryRun) { break }

        if ($pass -lt $MaxPasses) { Start-Sleep -Seconds 5 }
    }

    $result.Success = ($providerSucceeded -and $result.TotalFailed -eq 0)
    if ([string]::IsNullOrWhiteSpace($result.Message)) {
        $result.Message = if ($DryRun) {
            "Available: $($result.Available)"
        } else {
            "Installed: $($result.TotalInstalled), Failed: $($result.TotalFailed)"
        }
    }

    $label = if ($DryRun) { "Windows Update (available)" } else { "Windows Update" }
    Write-Log "$label`: $($result.Message)" $(if ($result.Success) { "SUCCESS" } else { "WARNING" })
    $script:WindowsUpdateCount = if ($DryRun) { $result.Available } else { $result.TotalInstalled }
    return $result
}

# ============================================================================
# HTML REPORT GENERATION
# ============================================================================

function New-HTMLReport {
    param(
        [hashtable]$SysInfo,
        [hashtable]$RunData
    )

    $generatedAt = Get-Date
    $reportFile = Join-Path $LogPath "SystemUpdatePro_Report_$($generatedAt.ToString('yyyyMMdd_HHmmss')).html"

    $encode = {
        param([AllowNull()][object]$Value)

        if ($null -eq $Value) { return "" }
        return [System.Net.WebUtility]::HtmlEncode([string]$Value)
    }

    $displayValue = {
        param(
            [AllowNull()][object]$Value,
            [string]$Fallback = "Not reported"
        )

        if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) {
            return & $encode $Fallback
        }
        return & $encode $Value
    }

    $exitCode = [int]$RunData.ExitCode
    $statusKey = switch ($exitCode) {
        0 { "success" }
        1 { "reboot" }
        2 { "partial" }
        default { "failed" }
    }
    $overallStatus = switch ($exitCode) {
        0 { "SUCCESS" }
        1 { "SUCCESS · REBOOT REQUIRED" }
        2 { "PARTIAL SUCCESS" }
        default { "FAILED" }
    }
    $statusGlyph = switch ($statusKey) {
        "success" { "&#10003;" }
        "reboot" { "&#8635;" }
        "partial" { "!" }
        default { "&#215;" }
    }

    $durationSeconds = [math]::Max(0, [int]$RunData.DurationSeconds)
    $duration = [TimeSpan]::FromSeconds($durationSeconds)
    if ($duration.TotalHours -ge 1) {
        $durationDisplay = "{0}h {1}m" -f [math]::Floor($duration.TotalHours), $duration.Minutes
    } elseif ($duration.TotalMinutes -ge 1) {
        $durationDisplay = "{0}m {1}s" -f [math]::Floor($duration.TotalMinutes), $duration.Seconds
    } else {
        $durationDisplay = "{0}s" -f $duration.Seconds
    }

    $oemCount = [math]::Max(0, [int]$RunData.OEMUpdates)
    $windowsCount = [math]::Max(0, [int]$RunData.WindowsUpdates)
    $wingetCount = [math]::Max(0, [int]$RunData.WingetUpdates)
    $totalUpdates = $oemCount + $windowsCount + $wingetCount
    $modeLabel = if ($DryRun) { "DRY RUN" } else { "LIVE RUN" }
    $summaryVerb = if ($DryRun) { "identified" } else { "processed" }
    $rebootLabel = if ($RunData.RebootRequired) { "Required" } else { "Not required" }
    $rebootClass = if ($RunData.RebootRequired) { "metric--reboot" } else { "metric--clear" }
    $componentRollbackDisplay = & $encode $(if ($ResetComponentBase) {
        "Disabled by irreversible /ResetBase"
    } elseif ($CleanupAfter) {
        "Retained after standard cleanup"
    } else {
        "Not changed"
    })

    $getChannelState = {
        param(
            [int]$Count,
            [bool]$Skipped
        )

        if ($Skipped) {
            return @{ Label = "Skipped"; Class = "state--quiet" }
        }
        if ($Count -gt 0 -and $DryRun) {
            return @{ Label = "Available"; Class = "state--attention" }
        }
        if ($Count -gt 0) {
            return @{ Label = "Processed"; Class = "state--good" }
        }
        return @{ Label = "No changes"; Class = "state--quiet" }
    }

    $oemState = & $getChannelState $oemCount $SkipOEM.IsPresent
    $windowsState = & $getChannelState $windowsCount $SkipWindows.IsPresent
    $wingetState = & $getChannelState $wingetCount $SkipWinget.IsPresent

    $buildUpdateList = {
        param(
            [AllowNull()][object[]]$Items,
            [string]$EmptyMessage
        )

        $safeItems = @($Items | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
        if ($safeItems.Count -eq 0) {
            return "<span class='channel-empty'>$(& $encode $EmptyMessage)</span>"
        }

        $listItems = ""
        $visibleItems = @($safeItems | Select-Object -First 2)
        foreach ($item in $visibleItems) {
            $listItems += "<li>$(& $encode $item)</li>"
        }
        if ($safeItems.Count -gt $visibleItems.Count) {
            $remaining = $safeItems.Count - $visibleItems.Count
            $listItems += "<li class='more'>+$remaining more</li>"
        }
        return "<ul class='channel-details'>$listItems</ul>"
    }

    $oemEmpty = if ($SkipOEM) { "Channel skipped by run configuration" } elseif ($DryRun) { "No OEM updates available" } else { "No OEM changes recorded" }
    $windowsEmpty = if ($SkipWindows) { "Channel skipped by run configuration" } elseif ($DryRun) { "No Windows updates available" } else { "No Windows changes recorded" }
    $wingetEmpty = if ($SkipWinget) {
        "Channel skipped by run configuration"
    } elseif ($DryRun) {
        "$wingetCount package upgrades detected by winget"
    } elseif ($wingetCount -gt 0) {
        "$wingetCount package upgrades processed by winget"
    } else {
        "No application package changes recorded"
    }

    $oemDetails = & $buildUpdateList @($script:OEMUpdates) $oemEmpty
    $windowsDetails = & $buildUpdateList @($script:WindowsUpdates) $windowsEmpty
    $wingetDetails = & $buildUpdateList @($script:WingetUpdates) $wingetEmpty

    $noticeRows = ""
    $errorCount = @($RunData.Errors).Count
    $warningCount = @($RunData.Warnings).Count
    $attentionCount = $errorCount + $warningCount

    foreach ($err in @($RunData.Errors)) {
        if ([string]::IsNullOrWhiteSpace([string]$err)) { continue }
        $noticeRows += @"
<li class="notice notice--error">
  <span class="notice-icon" aria-hidden="true">!</span>
  <div><strong>Error</strong><p>$(& $encode $err)</p></div>
</li>
"@
    }
    foreach ($warning in @($RunData.Warnings)) {
        if ([string]::IsNullOrWhiteSpace([string]$warning)) { continue }
        $noticeRows += @"
<li class="notice notice--warning">
  <span class="notice-icon" aria-hidden="true">!</span>
  <div><strong>Warning</strong><p>$(& $encode $warning)</p></div>
</li>
"@
    }
    if ([string]::IsNullOrWhiteSpace($noticeRows)) {
        $noticeRows = @"
<li class="notice notice--clear">
  <span class="notice-icon" aria-hidden="true">&#10003;</span>
  <div><strong>No attention needed</strong><p>The run completed without recorded errors or warnings.</p></div>
</li>
"@
    }

    $attentionSummary = if ($attentionCount -eq 0) {
        "All clear"
    } elseif ($attentionCount -eq 1) {
        "1 item"
    } else {
        "$attentionCount items"
    }

    $biosDate = "Not reported"
    if ($SysInfo.BIOSDate) {
        try {
            $biosDate = ([DateTime]$SysInfo.BIOSDate).ToString("yyyy-MM-dd")
        } catch {
            $biosDate = [string]$SysInfo.BIOSDate
        }
    }

    $computerName = & $displayValue $env:COMPUTERNAME "Unknown endpoint"
    $manufacturer = & $displayValue $SysInfo.Manufacturer
    $model = & $displayValue $SysInfo.Model
    $serialNumber = & $displayValue $SysInfo.SerialNumber
    $osName = & $displayValue $SysInfo.OSName
    $osBuild = & $displayValue $SysInfo.OSBuild
    $biosVersion = & $displayValue $SysInfo.BIOSVersion
    $biosDateDisplay = & $displayValue $biosDate
    $processor = & $displayValue $SysInfo.Processor
    $totalRam = if ($null -eq $SysInfo.TotalRAM -or [string]::IsNullOrWhiteSpace([string]$SysInfo.TotalRAM)) {
        "Not reported"
    } else {
        "$(& $encode $SysInfo.TotalRAM) GB"
    }
    $logFileDisplay = & $displayValue $script:LogFile
    $generatedDisplay = $generatedAt.ToString("yyyy-MM-dd HH:mm:ss")
    $versionDisplay = & $encode $script:Version

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta name="color-scheme" content="dark">
<title>SystemUpdatePro Report - $computerName</title>
<style>
  :root {
    color-scheme: dark;
    --ink: #071019;
    --ink-soft: #0a1520;
    --surface: #0f1b27;
    --surface-strong: #122231;
    --surface-muted: #0b1722;
    --line: #26394b;
    --line-soft: #1b2c3b;
    --text: #eef5fb;
    --muted: #93a8ba;
    --muted-strong: #b8c7d3;
    --cyan: #36c5f0;
    --mint: #52d6a7;
    --amber: #f5b942;
    --coral: #ff7a74;
    --violet: #9f8cff;
    --shadow: 0 20px 55px rgba(0, 0, 0, 0.22);
  }

  * {
    box-sizing: border-box;
  }

  html {
    background: var(--ink);
  }

  body {
    margin: 0;
    min-width: 320px;
    background:
      radial-gradient(circle at 8% -10%, rgba(54, 197, 240, 0.08), transparent 30rem),
      var(--ink);
    color: var(--text);
    font-family: "Segoe UI Variable Text", "Segoe UI", system-ui, -apple-system, sans-serif;
    font-size: 15px;
    line-height: 1.5;
    -webkit-font-smoothing: antialiased;
  }

  .page-shell {
    width: min(1320px, calc(100% - 48px));
    margin: 0 auto;
    padding: 28px 0 36px;
  }

  .brand-bar {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 24px;
    min-height: 44px;
    margin-bottom: 18px;
  }

  .brand {
    display: inline-flex;
    align-items: center;
    gap: 12px;
    min-width: 0;
  }

  .brand-mark {
    display: grid;
    width: 38px;
    height: 38px;
    place-items: center;
    color: var(--cyan);
    background: rgba(54, 197, 240, 0.08);
    border: 1px solid rgba(54, 197, 240, 0.34);
    border-radius: 10px;
  }

  .brand-mark svg,
  .metric-icon svg,
  .channel-icon svg {
    width: 100%;
    height: 100%;
  }

  .brand-name {
    font-size: 17px;
    font-weight: 680;
    letter-spacing: -0.015em;
  }

  .brand-version {
    color: var(--muted);
    font-family: "Cascadia Mono", "SFMono-Regular", Consolas, monospace;
    font-size: 12px;
  }

  .brand-meta {
    color: var(--muted);
    font-size: 12px;
    text-align: right;
  }

  .hero {
    --status-accent: var(--cyan);
    position: relative;
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 28px;
    min-height: 174px;
    padding: 32px 34px;
    overflow: hidden;
    background:
      linear-gradient(rgba(54, 197, 240, 0.035) 1px, transparent 1px),
      linear-gradient(90deg, rgba(54, 197, 240, 0.035) 1px, transparent 1px),
      linear-gradient(120deg, #102130 0%, #0d1b28 58%, #0a1722 100%);
    background-size: 32px 32px, 32px 32px, auto;
    border: 1px solid var(--line);
    border-radius: 14px;
    box-shadow: var(--shadow);
  }

  .hero::after {
    position: absolute;
    inset: auto -60px -120px auto;
    width: 360px;
    height: 260px;
    background: radial-gradient(circle, color-mix(in srgb, var(--status-accent) 14%, transparent), transparent 68%);
    content: "";
    pointer-events: none;
  }

  .hero--success { --status-accent: var(--mint); }
  .hero--reboot,
  .hero--partial { --status-accent: var(--amber); }
  .hero--failed { --status-accent: var(--coral); }

  .hero-main {
    position: relative;
    z-index: 1;
    display: flex;
    align-items: center;
    gap: 24px;
    min-width: 0;
  }

  .status-symbol {
    display: grid;
    flex: 0 0 auto;
    width: 74px;
    height: 74px;
    place-items: center;
    color: var(--status-accent);
    background: color-mix(in srgb, var(--status-accent) 9%, transparent);
    border: 2px solid color-mix(in srgb, var(--status-accent) 72%, transparent);
    border-radius: 22px;
    font-family: "Cascadia Mono", Consolas, monospace;
    font-size: 36px;
    font-weight: 800;
  }

  .eyebrow {
    margin: 0 0 4px;
    color: var(--muted-strong);
    font-family: "Cascadia Mono", "SFMono-Regular", Consolas, monospace;
    font-size: 13px;
    font-weight: 650;
    letter-spacing: 0.11em;
    text-transform: uppercase;
  }

  .hero h1 {
    margin: 0;
    color: var(--status-accent);
    font-size: clamp(30px, 4vw, 50px);
    font-weight: 760;
    letter-spacing: -0.035em;
    line-height: 1.06;
  }

  .hero-summary {
    margin: 10px 0 0;
    color: var(--muted-strong);
    font-size: 16px;
  }

  .hero-summary strong {
    color: var(--cyan);
    font-weight: 680;
  }

  .hero-side {
    position: relative;
    z-index: 1;
    display: flex;
    flex: 0 0 auto;
    flex-direction: column;
    align-items: flex-end;
    gap: 10px;
  }

  .run-pill,
  .exit-token,
  .section-count,
  .channel-state {
    display: inline-flex;
    align-items: center;
    width: fit-content;
    border-radius: 999px;
    white-space: nowrap;
  }

  .run-pill {
    gap: 8px;
    padding: 9px 14px;
    color: var(--status-accent);
    background: color-mix(in srgb, var(--status-accent) 10%, transparent);
    border: 1px solid color-mix(in srgb, var(--status-accent) 52%, transparent);
    font-size: 12px;
    font-weight: 750;
    letter-spacing: 0.08em;
  }

  .run-pill::before {
    width: 7px;
    height: 7px;
    background: currentColor;
    border-radius: 50%;
    box-shadow: 0 0 0 4px color-mix(in srgb, currentColor 12%, transparent);
    content: "";
  }

  .exit-token {
    padding: 5px 10px;
    color: var(--muted);
    background: rgba(7, 16, 25, 0.46);
    border: 1px solid var(--line-soft);
    font-family: "Cascadia Mono", Consolas, monospace;
    font-size: 11px;
  }

  .metrics {
    display: grid;
    grid-template-columns: repeat(4, minmax(0, 1fr));
    gap: 14px;
    margin: 16px 0;
  }

  .metric {
    display: grid;
    grid-template-columns: 46px 1fr;
    gap: 14px;
    align-items: center;
    min-height: 96px;
    padding: 18px;
    background: var(--surface);
    border: 1px solid var(--line);
    border-radius: 12px;
  }

  .metric-icon {
    display: grid;
    width: 46px;
    height: 46px;
    place-items: center;
    padding: 10px;
    color: var(--cyan);
    background: rgba(54, 197, 240, 0.08);
    border: 1px solid rgba(54, 197, 240, 0.28);
    border-radius: 11px;
  }

  .metric--windows .metric-icon {
    color: #53a8ff;
    background: rgba(83, 168, 255, 0.08);
    border-color: rgba(83, 168, 255, 0.28);
  }

  .metric--apps .metric-icon {
    color: var(--mint);
    background: rgba(82, 214, 167, 0.08);
    border-color: rgba(82, 214, 167, 0.28);
  }

  .metric--reboot .metric-icon {
    color: var(--amber);
    background: rgba(245, 185, 66, 0.08);
    border-color: rgba(245, 185, 66, 0.32);
  }

  .metric--clear .metric-icon {
    color: var(--mint);
    background: rgba(82, 214, 167, 0.08);
    border-color: rgba(82, 214, 167, 0.28);
  }

  .metric-value {
    display: block;
    color: var(--text);
    font-size: 28px;
    font-weight: 740;
    letter-spacing: -0.03em;
    line-height: 1.08;
  }

  .metric-value--text {
    font-size: 18px;
    letter-spacing: -0.015em;
  }

  .metric-label {
    display: block;
    margin-top: 5px;
    color: var(--muted);
    font-size: 12px;
    font-weight: 600;
    letter-spacing: 0.035em;
    text-transform: uppercase;
  }

  .dashboard-grid {
    display: grid;
    grid-template-columns: minmax(0, 1.72fr) minmax(310px, 0.88fr);
    gap: 16px;
  }

  .panel {
    min-width: 0;
    overflow: hidden;
    background: var(--surface);
    border: 1px solid var(--line);
    border-radius: 12px;
  }

  .panel-heading {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 16px;
    padding: 17px 20px;
    background: var(--surface-muted);
    border-bottom: 1px solid var(--line-soft);
  }

  .panel-heading h2 {
    margin: 0;
    font-size: 16px;
    font-weight: 690;
    letter-spacing: -0.012em;
  }

  .panel-heading p {
    margin: 3px 0 0;
    color: var(--muted);
    font-size: 12px;
  }

  .section-count {
    padding: 4px 9px;
    color: var(--muted-strong);
    background: rgba(147, 168, 186, 0.08);
    border: 1px solid var(--line);
    font-family: "Cascadia Mono", Consolas, monospace;
    font-size: 11px;
  }

  .channels {
    grid-column: 1;
  }

  .channel-list {
    padding: 6px 18px 10px;
  }

  .channel {
    display: grid;
    grid-template-columns: 44px minmax(144px, 0.95fr) 72px 104px minmax(200px, 1.2fr);
    gap: 14px;
    align-items: center;
    min-height: 96px;
    padding: 16px 4px;
    border-bottom: 1px solid var(--line-soft);
  }

  .channel:last-child {
    border-bottom: 0;
  }

  .channel-icon {
    display: grid;
    width: 42px;
    height: 42px;
    place-items: center;
    padding: 10px;
    color: var(--violet);
    background: rgba(159, 140, 255, 0.1);
    border: 1px solid rgba(159, 140, 255, 0.28);
    border-radius: 10px;
  }

  .channel--windows .channel-icon {
    color: #53a8ff;
    background: rgba(83, 168, 255, 0.09);
    border-color: rgba(83, 168, 255, 0.28);
  }

  .channel--apps .channel-icon {
    color: var(--mint);
    background: rgba(82, 214, 167, 0.09);
    border-color: rgba(82, 214, 167, 0.28);
  }

  .channel-title {
    display: block;
    color: var(--text);
    font-weight: 660;
  }

  .channel-caption {
    display: block;
    margin-top: 3px;
    color: var(--muted);
    font-size: 12px;
  }

  .channel-count {
    color: var(--text);
    font-size: 26px;
    font-weight: 730;
    line-height: 1;
    text-align: center;
  }

  .channel-count small {
    display: block;
    margin-top: 7px;
    color: var(--muted);
    font-size: 10px;
    font-weight: 650;
    letter-spacing: 0.06em;
    text-transform: uppercase;
  }

  .channel-state {
    justify-content: center;
    gap: 7px;
    padding: 5px 9px;
    border: 1px solid;
    font-size: 10px;
    font-weight: 760;
    letter-spacing: 0.045em;
    text-transform: uppercase;
  }

  .channel-state::before {
    width: 6px;
    height: 6px;
    background: currentColor;
    border-radius: 50%;
    content: "";
  }

  .state--good {
    color: var(--mint);
    background: rgba(82, 214, 167, 0.07);
    border-color: rgba(82, 214, 167, 0.24);
  }

  .state--attention {
    color: var(--amber);
    background: rgba(245, 185, 66, 0.07);
    border-color: rgba(245, 185, 66, 0.26);
  }

  .state--quiet {
    color: var(--muted-strong);
    background: rgba(147, 168, 186, 0.06);
    border-color: var(--line);
  }

  .channel-details {
    min-width: 0;
    margin: 0;
    padding: 0;
    color: var(--muted-strong);
    font-family: "Cascadia Mono", "SFMono-Regular", Consolas, monospace;
    font-size: 11px;
    line-height: 1.55;
    list-style: none;
  }

  .channel-details li {
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
  }

  .channel-details .more,
  .channel-empty {
    color: var(--muted);
  }

  .channel-empty {
    font-size: 12px;
  }

  .device-profile {
    grid-column: 2;
    grid-row: 1 / span 2;
  }

  .profile-list {
    margin: 0;
    padding: 8px 20px 12px;
  }

  .profile-row {
    display: grid;
    grid-template-columns: minmax(100px, 0.78fr) minmax(0, 1.22fr);
    gap: 16px;
    margin: 0;
    padding: 11px 0;
    border-bottom: 1px solid var(--line-soft);
  }

  .profile-row:last-child {
    border-bottom: 0;
  }

  .profile-row dt {
    color: var(--muted);
  }

  .profile-row dd {
    min-width: 0;
    margin: 0;
    overflow-wrap: anywhere;
    color: var(--text);
    font-family: "Cascadia Mono", "SFMono-Regular", Consolas, monospace;
    font-size: 12px;
  }

  .attention {
    grid-column: 1;
  }

  .notice-list {
    margin: 0;
    padding: 10px 18px 14px;
    list-style: none;
  }

  .notice {
    display: grid;
    grid-template-columns: 30px 1fr;
    gap: 12px;
    align-items: start;
    margin-top: 8px;
    padding: 12px 14px;
    border: 1px solid;
    border-radius: 9px;
  }

  .notice-icon {
    display: grid;
    width: 26px;
    height: 26px;
    place-items: center;
    border: 1px solid currentColor;
    border-radius: 8px;
    font-family: "Cascadia Mono", Consolas, monospace;
    font-size: 13px;
    font-weight: 800;
  }

  .notice strong {
    display: block;
    margin-bottom: 2px;
    font-size: 12px;
    letter-spacing: 0.045em;
    text-transform: uppercase;
  }

  .notice p {
    margin: 0;
    color: var(--muted-strong);
    font-size: 13px;
  }

  .notice--error {
    color: var(--coral);
    background: rgba(255, 122, 116, 0.06);
    border-color: rgba(255, 122, 116, 0.22);
  }

  .notice--warning {
    color: var(--amber);
    background: rgba(245, 185, 66, 0.06);
    border-color: rgba(245, 185, 66, 0.22);
  }

  .notice--clear {
    color: var(--mint);
    background: rgba(82, 214, 167, 0.06);
    border-color: rgba(82, 214, 167, 0.2);
  }

  .run-details {
    grid-column: 1 / -1;
  }

  .detail-grid {
    display: grid;
    grid-template-columns: repeat(4, minmax(0, 1fr));
    margin: 0;
    padding: 18px 20px;
  }

  .detail-item {
    min-width: 0;
    margin: 0;
    padding: 0 18px;
    border-left: 1px solid var(--line-soft);
  }

  .detail-item:first-child {
    padding-left: 0;
    border-left: 0;
  }

  .detail-item dt {
    margin-bottom: 5px;
    color: var(--muted);
    font-size: 10px;
    font-weight: 650;
    letter-spacing: 0.07em;
    text-transform: uppercase;
  }

  .detail-item dd {
    margin: 0;
    color: var(--text);
    font-family: "Cascadia Mono", "SFMono-Regular", Consolas, monospace;
    font-size: 12px;
    overflow-wrap: anywhere;
  }

  .log-path {
    grid-column: 1 / -1;
    display: grid;
    grid-template-columns: 112px minmax(0, 1fr);
    gap: 14px;
    margin: 16px 20px 20px;
    padding-top: 16px;
    border-top: 1px solid var(--line-soft);
  }

  .log-path span {
    color: var(--muted);
    font-size: 11px;
    font-weight: 650;
    letter-spacing: 0.055em;
    text-transform: uppercase;
  }

  .log-path code {
    color: var(--muted-strong);
    font-family: "Cascadia Mono", "SFMono-Regular", Consolas, monospace;
    font-size: 11px;
    overflow-wrap: anywhere;
  }

  .footer {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 20px;
    margin-top: 20px;
    padding: 0 2px;
    color: var(--muted);
    font-size: 11px;
  }

  .footer strong {
    color: var(--muted-strong);
    font-weight: 650;
  }

  @media (max-width: 980px) {
    .metrics {
      grid-template-columns: repeat(2, minmax(0, 1fr));
    }

    .dashboard-grid {
      grid-template-columns: 1fr;
    }

    .channels,
    .device-profile,
    .attention,
    .run-details {
      grid-column: 1;
      grid-row: auto;
    }

    .device-profile {
      order: 3;
    }

    .run-details {
      order: 4;
    }
  }

  @media (max-width: 700px) {
    .page-shell {
      width: min(100% - 24px, 1320px);
      padding-top: 16px;
    }

    .brand-meta {
      display: none;
    }

    .hero {
      align-items: flex-start;
      min-height: 0;
      padding: 24px;
    }

    .hero-main {
      align-items: flex-start;
    }

    .status-symbol {
      width: 52px;
      height: 52px;
      border-radius: 15px;
      font-size: 26px;
    }

    .hero-side {
      position: absolute;
      top: 18px;
      right: 18px;
    }

    .exit-token {
      display: none;
    }

    .eyebrow {
      padding-right: 108px;
      font-size: 11px;
    }

    .hero h1 {
      font-size: clamp(26px, 8vw, 38px);
    }

    .metrics {
      gap: 10px;
    }

    .metric {
      grid-template-columns: 38px 1fr;
      min-height: 82px;
      padding: 14px;
    }

    .metric-icon {
      width: 38px;
      height: 38px;
      padding: 8px;
    }

    .metric-value {
      font-size: 23px;
    }

    .metric-value--text {
      font-size: 15px;
    }

    .channel {
      grid-template-columns: 42px minmax(0, 1fr) 64px;
      gap: 12px;
    }

    .channel-state {
      grid-column: 2;
      justify-self: start;
    }

    .channel > :last-child {
      grid-column: 2 / -1;
    }

    .detail-grid {
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 18px 0;
    }

    .detail-item:nth-child(odd) {
      padding-left: 0;
      border-left: 0;
    }
  }

  @media (max-width: 480px) {
    .brand-version {
      display: none;
    }

    .hero-main {
      gap: 14px;
    }

    .status-symbol {
      display: none;
    }

    .eyebrow {
      padding-right: 94px;
    }

    .run-pill {
      padding: 7px 10px;
      font-size: 10px;
    }

    .metrics {
      grid-template-columns: 1fr;
    }

    .profile-row {
      grid-template-columns: 1fr;
      gap: 4px;
    }

    .detail-grid {
      grid-template-columns: 1fr;
    }

    .detail-item,
    .detail-item:nth-child(odd) {
      padding: 0;
      border: 0;
    }

    .log-path {
      grid-template-columns: 1fr;
      gap: 5px;
    }

    .footer {
      align-items: flex-start;
      flex-direction: column;
      gap: 4px;
    }
  }

  @media print {
    :root {
      color-scheme: light;
      --ink: #ffffff;
      --ink-soft: #ffffff;
      --surface: #ffffff;
      --surface-strong: #f7f9fb;
      --surface-muted: #f3f6f8;
      --line: #c9d2d9;
      --line-soft: #dde3e8;
      --text: #13202b;
      --muted: #526779;
      --muted-strong: #334a5c;
      --shadow: none;
    }

    @page {
      size: landscape;
      margin: 10mm;
    }

    body {
      background: #ffffff;
      print-color-adjust: exact;
      -webkit-print-color-adjust: exact;
    }

    .page-shell {
      width: 100%;
      padding: 0;
    }

    .hero {
      background: #f3f6f8;
    }

    .hero::after {
      display: none;
    }

    .panel,
    .metric,
    .hero {
      break-inside: avoid;
      box-shadow: none;
    }
  }
</style>
</head>
<body>
<main class="page-shell">
  <header class="brand-bar">
    <div class="brand" aria-label="SystemUpdatePro">
      <span class="brand-mark" aria-hidden="true">
        <svg viewBox="0 0 24 24" fill="none">
          <path d="M12 2.7 19 5.4v5.4c0 4.7-2.7 8.5-7 10.5-4.3-2-7-5.8-7-10.5V5.4L12 2.7Z" stroke="currentColor" stroke-width="1.7"/>
          <path d="m8.6 11.9 2.1 2.1 4.7-5" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"/>
        </svg>
      </span>
      <span class="brand-name">SystemUpdatePro</span>
      <span class="brand-version">v$versionDisplay</span>
    </div>
    <div class="brand-meta">Generated $generatedDisplay</div>
  </header>

  <section class="hero hero--$statusKey" aria-labelledby="run-status">
    <div class="hero-main">
      <div class="status-symbol" aria-hidden="true">$statusGlyph</div>
      <div>
        <p class="eyebrow">$computerName</p>
        <h1 id="run-status">$overallStatus</h1>
        <p class="hero-summary"><strong>$totalUpdates updates</strong> $summaryVerb in <strong>$durationDisplay</strong></p>
      </div>
    </div>
    <div class="hero-side">
      <span class="run-pill">$modeLabel</span>
      <span class="exit-token">EXIT $exitCode</span>
    </div>
  </section>

  <section class="metrics" aria-label="Run metrics">
    <article class="metric metric--oem">
      <span class="metric-icon" aria-hidden="true">
        <svg viewBox="0 0 24 24" fill="none">
          <path d="M9 3h6v3h3v3h3v6h-3v3h-3v3H9v-3H6v-3H3V9h3V6h3V3Z" stroke="currentColor" stroke-width="1.6"/>
          <path d="M9 9h6v6H9z" stroke="currentColor" stroke-width="1.6"/>
        </svg>
      </span>
      <div><span class="metric-value">$oemCount</span><span class="metric-label">OEM updates</span></div>
    </article>
    <article class="metric metric--windows">
      <span class="metric-icon" aria-hidden="true">
        <svg viewBox="0 0 24 24" fill="currentColor">
          <path d="m3 5.2 7.4-1v7.1H3V5.2Zm8.5-1.1L21 2.8v8.5h-9.5V4.1ZM3 12.5h7.4v7.2L3 18.7v-6.2Zm8.5 0H21v8.6l-9.5-1.3v-7.3Z"/>
        </svg>
      </span>
      <div><span class="metric-value">$windowsCount</span><span class="metric-label">Windows updates</span></div>
    </article>
    <article class="metric metric--apps">
      <span class="metric-icon" aria-hidden="true">
        <svg viewBox="0 0 24 24" fill="none">
          <path d="m12 3 8 4.5v9L12 21l-8-4.5v-9L12 3Z" stroke="currentColor" stroke-width="1.7"/>
          <path d="m4.4 7.7 7.6 4.4 7.6-4.4M12 12.1V21" stroke="currentColor" stroke-width="1.5"/>
        </svg>
      </span>
      <div><span class="metric-value">$wingetCount</span><span class="metric-label">App packages</span></div>
    </article>
    <article class="metric $rebootClass">
      <span class="metric-icon" aria-hidden="true">
        <svg viewBox="0 0 24 24" fill="none">
          <path d="M18.4 8.1A8 8 0 1 1 12 4" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/>
          <path d="M12 1.8 15.1 5 12 8.2" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/>
        </svg>
      </span>
      <div><span class="metric-value metric-value--text">$rebootLabel</span><span class="metric-label">Reboot</span></div>
    </article>
  </section>

  <div class="dashboard-grid">
    <section class="panel channels" aria-labelledby="channels-heading">
      <div class="panel-heading">
        <div>
          <h2 id="channels-heading">Update channels</h2>
          <p>Work completed across the device update stack</p>
        </div>
        <span class="section-count">$totalUpdates total</span>
      </div>
      <div class="channel-list">
        <article class="channel channel--oem">
          <span class="channel-icon" aria-hidden="true">
            <svg viewBox="0 0 24 24" fill="none">
              <path d="M8 3h8v3h3v3h2v6h-2v3h-3v3H8v-3H5v-3H3V9h2V6h3V3Z" stroke="currentColor" stroke-width="1.5"/>
              <path d="M9 9h6v6H9z" stroke="currentColor" stroke-width="1.5"/>
            </svg>
          </span>
          <div><span class="channel-title">OEM &amp; firmware</span><span class="channel-caption">Vendor drivers and BIOS</span></div>
          <div class="channel-count">$oemCount<small>Updates</small></div>
          <span class="channel-state $($oemState.Class)">$($oemState.Label)</span>
          <div>$oemDetails</div>
        </article>

        <article class="channel channel--windows">
          <span class="channel-icon" aria-hidden="true">
            <svg viewBox="0 0 24 24" fill="currentColor">
              <path d="m3 5.2 7.4-1v7.1H3V5.2Zm8.5-1.1L21 2.8v8.5h-9.5V4.1ZM3 12.5h7.4v7.2L3 18.7v-6.2Zm8.5 0H21v8.6l-9.5-1.3v-7.3Z"/>
            </svg>
          </span>
          <div><span class="channel-title">Windows Update</span><span class="channel-caption">OS and Microsoft updates</span></div>
          <div class="channel-count">$windowsCount<small>Updates</small></div>
          <span class="channel-state $($windowsState.Class)">$($windowsState.Label)</span>
          <div>$windowsDetails</div>
        </article>

        <article class="channel channel--apps">
          <span class="channel-icon" aria-hidden="true">
            <svg viewBox="0 0 24 24" fill="none">
              <path d="m12 3 8 4.5v9L12 21l-8-4.5v-9L12 3Z" stroke="currentColor" stroke-width="1.6"/>
              <path d="m4.4 7.7 7.6 4.4 7.6-4.4M12 12.1V21" stroke="currentColor" stroke-width="1.4"/>
            </svg>
          </span>
          <div><span class="channel-title">Application packages</span><span class="channel-caption">winget package upgrades</span></div>
          <div class="channel-count">$wingetCount<small>Updates</small></div>
          <span class="channel-state $($wingetState.Class)">$($wingetState.Label)</span>
          <div>$wingetDetails</div>
        </article>
      </div>
    </section>

    <aside class="panel device-profile" aria-labelledby="device-heading">
      <div class="panel-heading">
        <div>
          <h2 id="device-heading">Device profile</h2>
          <p>Inventory captured at run time</p>
        </div>
      </div>
      <dl class="profile-list">
        <div class="profile-row"><dt>Device name</dt><dd>$computerName</dd></div>
        <div class="profile-row"><dt>Manufacturer</dt><dd>$manufacturer</dd></div>
        <div class="profile-row"><dt>Model</dt><dd>$model</dd></div>
        <div class="profile-row"><dt>Serial number</dt><dd>$serialNumber</dd></div>
        <div class="profile-row"><dt>Operating system</dt><dd>$osName</dd></div>
        <div class="profile-row"><dt>OS build</dt><dd>$osBuild</dd></div>
        <div class="profile-row"><dt>BIOS version</dt><dd>$biosVersion</dd></div>
        <div class="profile-row"><dt>BIOS date</dt><dd>$biosDateDisplay</dd></div>
        <div class="profile-row"><dt>Processor</dt><dd>$processor</dd></div>
        <div class="profile-row"><dt>Memory</dt><dd>$totalRam</dd></div>
      </dl>
    </aside>

    <section class="panel attention" aria-labelledby="attention-heading">
      <div class="panel-heading">
        <div>
          <h2 id="attention-heading">Attention needed</h2>
          <p>Exceptions and follow-up from this run</p>
        </div>
        <span class="section-count">$attentionSummary</span>
      </div>
      <ul class="notice-list">$noticeRows</ul>
    </section>

    <section class="panel run-details" aria-labelledby="details-heading">
      <div class="panel-heading">
        <div>
          <h2 id="details-heading">Run details</h2>
          <p>Execution metadata for audit and support</p>
        </div>
      </div>
      <dl class="detail-grid">
        <div class="detail-item"><dt>Mode</dt><dd>$(if ($DryRun) { "Dry run · no changes" } else { "Live update" })</dd></div>
        <div class="detail-item"><dt>Exit code</dt><dd>$exitCode</dd></div>
        <div class="detail-item"><dt>Duration</dt><dd>$durationDisplay · $durationSeconds seconds</dd></div>
        <div class="detail-item"><dt>Reboot</dt><dd>$rebootLabel</dd></div>
        <div class="detail-item"><dt>Component rollback</dt><dd>$componentRollbackDisplay</dd></div>
      </dl>
      <div class="log-path"><span>Log file</span><code>$logFileDisplay</code></div>
    </section>
  </div>

  <footer class="footer">
    <span><strong>SystemUpdatePro</strong> v$versionDisplay</span>
    <span>$computerName · Report generated $generatedDisplay</span>
  </footer>
</main>
</body>
</html>
"@

    try {
        $html | Set-Content -Path $reportFile -Encoding UTF8 -Force
        Write-Log "HTML report: $reportFile" "SUCCESS"

        # Auto-open in browser unless headless/unattended
        $isHeadless = [Environment]::UserInteractive -eq $false
        $isSystem = ([Security.Principal.WindowsIdentity]::GetCurrent().Name -match 'SYSTEM$')

        if (-not $isHeadless -and -not $isSystem) {
            try {
                Start-Process $reportFile -ErrorAction SilentlyContinue
            } catch {
                Write-Log "Could not auto-open HTML report: $($_.Exception.Message)" "DEBUG"
            }
        }
    } catch {
        Write-Log "Failed to write HTML report: $($_.Exception.Message)" "WARNING"
    }

    return $reportFile
}

# ============================================================================
# WEBHOOK NOTIFICATIONS
# ============================================================================

function Send-WebhookNotification {
    param(
        [string]$Url,
        [hashtable]$RunData
    )

    if (-not $Url) { return $false }

    Write-Log "Sending webhook notification..." "DEBUG"

    $overallStatus = switch ($RunData.ExitCode) {
        0 { "success" }
        1 { "success" }
        2 { "partial" }
        default { "failed" }
    }

    # Generic payload
    $payload = @{
        schema_version  = $RunData.SchemaVersion
        run_id          = $RunData.RunId
        started_at      = $RunData.StartedAt
        completed_at    = $RunData.CompletedAt
        hostname        = $env:COMPUTERNAME
        status          = $overallStatus
        oem_updates     = $RunData.OEMUpdates
        windows_updates = $RunData.WindowsUpdates
        winget_updates  = $RunData.WingetUpdates
        total_installed = $RunData.TotalInstalled
        total_available = $RunData.TotalAvailable
        total_failed    = $RunData.TotalFailed
        reboot_required = $RunData.RebootRequired
        exit_code       = $RunData.ExitCode
        errors          = @($RunData.Errors)
        warnings        = @($RunData.Warnings)
        stages          = @($RunData.Stages)
        runtime_seconds = $RunData.DurationSeconds
    }

    try {
        # Detect webhook type and format accordingly
        if ($Url -match 'hooks\.slack\.com') {
            # Slack format
            $statusIcon = switch ($overallStatus) {
                "success" { "OK" }
                "partial" { "WARN" }
                "failed"  { "FAIL" }
            }
            $dryLabel = if ($DryRun) { " [DRY RUN]" } else { "" }
            $slackPayload = @{
                text = "SystemUpdatePro$dryLabel - $($env:COMPUTERNAME)`nStatus: $statusIcon $overallStatus | OEM: $($RunData.OEMUpdates) | WinUpd: $($RunData.WindowsUpdates) | Winget: $($RunData.WingetUpdates) | Runtime: $($RunData.DurationSeconds)s"
            }
            $body = $slackPayload | ConvertTo-Json -Depth 3
        }
        elseif ($Url -match 'webhook\.office\.com' -or $Url -match 'workflows.*\.logic\.azure\.com') {
            # Microsoft Teams format
            $dryLabel = if ($DryRun) { " [DRY RUN]" } else { "" }
            $teamsPayload = @{
                "@type"    = "MessageCard"
                "@context" = "http://schema.org/extensions"
                summary    = "SystemUpdatePro Report"
                title      = "SystemUpdatePro$dryLabel - $($env:COMPUTERNAME)"
                themeColor = switch ($overallStatus) { "success" { "00FF00" }; "partial" { "FFFF00" }; "failed" { "FF0000" } }
                sections   = @(
                    @{
                        facts = @(
                            @{ name = "Status"; value = $overallStatus.ToUpper() },
                            @{ name = "OEM Updates"; value = "$($RunData.OEMUpdates)" },
                            @{ name = "Windows Updates"; value = "$($RunData.WindowsUpdates)" },
                            @{ name = "Winget Updates"; value = "$($RunData.WingetUpdates)" },
                            @{ name = "Runtime"; value = "$($RunData.DurationSeconds) seconds" },
                            @{ name = "Errors"; value = "$($RunData.Errors.Count)" }
                        )
                    }
                )
            }
            $body = $teamsPayload | ConvertTo-Json -Depth 5
        }
        else {
            # Generic webhook
            $body = $payload | ConvertTo-Json -Depth 12
        }

        Invoke-RestMethod -Uri $Url -Method Post -Body $body -ContentType "application/json" -TimeoutSec 30 -ErrorAction Stop | Out-Null
        Write-Log "Webhook notification sent" "SUCCESS"
        return $true
    }
    catch {
        Write-Log "Webhook notification failed: $($_.Exception.Message)" "WARNING"
        return $false
    }
}

function Invoke-TerminalEvidence {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$SysInfo,
        [Parameter(Mandatory = $true)]
        [hashtable]$RunData,
        [string]$WebhookEndpoint = ""
    )

    if ($script:RunFinalized) {
        if ($script:LastEvidenceDelivery) {
            $RunData.EvidenceDelivery = $script:LastEvidenceDelivery
        }
        return $RunData
    }

    # Mark first so a writer failure or nested error can never duplicate terminal
    # history, webhook, report, or Event Log evidence.
    $script:RunFinalized = $true
    $delivery = [ordered]@{
        Report  = [ordered]@{ Attempted = $true; Status = "Attempted"; Detail = "" }
        Event   = [ordered]@{ Attempted = $true; Status = "Attempted"; Detail = "" }
        Webhook = [ordered]@{ Attempted = [bool]$WebhookEndpoint; Status = $(if ($WebhookEndpoint) { "Attempted" } else { "Skipped" }); Detail = "" }
        History = [ordered]@{ Attempted = $true; Status = "Attempted"; Detail = "" }
    }
    $RunData.EvidenceDelivery = $delivery

    try {
        $reportPath = New-HTMLReport -SysInfo $SysInfo -RunData $RunData
        if ($reportPath -and (Test-Path -LiteralPath $reportPath)) {
            $delivery.Report.Status = "Succeeded"
            $delivery.Report.Detail = $reportPath
        } else {
            $delivery.Report.Status = "Failed"
            $delivery.Report.Detail = "Report file was not created"
        }
    } catch {
        $delivery.Report.Status = "Failed"
        $delivery.Report.Detail = $_.Exception.Message
    }

    try {
        [void](Initialize-EventLog)
        $eventMessage = @"
SystemUpdatePro completed$(if ($DryRun) { ' (DRY RUN)' })
Run ID: $($RunData.RunId)
Status: $($RunData.Status)
System: $($SysInfo.Manufacturer) $($SysInfo.Model)
Updates $(if ($DryRun) { 'Available' } else { 'Applied' }): $(if ($DryRun) { $RunData.TotalAvailable } else { $RunData.TotalInstalled })
Updates Failed: $($RunData.TotalFailed)
Reboot Required: $($RunData.RebootRequired)
Duration: $($RunData.DurationSeconds) seconds
"@
        $eventType = if ($RunData.ExitCode -eq 0) { "Information" } elseif ($RunData.ExitCode -le 2) { "Warning" } else { "Error" }
        if (Write-EventLogEntry -Message $eventMessage -EntryType $eventType -EventId (1000 + $RunData.ExitCode)) {
            $delivery.Event.Status = "Succeeded"
        } else {
            $delivery.Event.Status = "Failed"
            $delivery.Event.Detail = "Windows Event Log write failed"
        }
    } catch {
        $delivery.Event.Status = "Failed"
        $delivery.Event.Detail = $_.Exception.Message
    }

    if ($WebhookEndpoint) {
        try {
            if (Send-WebhookNotification -Url $WebhookEndpoint -RunData $RunData) {
                $delivery.Webhook.Status = "Succeeded"
            } else {
                $delivery.Webhook.Status = "Failed"
                $delivery.Webhook.Detail = "Webhook delivery failed"
            }
        } catch {
            $delivery.Webhook.Status = "Failed"
            $delivery.Webhook.Detail = $_.Exception.Message
        }
    }

    try {
        # A persisted record can truthfully claim success for its own write; if
        # the write fails there is no record, and the in-memory status is reset.
        $delivery.History.Status = "Succeeded"
        if (Save-UpdateHistory -RunData $RunData) {
            $delivery.History.Status = "Succeeded"
        } else {
            $delivery.History.Status = "Failed"
            $delivery.History.Detail = "History write failed"
        }
    } catch {
        $delivery.History.Status = "Failed"
        $delivery.History.Detail = $_.Exception.Message
    }

    $RunData.EvidenceDelivery = $delivery
    $script:LastEvidenceDelivery = $delivery
    return $RunData
}

# ============================================================================
# MAIN EXECUTION
# ============================================================================

# Handle -ShowHistory early exit
if ($ShowHistory) {
    Show-UpdateHistory -Count $HistoryCount
    exit 0
}

$scriptStart = $script:RunStartedAt
$sysInfo = @{
    Manufacturer = "Unknown"
    Model = "Unknown"
    SerialNumber = ""
    BIOSVersion = ""
    BIOSDate = $null
    OSName = ""
    OSBuild = ""
    Processor = ""
    TotalRAM = 0
}
$lockAcquired = $false
$wsusBypassApplied = $false
$shutdownRequested = $false
$runData = $null
$state = @{}

try {
    :run do {
        $initializationStart = Get-Date
        if (-not (Test-Administrator)) {
            $message = "This script requires administrator privileges"
            Write-Host "[X] $message." -ForegroundColor Red
            [void]$script:Errors.Add($message)
            [void](Add-StageResult (New-StageResult -Name "Initialization" -Status "Failed" `
                -Attempted 1 -Failed 1 -ProviderCode 3 -Message $message -StartedAt $initializationStart))
            $script:ExitCode = 3
            break run
        }

        $state = Get-State
        if ($state.Count -gt 0) {
            $resumeResult = Import-ContinuationState -State $state
            if (-not $resumeResult.Success) {
                $message = "Continuation state could not be resumed: $($resumeResult.Message)"
                [void]$script:Errors.Add($message)
                [void](Add-StageResult (New-StageResult -Name "Resume" -Provider "Task Scheduler" `
                    -Status "Failed" -Attempted 1 -Failed 1 -ProviderCode 3 -Message $message `
                    -StartedAt $initializationStart))
                $script:ExitCode = 3
                break run
            }
            $scriptStart = $script:RunStartedAt
        }

        if (-not (Initialize-RunEnvironment)) {
            $message = "Run storage initialization failed"
            [void]$script:Errors.Add($message)
            [void](Add-StageResult (New-StageResult -Name "Initialization" -Status "Failed" `
                -Attempted 1 -Failed 1 -ProviderCode 3 -Message $message -StartedAt $initializationStart))
            $script:ExitCode = 3
            break run
        }

        [void](Add-StageResult (New-StageResult -Name "Initialization" -Status "Succeeded" `
            -Attempted 1 -Message "Run environment initialized" -StartedAt $initializationStart `
            -DurationSeconds ([int]((Get-Date) - $initializationStart).TotalSeconds)))

        Write-Banner
        [void](Initialize-EventLog)

        if ($DryRun) {
            Write-Log "DRY RUN MODE - No updates will be installed" "STEP"
            Write-Host ""
        }

        $coordinationStart = Get-Date
        if (Test-LockFile) {
            $message = "Another instance is already running"
            Write-Log $message "ERROR"
            [void](Add-StageResult (New-StageResult -Name "Coordination" -Status "Failed" `
                -Attempted 1 -Failed 1 -ProviderCode 6 -Message $message -StartedAt $coordinationStart))
            $script:ExitCode = 6
            break run
        }

        New-LockFile
        $lockAcquired = $true
        [void](Add-StageResult (New-StageResult -Name "Coordination" -Status "Succeeded" `
            -Attempted 1 -Message "Single-instance lock acquired" -StartedAt $coordinationStart))

        Invoke-LogRotation -RetentionDays $LogRetentionDays
        Write-Log "Run ID: $($script:RunId)" "DEBUG"
        Write-Log "Log: $script:LogFile" "DEBUG"

        if ($script:ContinuationActive) {
            $continuationStart = Get-Date
            Write-Log "Resuming run $($script:RunId) after reboot at stage '$($script:ResumeStageCursor)' (attempt $($script:ContinuationAttempt)/$($script:MaxContinuationAttempts))..." "STEP"
            [void](Add-StageResult (New-StageResult -Name "Resume" -Status "Succeeded" `
                -Attempted 1 -Message "Post-reboot continuation resumed at $($script:ResumeStageCursor)" `
                -ProviderCode $script:ContinuationAttempt -StartedAt $continuationStart))
        }

        Write-Log "Running pre-flight checks..." "STEP"
        $preflightStart = Get-Date
        $preflightItems = [System.Collections.ArrayList]::new()

        $sysInfo = Get-SystemInfo
        Write-Log "System: $($sysInfo.Manufacturer) $($sysInfo.Model)" "INFO"
        Write-Log "OS: $($sysInfo.OSName) (Build $($sysInfo.OSBuild))" "DEBUG"

        if (Test-InternetConnection) {
            [void]$preflightItems.Add((New-UpdateItemResult -Name "Internet connectivity" -Status "Succeeded"))
        } else {
            Write-Log "No internet connection" "WARNING"
            [void]$preflightItems.Add((New-UpdateItemResult -Name "Internet connectivity" -Status "Warning" -Message "No internet connection detected"))
        }

        $disk = Test-DiskSpace -MinGB $MinDiskSpaceGB
        Write-Log "Free disk space: $($disk.FreeGB) GB (required: $($disk.RequiredGB) GB)" "DEBUG"
        if (-not $disk.Sufficient) {
            $message = "Insufficient disk space"
            if (-not $Force) {
                Write-Log $message "ERROR"
                [void]$preflightItems.Add((New-UpdateItemResult -Name "Disk space" -Status "Blocked" -ProviderCode 4 -Message $message))
                [void](Add-StageResult (New-StageResult -Name "Preflight" -Status "Failed" `
                    -Attempted $preflightItems.Count -Failed 1 -ProviderCode 4 -Message $message `
                    -Items @($preflightItems) -StartedAt $preflightStart))
                $script:ExitCode = 4
                break run
            }
            Write-Log "$message overridden by -Force" "WARNING"
            [void]$preflightItems.Add((New-UpdateItemResult -Name "Disk space" -Status "Warning" -Message "$message overridden"))
        } else {
            [void]$preflightItems.Add((New-UpdateItemResult -Name "Disk space" -Status "Succeeded" -Message "$($disk.FreeGB) GB free"))
        }

        $reboot = Test-PendingReboot
        if ($reboot.Pending) {
            $message = "Pending reboot: $($reboot.Reasons -join ', ')"
            Write-Log $message "WARNING"
            if (-not $Force) {
                Write-Log "Use -Force to override or reboot first" "ERROR"
                [void]$preflightItems.Add((New-UpdateItemResult -Name "Pending reboot" -Status "Blocked" -ProviderCode 5 -Message $message))
                [void](Add-StageResult (New-StageResult -Name "Preflight" -Status "Failed" `
                    -Attempted $preflightItems.Count -Failed 1 -ProviderCode 5 -Message $message `
                    -Items @($preflightItems) -StartedAt $preflightStart))
                $script:ExitCode = 5
                break run
            }
            [void]$preflightItems.Add((New-UpdateItemResult -Name "Pending reboot" -Status "Warning" -Message "$message; overridden"))
        } else {
            [void]$preflightItems.Add((New-UpdateItemResult -Name "Pending reboot" -Status "Succeeded"))
        }

        if ($IncludeBIOS) {
            $battery = Test-BatteryPower
            if ($battery.OnBattery) {
                $message = "Cannot update BIOS on battery power"
                if (-not $Force) {
                    Write-Log $message "ERROR"
                    [void]$preflightItems.Add((New-UpdateItemResult -Name "Firmware power" -Status "Blocked" -ProviderCode 7 -Message $message))
                    [void](Add-StageResult (New-StageResult -Name "Preflight" -Status "Failed" `
                        -Attempted $preflightItems.Count -Failed 1 -ProviderCode 7 -Message $message `
                        -Items @($preflightItems) -StartedAt $preflightStart))
                    $script:ExitCode = 7
                    break run
                }
                Write-Log "$message overridden by -Force" "WARNING"
                [void]$preflightItems.Add((New-UpdateItemResult -Name "Firmware power" -Status "Warning" -Message "$message overridden"))
            } else {
                [void]$preflightItems.Add((New-UpdateItemResult -Name "Firmware power" -Status "Succeeded" -Message "AC power detected"))
            }
        } else {
            [void]$preflightItems.Add((New-UpdateItemResult -Name "Firmware power" -Status "Skipped" -Message "BIOS updates not requested"))
        }

        if (Test-MeteredConnection) {
            Write-Log "Metered connection detected - large downloads may incur charges" "WARNING"
            [void]$preflightItems.Add((New-UpdateItemResult -Name "Network cost" -Status "Warning" -Message "Metered connection detected"))
        } else {
            [void]$preflightItems.Add((New-UpdateItemResult -Name "Network cost" -Status "Succeeded"))
        }

        $bitlocker = Test-BitLockerEnabled
        $bitlockerStatus = if ($bitlocker.Status -eq "Unknown") { "Warning" } else { "Succeeded" }
        [void]$preflightItems.Add((New-UpdateItemResult -Name "BitLocker state" -Status $bitlockerStatus -Message $bitlocker.Status))
        if ($bitlocker.Enabled) { Write-Log "BitLocker: Active" "INFO" }

        [void](Add-StageResult (New-StageResult -Name "Preflight" -Status "Succeeded" `
            -Attempted $preflightItems.Count -Message "Preflight checks completed" -Items @($preflightItems) `
            -StartedAt $preflightStart -DurationSeconds ([int]((Get-Date) - $preflightStart).TotalSeconds)))
        Write-Host ""

        if (-not $script:ContinuationActive) {
            $repairStart = Get-Date
            if ($RepairWindowsUpdate) {
            $repairSucceeded = [bool](Repair-WindowsUpdateServices)
            $repairStage = ConvertTo-StageResult -Name "WindowsUpdateRepair" -Provider "Windows servicing" `
                -Result @{ Success = $repairSucceeded; Message = $(if ($repairSucceeded) { "Repair completed" } else { "Repair failed" }) } `
                -StartedAt $repairStart
            [void](Add-StageResult $repairStage)
            if ($repairStage.Status -eq "Failed") { $script:ExitCode = 2 }
            Write-Host ""
            } else {
                [void](Add-StageResult (New-StageResult -Name "WindowsUpdateRepair" -Provider "Windows servicing" `
                    -Status "Skipped" -Skipped 1 -Message "Repair not requested" -StartedAt $repairStart))
            }

            $wsusStart = Get-Date
            if ($BypassWSUS) {
                Set-WSUSBypass -Enable
                $wsusBypassApplied = $true
                [void](Add-StageResult (New-StageResult -Name "WSUSBypass" -Provider "Windows Update policy" `
                    -Status "Succeeded" -Attempted 1 -Message "Temporary WSUS bypass applied" -StartedAt $wsusStart))
            } else {
                [void](Add-StageResult (New-StageResult -Name "WSUSBypass" -Provider "Windows Update policy" `
                    -Status "Skipped" -Skipped 1 -Message "WSUS bypass not requested" -StartedAt $wsusStart))
            }

            $backupStart = Get-Date
            if ($BackupDrivers -and -not $SkipOEM) {
                $backupPath = Invoke-DriverBackup
                $backupSuccess = [bool]$backupPath
                $backupStage = ConvertTo-StageResult -Name "DriverBackup" -Provider "Export-WindowsDriver" `
                    -Result @{ Success = $backupSuccess; Message = $(if ($backupSuccess) { "Driver backup completed" } else { "Driver backup failed" }); Evidence = @($backupPath) } `
                    -StartedAt $backupStart
                [void](Add-StageResult $backupStage)
                if ($backupStage.Status -eq "Failed") { $script:ExitCode = 2 }
                Write-Host ""
            } else {
                [void](Add-StageResult (New-StageResult -Name "DriverBackup" -Provider "Export-WindowsDriver" `
                    -Status "Skipped" -Skipped 1 -Message "Driver backup not requested" -StartedAt $backupStart))
            }

            $oemStart = Get-Date
            if ($SkipOEM) {
                [void](Add-StageResult (New-StageResult -Name "OEM" -Provider "OEM" -Status "Skipped" `
                    -Skipped 1 -Message "OEM updates skipped by run configuration" -StartedAt $oemStart))
            } else {
                $manufacturer = [string]$sysInfo.Manufacturer
                $provider = $manufacturer
                $oemResult = $null

                if ($manufacturer -match "DELL|ALIENWARE") {
                    $provider = "Dell Command Update"
                    $oemResult = Invoke-DellUpdate -IncludeBIOS:$IncludeBIOS
                } elseif ($manufacturer -match "LENOVO") {
                    $provider = "LSUClient"
                    $oemResult = Invoke-LenovoUpdate -IncludeBIOS:$IncludeBIOS
                } elseif ($manufacturer -match "HP|HEWLETT") {
                    $provider = "HP Image Assistant"
                    $oemResult = Invoke-HPUpdate -IncludeBIOS:$IncludeBIOS
                } else {
                    Write-Log "========== OEM UPDATES ==========" "HEADER"
                    Write-Log "Manufacturer '$manufacturer' not supported" "INFO"
                    [void](Add-StageResult (New-StageResult -Name "OEM" -Provider $manufacturer -Status "Skipped" `
                        -Skipped 1 -Message "Manufacturer is not supported" -StartedAt $oemStart))
                }

                if ($oemResult) {
                    $oemStage = ConvertTo-StageResult -Name "OEM" -Provider $provider -Result $oemResult `
                        -ItemNames @($script:OEMUpdates) -StartedAt $oemStart
                    [void](Add-StageResult $oemStage)
                    if ($oemStage.Status -in @("Failed", "Partial")) { $script:ExitCode = 2 }
                }
                Write-Host ""
            }
        } else {
            Write-Log "Stages before '$($script:ResumeStageCursor)' were restored from continuation state" "INFO"
        }

        if (Test-ShouldRunContinuationStage -Stage "WindowsUpdate") {
            $windowsStart = Get-Date
            if ($SkipWindows) {
                [void](Add-StageResult (New-StageResult -Name "WindowsUpdate" -Provider "Windows Update" `
                    -Status "Skipped" -Skipped 1 -Message "Windows Update skipped by run configuration" -StartedAt $windowsStart))
            } else {
                $wuResult = Invoke-WindowsUpdate -MaxPasses $MaxUpdatePasses
                $wuStage = ConvertTo-StageResult -Name "WindowsUpdate" -Provider "Windows Update" `
                    -Result $wuResult -ItemNames @($script:WindowsUpdates) -StartedAt $windowsStart
                [void](Add-StageResult $wuStage)
                if ($wuStage.Status -in @("Failed", "Partial")) { $script:ExitCode = 2 }
                Write-Host ""
            }
            if (-not (Set-ContinuationCursor -StageCursor "Winget")) {
                throw "Failed to persist continuation cursor after Windows Update"
            }
        } else {
            Write-Log "Windows Update stage already completed before continuation resumed" "INFO"
        }

        if (Test-ShouldRunContinuationStage -Stage "Winget") {
            $wingetStart = Get-Date
            if ($SkipWinget) {
                [void](Add-StageResult (New-StageResult -Name "Winget" -Provider "WinGet" -Status "Skipped" `
                    -Skipped 1 -Message "WinGet skipped by run configuration" -StartedAt $wingetStart))
            } else {
                $wingetResult = Invoke-WingetUpgradeAll
                $wingetStage = ConvertTo-StageResult -Name "Winget" -Provider "WinGet" `
                    -Result $wingetResult -ItemNames @($script:WingetUpdates) -StartedAt $wingetStart
                [void](Add-StageResult $wingetStage)
                if ($wingetStage.Status -in @("Failed", "Partial")) { $script:ExitCode = 2 }
                Write-Host ""
            }
            if (-not (Set-ContinuationCursor -StageCursor "Cleanup")) {
                throw "Failed to persist continuation cursor after WinGet"
            }
        } else {
            Write-Log "WinGet stage already completed before continuation resumed" "INFO"
        }

        if (Test-ShouldRunContinuationStage -Stage "Cleanup") {
            $cleanupStart = Get-Date
            if ($CleanupAfter -or $ResetComponentBase) {
                $cleanupResult = Invoke-ComponentCleanup
                $cleanupStage = ConvertTo-StageResult -Name "Cleanup" -Provider "DISM and cleanmgr" `
                    -Result $cleanupResult `
                    -StartedAt $cleanupStart
                [void](Add-StageResult $cleanupStage)
                if ($cleanupStage.Status -in @("Failed", "Partial")) { $script:ExitCode = 2 }
                Write-Host ""
            } else {
                [void](Add-StageResult (New-StageResult -Name "Cleanup" -Provider "DISM and cleanmgr" `
                    -Status "Skipped" -Skipped 1 -Message "Cleanup not requested" -StartedAt $cleanupStart))
            }
            if (-not (Set-ContinuationCursor -StageCursor "Complete")) {
                throw "Failed to persist completed continuation cursor"
            }
        } else {
            Write-Log "Cleanup stage already completed before continuation resumed" "INFO"
        }

        $script:UpdatesInstalled = [int](($script:StageResults | Measure-Object -Property Installed -Sum).Sum)
        $script:UpdatesFailed = [int](($script:StageResults | Measure-Object -Property Failed -Sum).Sum)

        if ($script:RebootRequired -and -not $DryRun -and $ContinueAfterReboot) {
            $continuationStart = Get-Date
            $continuationSucceeded = [bool](Register-ContinuationTask)
            $continuationStage = ConvertTo-StageResult -Name "Continuation" -Provider "Task Scheduler" `
                -Result @{ Success = $continuationSucceeded; Message = $(if ($continuationSucceeded) { "Continuation task registered" } else { "Continuation task registration failed" }) } `
                -StartedAt $continuationStart
            [void](Add-StageResult $continuationStage)
            if ($continuationStage.Status -eq "Failed") { $script:ExitCode = 2 }
        }

        if ($script:RebootRequired -and -not $DryRun -and $Reboot) {
            if (-not $ContinueAfterReboot -or $script:ContinuationRegistered) {
                $shutdownRequested = $true
            } else {
                Write-Log "Automatic reboot cancelled because continuation was not registered" "ERROR"
            }
        }
    } while ($false)
} catch {
    $message = "Unhandled runtime error: $($_.Exception.Message)"
    try { Write-Log $message "ERROR" } catch { Write-Host "[X] $message" -ForegroundColor Red }
    if (-not ($script:Errors -contains $message)) { [void]$script:Errors.Add($message) }
    [void](Add-StageResult (New-StageResult -Name "Runtime" -Status "Failed" -Attempted 1 `
        -Failed 1 -ProviderCode 3 -HResult $_.Exception.HResult -Message $message -StartedAt (Get-Date)))
    $script:ExitCode = 3
} finally {
    if ($wsusBypassApplied) {
        try {
            Set-WSUSBypass
        } catch {
            $message = "WSUS settings restoration failed: $($_.Exception.Message)"
            try { Write-Log $message "ERROR" } catch {}
            if (-not ($script:Errors -contains $message)) { [void]$script:Errors.Add($message) }
            $script:ExitCode = 3
        }
    }

    if (-not $DryRun -and -not $script:ContinuationRegistered) {
        $continuationArtifactsPresent = $script:ContinuationActive -or
            (Test-Path -LiteralPath $script:StateFile) -or (Test-ContinuationTask)
        $continuationCleanupStart = Get-Date
        $continuationCleanupSucceeded = [bool](Unregister-ContinuationTask)
        if ($continuationArtifactsPresent) {
            $continuationCleanupStage = ConvertTo-StageResult -Name "ContinuationCleanup" -Provider "Task Scheduler" `
                -Result @{
                    Success = $continuationCleanupSucceeded
                    Message = $(if ($continuationCleanupSucceeded) {
                        "One-shot continuation task and state removed"
                    } else {
                        "Continuation task or state cleanup failed"
                    })
                } -StartedAt $continuationCleanupStart
            [void](Add-StageResult $continuationCleanupStage)
            if (-not $continuationCleanupSucceeded) { $script:ExitCode = 3 }
        }
    }

    if ($lockAcquired) {
        Remove-LockFile
    }

    $completedAt = Get-Date
    $runData = New-RunData -StartedAt $scriptStart -CompletedAt $completedAt -RequestedExitCode $script:ExitCode
    $script:ExitCode = $runData.ExitCode
    $script:RebootRequired = $runData.RebootRequired
    $script:UpdatesInstalled = $runData.TotalInstalled
    $script:UpdatesFailed = $runData.TotalFailed
    $script:OEMUpdateCount = $runData.OEMUpdates
    $script:WindowsUpdateCount = $runData.WindowsUpdates
    $script:WingetUpdateCount = $runData.WingetUpdates

    try {
        Write-Log "================================================================" "HEADER"
        $summaryLabel = if ($DryRun) { "  DRY RUN COMPLETE" } else { "  UPDATE COMPLETE" }
        Write-Log $summaryLabel "HEADER"
        Write-Log "================================================================" "HEADER"
        Write-Host ""
        Write-Log "Run ID:          $($runData.RunId)" "INFO"
        Write-Log "System:          $($sysInfo.Manufacturer) $($sysInfo.Model)" "INFO"
        Write-Log "Duration:        $([math]::Round($runData.DurationSeconds / 60, 1)) minutes" "INFO"
        $updateCount = if ($DryRun) { $runData.TotalAvailable } else { $runData.TotalInstalled }
        $updateLabel = if ($DryRun) { "Updates Available" } else { "Updates Applied" }
        Write-Log "$updateLabel`:  $updateCount" "INFO"
        if ($runData.TotalFailed -gt 0) { Write-Log "Updates Failed:  $($runData.TotalFailed)" "WARNING" }
        Write-Log "Reboot Required: $($runData.RebootRequired)" $(if ($runData.RebootRequired) { "WARNING" } else { "INFO" })
        Write-Log "Exit Code:       $($runData.ExitCode)" "DEBUG"
        Write-Log "Log File:        $script:LogFile" "DEBUG"
    } catch {}

    try {
        $runData = Invoke-TerminalEvidence -SysInfo $sysInfo -RunData $runData -WebhookEndpoint $WebhookUrl
    } catch {
        [Console]::Error.WriteLine("[X] Terminal evidence finalization failed: $($_.Exception.Message)")
        if ($script:ExitCode -lt 3) { $script:ExitCode = 3 }
    }

    if ($script:TranscriptStarted) {
        try { Stop-Transcript | Out-Null } catch {}
        $script:TranscriptStarted = $false
    }
}

if ($shutdownRequested) {
    Write-Log "Rebooting in 30 seconds..." "WARNING"
    shutdown.exe /r /t 30 /c "SystemUpdatePro - Reboot Required"
}

exit $script:ExitCode
