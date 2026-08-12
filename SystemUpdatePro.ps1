<#
.SYNOPSIS
    SystemUpdatePro v4.2.0 - Enterprise Multi-OEM System Update Utility
.DESCRIPTION
    Bulletproof MSP-grade unattended update tool with self-healing capabilities.

    FEATURES:
    - Multi-OEM Support: Dell, Lenovo, HP (auto-detects manufacturer)
    - Windows Update with automatic service repair
    - Winget upgrade all (auto-installs winget on Windows 10)
    - Self-healing: repairs corrupted Windows Update components
    - Fail-closed firmware gating for OEM/model, tool, disk, AC/charge, and BitLocker readiness
    - Disk space verification before updates
    - Post-reboot continuation via scheduled task
    - Event Log integration for RMM visibility
    - Concurrent execution prevention (lock file)
    - Automatic log rotation
    - WSUS bypass option for direct Microsoft updates
    - Comprehensive retry logic with exponential backoff
    - DryRun mode for safe preview of available updates
    - HTML summary report generation
    - Versioned, retryable webhook notifications (Slack, Teams Workflows, generic)
    - Driver backup before OEM updates
    - Automatic throttled System Restore point creation and DISM driver rollback
    - Update history tracking with JSON log
    - Bounded, redacted diagnostic and recovery archives

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
    Include BIOS and firmware updates only after every safety prerequisite is known-ready
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
.PARAMETER RollbackDrivers
    Restore the newest protected driver backup with DISM and skip forward-update stages
.PARAMETER ShowHistory
    Display update history from previous runs
.PARAMETER WebhookSecretReference
    Environment-variable or protected JSON-file reference containing an HTTPS webhook endpoint
.PARAMETER HistoryCount
    Number of history entries to display with -ShowHistory (default: 10)
.PARAMETER MaxRetries
    Maximum attempts for retryable operations and webhook delivery (default: 3)
.PARAMETER MaxUpdatePasses
    Maximum Windows Update passes (default: 3)
.PARAMETER MinDiskSpaceGB
    Minimum free disk space required in GB (default: 10)
.PARAMETER MinFirmwareChargePercent
    Minimum battery charge required for firmware updates (default: 50)
.PARAMETER LogPath
    Custom log directory (default: C:\ProgramData\SystemUpdatePro\Logs)
.PARAMETER LogRetentionDays
    Days to keep old logs (default: 30)
.PARAMETER EvidenceMaxSizeMB
    Maximum combined size of retained logs, reports, and driver backups (default: 512 MB)
.PARAMETER RedactionMode
    Redact secrets only, or secrets and device serial numbers (default: SecretsAndSerials)
.PARAMETER CreateDiagnosticBundle
    Collect a bounded, redacted diagnostic archive from the latest failed or completed run, then exit
.PARAMETER DiagnosticBundleMaxSizeMB
    Maximum diagnostic archive size (default: 50 MB)
.PARAMETER Reboot
    Allow automatic reboot if required
.PARAMETER Force
    Continue despite non-firmware warnings (pending reboot, low disk); never overrides unknown firmware safety state
.PARAMETER Offline
    Disable network acquisition and use only verified content-addressed dependency cache artifacts
.PARAMETER DependencyCachePath
    Optional administrator-prefilled content-addressed dependency cache directory
.PARAMETER SourceTimeoutSeconds
    Per-origin readiness and dependency acquisition timeout in seconds
.PARAMETER AllowMeteredNetwork
    Audited override that permits provider downloads on a known metered connection
.PARAMETER PolicyPath
    Optional protected JSON policy containing package, Windows Update, maintenance, and conflict rules
.PARAMETER RolloutPolicyPath
    Optional protected JSON policy used for deterministic cohort promotion decisions
.PARAMETER FeatureDeferralDays
    Keep feature updates deferred until their release age reaches this number of days
.PARAMETER SecurityOnly
    Only apply updates classified as critical or security updates
.PARAMETER PreStage
    Download approved Windows Updates and persist an install plan for a later run
.PARAMETER Interactive
    Ask for confirmation before requesting an automatic reboot
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
    .\SystemUpdatePro.ps1 -WebhookSecretReference "env:SYSTEMUPDATEPRO_WEBHOOK_URL"
    # Resolve the webhook from an environment variable without exposing it in argv
.EXAMPLE
    .\SystemUpdatePro.ps1 -ShowHistory -HistoryCount 20
    # Show last 20 update runs
.EXAMPLE
    .\SystemUpdatePro.ps1 -CreateDiagnosticBundle
    # Collect the latest run, provider, Windows Update, and recovery evidence, then exit
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
    Version: 4.2.0
    Requires: Administrator, PowerShell 5.1+, and a supported provider capability

    EXIT CODES:
        0 = Success, no reboot needed
        1 = Success, reboot required
        2 = Partial success (some failed)
        3 = Critical failure
        4 = Insufficient disk space
        5 = Pending reboot blocked execution
        6 = Already running (lock file exists)
        7 = Firmware safety prerequisites blocked
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
    [switch]$RollbackDrivers,
    [switch]$ShowHistory,
    [ValidateScript({
        if ([string]::IsNullOrWhiteSpace([string]$_)) { return $true }
        if ([string]$_ -match "(?i)^env:[A-Z_][A-Z0-9_]*$") { return $true }
        if ([string]$_ -match "(?i)^file:(?<ConfigPath>[A-Z]:[\\/].+)$") {
            $candidatePath = [string]$Matches.ConfigPath
            if ($candidatePath -match '[*?"<>|]' -or
                -not [IO.Path]::IsPathRooted($candidatePath)) {
                throw "Webhook config reference must use an absolute local path without wildcards"
            }
            return $true
        }
        throw "WebhookSecretReference must be empty, env:VARIABLE_NAME, or file:C:\path\config.json"
    })]
    [string]$WebhookSecretReference = "",
    [ValidateRange(1, 100)]
    [int]$HistoryCount = 10,
    [ValidateRange(1, 10)]
    [int]$MaxRetries = 3,
    [ValidateRange(1, 10)]
    [int]$MaxUpdatePasses = 3,
    [ValidateRange(1, 1024)]
    [int]$MinDiskSpaceGB = 10,
    [ValidateRange(10, 100)]
    [int]$MinFirmwareChargePercent = 50,
    [ValidateScript({
        $candidatePath = [string]$_
        if ([string]::IsNullOrWhiteSpace($candidatePath) -or
            $candidatePath -notmatch "^[A-Z]:[\\/]" -or
            $candidatePath -match '[*?"<>|]') {
            throw "LogPath must be an absolute local path without wildcards"
        }
        try {
            $fullPath = [IO.Path]::GetFullPath($candidatePath).TrimEnd("\", "/")
            $rootPath = [IO.Path]::GetPathRoot($fullPath).TrimEnd("\", "/")
            if ($fullPath -eq $rootPath) {
                throw "LogPath must be a dedicated directory, not a drive root"
            }
        } catch {
            throw "LogPath is invalid: $($_.Exception.Message)"
        }
        return $true
    })]
    [string]$LogPath = "C:\ProgramData\SystemUpdatePro\Logs",
    [ValidateRange(1, 3650)]
    [int]$LogRetentionDays = 30,
    [ValidateRange(10, 10240)]
    [int]$EvidenceMaxSizeMB = 512,
    [ValidateSet("Secrets", "SecretsAndSerials")]
    [string]$RedactionMode = "SecretsAndSerials",
    [switch]$CreateDiagnosticBundle,
    [ValidateRange(5, 512)]
    [int]$DiagnosticBundleMaxSizeMB = 50,
    [switch]$Reboot,
    [switch]$Force,
    [switch]$Offline,
    [ValidateScript({
        $candidatePath = [string]$_
        if ([string]::IsNullOrWhiteSpace($candidatePath) -or
            $candidatePath -notmatch "^[A-Z]:[\\/]" -or
            $candidatePath -match '[*?"<>|]') {
            throw "DependencyCachePath must be an absolute local path without wildcards"
        }
        return $true
    })]
    [string]$DependencyCachePath = "C:\ProgramData\SystemUpdatePro\Cache",
    [ValidateRange(1, 600)]
    [int]$SourceTimeoutSeconds = 30,
    [switch]$AllowMeteredNetwork,
    [ValidateScript({
        $candidatePath = [string]$_
        if (-not [string]::IsNullOrWhiteSpace($candidatePath) -and
            ($candidatePath -notmatch "^[A-Z]:[\\/]" -or $candidatePath -match '[*?"<>|]')) {
            throw "PolicyPath must be empty or an absolute local path without wildcards"
        }
        return $true
    })]
    [string]$PolicyPath = "",
    [ValidateScript({
        $candidatePath = [string]$_
        if (-not [string]::IsNullOrWhiteSpace($candidatePath) -and
            ($candidatePath -notmatch "^[A-Z]:[\\/]" -or $candidatePath -match '[*?"<>|]')) {
            throw "RolloutPolicyPath must be empty or an absolute local path without wildcards"
        }
        return $true
    })]
    [string]$RolloutPolicyPath = "",
    [ValidateRange(0, 3650)]
    [int]$FeatureDeferralDays = 0,
    [switch]$SecurityOnly,
    [switch]$PreStage,
    [switch]$Interactive
)

# ============================================================================
# SCRIPT CONFIGURATION
# ============================================================================

$script:Version = "4.2.0"
$script:ProductName = "SystemUpdatePro"
$script:WebhookSecretReference = [string]$WebhookSecretReference
$script:WebhookUrl = ""
$script:Offline = [bool]$Offline
$script:DependencyCachePath = [string]$DependencyCachePath
$script:SourceTimeoutSeconds = [int]$SourceTimeoutSeconds
$script:AllowMeteredNetwork = [bool]$AllowMeteredNetwork
$script:PolicyPath = [string]$PolicyPath
$script:RolloutPolicyPath = [string]$RolloutPolicyPath
$script:FeatureDeferralDays = [int]$FeatureDeferralDays
$script:SecurityOnly = [bool]$SecurityOnly
$script:PreStage = [bool]$PreStage
$script:Interactive = [bool]$Interactive
$script:RollbackDrivers = [bool]$RollbackDrivers
$script:DependencyReadiness = $null
$script:DownloadPolicy = $null
$script:CurrentSystemInfo = $null
$script:WingetScopeResults = @()
$script:RolloutDecision = $null
$script:PackagePolicy = $null
$script:RolloutPolicy = $null
$script:WindowsUpdatePolicy = $null
$script:MaintenanceDecision = $null
$script:PowerPlanState = $null
$script:PreHealthCheck = $null
$script:PostHealthCheck = $null
$script:HealthRegression = $null
$script:DryRunMutationBaseline = $null
$script:MaxRetries = [int]$MaxRetries
$script:EvidenceMaxSizeMB = [int]$EvidenceMaxSizeMB
$script:RedactionMode = [string]$RedactionMode
$script:EventLogSource = "SystemUpdatePro"
$script:ResultSchemaVersion = 1
$script:StateSchemaVersion = 8
$script:CapabilitySchemaVersion = 1
$script:HistorySchemaVersion = 2
$script:LockSchemaVersion = 1
$script:DiagnosticBundleSchemaVersion = 1
$script:WebhookPayloadSchemaVersion = 2
$script:WebhookDeliverySchemaVersion = 1
$script:MaxContinuationAttempts = 3
$script:RunId = [guid]::NewGuid().ToString()
$script:RunStartedAt = Get-Date
$script:EntryScriptPath = [string]$PSCommandPath
$script:DataPath = "C:\ProgramData\SystemUpdatePro"
$script:LockFile = "C:\ProgramData\SystemUpdatePro\update.lock"
$script:StateFile = "C:\ProgramData\SystemUpdatePro\state.json"
$script:HistoryFile = "C:\ProgramData\SystemUpdatePro\update_history.json"
$script:TaskName = "SystemUpdatePro_Continue"
$script:MutationJournalSchemaVersion = 1
$script:MutationJournalDirectory = "C:\ProgramData\SystemUpdatePro\Journals"
$script:WebhookDeliveryDirectory = "C:\ProgramData\SystemUpdatePro\WebhookDeliveries"
$script:MutationJournal = $null
$script:MutationEvidence = [System.Collections.ArrayList]::new()
$script:WindowsRoot = [string]$env:SystemRoot

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
$script:FirmwarePrerequisites = $null
$script:AcquisitionManifestVersion = 1
$script:AcquisitionManifest = $null
$script:AcquisitionProvenance = [System.Collections.ArrayList]::new()
$script:CapabilityAssessment = $null
$script:RetentionResult = $null
$script:SensitiveEvidenceValues = [System.Collections.ArrayList]::new()
$script:ProtectedEvidenceDirectories = @{}
$script:PSModuleInstallRoot = Join-Path ([Environment]::GetFolderPath("ProgramFiles")) "WindowsPowerShell\Modules"
$script:HPIAInstallRoot = "C:\ProgramData\SystemUpdatePro\HPIA"

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
# PROTECTED LOCAL EVIDENCE STORE
# ============================================================================

function Test-CurrentIdentityIsAdministrator {
    try {
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch {
        return $false
    }
}

function Set-EvidencePathAccess {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseShouldProcessForStateChangingFunctions", "", Justification = "Applies the product's private evidence ACL to an explicitly supplied path.")]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    try {
        $item = Get-Item -LiteralPath $Path -Force -ErrorAction Stop
        $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent().User
        $systemIdentity = New-Object Security.Principal.SecurityIdentifier("S-1-5-18")
        $administratorIdentity = New-Object Security.Principal.SecurityIdentifier("S-1-5-32-544")
        $identities = @($systemIdentity, $administratorIdentity)
        # Unit tests and read-only development runs can execute without an
        # elevated administrator token. Production initialization has already
        # enforced elevation, so production ACLs contain only SYSTEM/Admins.
        if (-not (Test-CurrentIdentityIsAdministrator) -and
            $currentIdentity.Value -notin @($systemIdentity.Value, $administratorIdentity.Value)) {
            $identities += $currentIdentity
        }
        $identities = @($identities | Sort-Object -Property Value -Unique)

        if ($item.PSIsContainer) {
            $acl = New-Object System.Security.AccessControl.DirectorySecurity
            $inheritance = [Security.AccessControl.InheritanceFlags]::ContainerInherit -bor
                [Security.AccessControl.InheritanceFlags]::ObjectInherit
            $propagation = [Security.AccessControl.PropagationFlags]::None
        } else {
            $acl = New-Object System.Security.AccessControl.FileSecurity
            $inheritance = [Security.AccessControl.InheritanceFlags]::None
            $propagation = [Security.AccessControl.PropagationFlags]::None
        }
        $acl.SetOwner($currentIdentity)
        $acl.SetAccessRuleProtection($true, $false)
        foreach ($identity in $identities) {
            $rule = New-Object System.Security.AccessControl.FileSystemAccessRule(
                $identity,
                [Security.AccessControl.FileSystemRights]::FullControl,
                $inheritance,
                $propagation,
                [Security.AccessControl.AccessControlType]::Allow
            )
            [void]$acl.AddAccessRule($rule)
        }

        if ($PSVersionTable.PSEdition -eq "Core") {
            [IO.FileSystemAclExtensions]::SetAccessControl($item, $acl)
        } else {
            $item.SetAccessControl($acl)
        }
        return $true
    } catch {
        $script:LastEvidenceAccessError = $_.Exception.Message
        return $false
    }
}

function Test-EvidencePathAccess {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    try {
        $acl = Get-Acl -LiteralPath $Path -ErrorAction Stop
        if (-not $acl.AreAccessRulesProtected) {
            return [PSCustomObject]@{ Valid = $false; Reason = "Evidence path inherits access rules" }
        }

        $writeSafety = Test-EvidencePathWriteSafety -Path $Path
        if (-not $writeSafety.Valid) { return $writeSafety }

        $requiredWriterSids = @("S-1-5-18", "S-1-5-32-544")
        if (-not (Test-CurrentIdentityIsAdministrator)) {
            $requiredWriterSids += [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
        }
        foreach ($requiredSid in $requiredWriterSids) {
            $hasFullControl = @($acl.Access | Where-Object {
                $_.AccessControlType -eq [Security.AccessControl.AccessControlType]::Allow -and
                $_.IdentityReference.Translate(
                    [Security.Principal.SecurityIdentifier]
                ).Value -eq $requiredSid -and
                ($_.FileSystemRights -band [Security.AccessControl.FileSystemRights]::FullControl) -eq
                    [Security.AccessControl.FileSystemRights]::FullControl
            }).Count -gt 0
            if (-not $hasFullControl) {
                return [PSCustomObject]@{
                    Valid = $false
                    Reason = "Evidence path is missing the required full-control SID $requiredSid"
                }
            }
        }
        return [PSCustomObject]@{ Valid = $true; Reason = "" }
    } catch {
        return [PSCustomObject]@{
            Valid = $false
            Reason = "Evidence ACL could not be verified: $($_.Exception.Message)"
        }
    }
}

function Test-EvidencePathWriteSafety {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [switch]$AllowCurrentIdentity
    )

    try {
        $acl = Get-Acl -LiteralPath $Path -ErrorAction Stop
        $allowedWriterSids = @("S-1-5-18", "S-1-5-32-544")
        if ($AllowCurrentIdentity -or -not (Test-CurrentIdentityIsAdministrator)) {
            $allowedWriterSids += [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
        }
        $writeRights = [Security.AccessControl.FileSystemRights]::Write -bor
            [Security.AccessControl.FileSystemRights]::Modify -bor
            [Security.AccessControl.FileSystemRights]::FullControl -bor
            [Security.AccessControl.FileSystemRights]::WriteData -bor
            [Security.AccessControl.FileSystemRights]::CreateFiles
        foreach ($rule in $acl.Access) {
            if ($rule.AccessControlType -ne [Security.AccessControl.AccessControlType]::Allow -or
                -not ($rule.FileSystemRights -band $writeRights)) {
                continue
            }
            $sid = $rule.IdentityReference.Translate(
                [Security.Principal.SecurityIdentifier]
            ).Value
            if ($sid -notin $allowedWriterSids) {
                return [PSCustomObject]@{
                    Valid = $false
                    Reason = "Evidence path grants write access to unapproved SID $sid"
                }
            }
        }
        return [PSCustomObject]@{ Valid = $true; Reason = "" }
    } catch {
        return [PSCustomObject]@{
            Valid = $false
            Reason = "Evidence write ACL could not be verified: $($_.Exception.Message)"
        }
    }
}

function New-ProtectedDirectory {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseShouldProcessForStateChangingFunctions", "", Justification = "Creates and ACL-hardens an explicitly supplied evidence directory.")]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    try {
        $resolvedPath = [IO.Path]::GetFullPath($Path).TrimEnd("\", "/")
        if ($null -eq $script:ProtectedEvidenceDirectories) {
            $script:ProtectedEvidenceDirectories = @{}
        }
        if ($script:ProtectedEvidenceDirectories.ContainsKey($resolvedPath) -and
            (Test-Path -LiteralPath $resolvedPath -PathType Container)) {
            return $true
        }
        if (Test-Path -LiteralPath $resolvedPath -PathType Container) {
            $dataRoot = [IO.Path]::GetFullPath([string]$script:DataPath).TrimEnd("\", "/")
            $isProductDataPath = (
                $resolvedPath.Equals($dataRoot, [StringComparison]::OrdinalIgnoreCase) -or
                $resolvedPath.StartsWith(
                    "$dataRoot$([IO.Path]::DirectorySeparatorChar)",
                    [StringComparison]::OrdinalIgnoreCase
                )
            )
            if (-not $isProductDataPath) {
                $unownedEntries = @(Get-ChildItem -LiteralPath $resolvedPath -Force `
                    -ErrorAction Stop | Where-Object {
                        $_.Name -notmatch "^(SystemUpdatePro_|DCU_(?:Scan|Apply)_|HPIA_\d{8}_\d{6})"
                    })
                if ($unownedEntries.Count -gt 0) {
                    throw "Custom evidence directory contains unowned content; use a dedicated empty LogPath"
                }
            }
        }
        if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
            New-Item -ItemType Directory -Path $Path -Force -ErrorAction Stop | Out-Null
        }
        if (-not (Set-EvidencePathAccess -Path $Path)) {
            throw "Access controls could not be applied: $($script:LastEvidenceAccessError)"
        }
        $script:ProtectedEvidenceDirectories[$resolvedPath] = $true
        return $true
    } catch {
        $script:LastEvidenceAccessError = $_.Exception.Message
        return $false
    }
}

function Add-SensitiveEvidenceValue {
    param(
        [AllowNull()][string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) { return }
    if ($null -eq $script:SensitiveEvidenceValues) {
        $script:SensitiveEvidenceValues = [System.Collections.ArrayList]::new()
    }
    if ($Value -notin @($script:SensitiveEvidenceValues)) {
        [void]$script:SensitiveEvidenceValues.Add($Value)
    }
}

function Protect-EvidenceText {
    param(
        [AllowNull()][object]$Text,
        [AllowEmptyCollection()][string[]]$SensitiveValues = @()
    )

    if ($null -eq $Text) { return "" }
    $safeText = [string]$Text
    $safeText = [regex]::Replace(
        $safeText,
        "(?i)(https://hooks\.slack\.com/services/)[^\s`"'<]+",
        '${1}[REDACTED]'
    )
    $safeText = [regex]::Replace(
        $safeText,
        "(?i)([?&](?:sig|signature|token|access_token|code|key|api[_-]?key|secret|password)=)[^&\s`"'<>]+",
        '${1}[REDACTED]'
    )
    $safeText = [regex]::Replace(
        $safeText,
        "(?i)(Authorization\s*[:=]\s*(?:Bearer|Basic)\s+)[A-Za-z0-9._~+/=-]+",
        '${1}[REDACTED]'
    )
    $safeText = [regex]::Replace(
        $safeText,
        "(?i)(https?://)[^/\s:@]+:[^@/\s]+@",
        '${1}[REDACTED]@'
    )

    if (-not [string]::IsNullOrWhiteSpace([string]$WebhookUrl)) {
        $safeText = [regex]::Replace(
            $safeText,
            [regex]::Escape([string]$WebhookUrl),
            "[REDACTED WEBHOOK]",
            [Text.RegularExpressions.RegexOptions]::IgnoreCase
        )
    }
    if ($RedactionMode -eq "SecretsAndSerials") {
        $allSensitiveValues = @($SensitiveValues) + @($script:SensitiveEvidenceValues)
        foreach ($value in @($allSensitiveValues | Where-Object {
            -not [string]::IsNullOrWhiteSpace([string]$_)
        } | Select-Object -Unique)) {
            $safeText = [regex]::Replace(
                $safeText,
                [regex]::Escape([string]$value),
                "[REDACTED SERIAL]",
                [Text.RegularExpressions.RegexOptions]::IgnoreCase
            )
        }
    }
    return $safeText
}

function Protect-EvidenceObject {
    param(
        [AllowNull()][object]$InputObject,
        [string]$PropertyName = ""
    )

    if ($null -eq $InputObject) { return $null }
    $isSecretReferenceProperty = $PropertyName -match "(?i)(reference|ref)$"
    $isSecretProperty = (
        $PropertyName -match "(?i)(webhook|authorization|password|secret|token|api.?key|signature)" -and
        -not $isSecretReferenceProperty
    )
    $isSerialProperty = $PropertyName -match "(?i)^serial[_-]?(number)?$"
    if ($isSecretProperty) {
        if ($InputObject -is [string]) {
            Add-SensitiveEvidenceValue -Value ([string]$InputObject)
        }
        return "[REDACTED]"
    }
    if ($isSerialProperty) {
        if ($InputObject -is [string]) {
            Add-SensitiveEvidenceValue -Value ([string]$InputObject)
        }
    }
    if ($isSerialProperty -and $RedactionMode -eq "SecretsAndSerials") {
        return "[REDACTED]"
    }
    if ($InputObject -is [string]) {
        return Protect-EvidenceText -Text $InputObject
    }
    if ($InputObject -is [ValueType] -or $InputObject -is [datetime] -or
        $InputObject -is [version]) {
        return $InputObject
    }
    if ($InputObject -is [System.Collections.IDictionary]) {
        $safeDictionary = [ordered]@{}
        foreach ($key in $InputObject.Keys) {
            $safeDictionary[[string]$key] = Protect-EvidenceObject `
                -InputObject $InputObject[$key] -PropertyName ([string]$key)
        }
        return ,$safeDictionary
    }
    if ($InputObject -is [System.Management.Automation.PSCustomObject]) {
        $safeDictionary = [ordered]@{}
        foreach ($property in $InputObject.PSObject.Properties) {
            $safeDictionary[$property.Name] = Protect-EvidenceObject `
                -InputObject $property.Value -PropertyName $property.Name
        }
        return ,$safeDictionary
    }
    if ($InputObject -is [System.Collections.IEnumerable]) {
        return ,@($InputObject | ForEach-Object {
            Protect-EvidenceObject -InputObject $_ -PropertyName $PropertyName
        })
    }
    return Protect-EvidenceText -Text $InputObject
}

function Write-ProtectedAtomicFile {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseShouldProcessForStateChangingFunctions", "", Justification = "Writes only to an explicitly supplied evidence path through a validated atomic replace.")]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Content,
        [AllowNull()][scriptblock]$ValidationScript = $null,
        [bool]$KeepLastKnownGood = $true
    )

    $temporaryPath = "$Path.tmp.$PID.$([guid]::NewGuid().ToString('N'))"
    $backupPath = "$Path.previous"
    $stream = $null
    try {
        $directory = Split-Path -Parent $Path
        if ([string]::IsNullOrWhiteSpace($directory) -or
            -not (New-ProtectedDirectory -Path $directory)) {
            throw "Evidence directory could not be protected: $($script:LastEvidenceAccessError)"
        }

        $bytes = (New-Object Text.UTF8Encoding($false)).GetBytes($Content)
        $stream = [IO.FileStream]::new(
            $temporaryPath,
            [IO.FileMode]::CreateNew,
            [IO.FileAccess]::Write,
            [IO.FileShare]::None,
            4096,
            [IO.FileOptions]::WriteThrough
        )
        $stream.Write($bytes, 0, $bytes.Length)
        $stream.Flush($true)
        $stream.Dispose()
        $stream = $null

        if ($null -ne $ValidationScript) {
            $validation = & $ValidationScript $temporaryPath
            if (($validation -is [bool] -and -not $validation) -or
                ($validation.PSObject.Properties["Valid"] -and -not [bool]$validation.Valid)) {
                $reason = if ($validation.PSObject.Properties["Reason"]) {
                    [string]$validation.Reason
                } else {
                    "Validation rejected the temporary payload"
                }
                throw $reason
            }
        }
        if (-not (Set-EvidencePathAccess -Path $temporaryPath)) {
            throw "Temporary evidence access controls could not be hardened: $($script:LastEvidenceAccessError)"
        }

        if (Test-Path -LiteralPath $Path) {
            Remove-Item -LiteralPath $backupPath -Force -ErrorAction SilentlyContinue
            [IO.File]::Replace($temporaryPath, $Path, $backupPath, $true)
            if ($KeepLastKnownGood) {
                if (-not (Set-EvidencePathAccess -Path $backupPath)) {
                    throw "Last-known-good evidence backup could not be protected"
                }
            } else {
                Remove-Item -LiteralPath $backupPath -Force -ErrorAction Stop
            }
        } else {
            [IO.File]::Move($temporaryPath, $Path)
        }
        if (-not (Set-EvidencePathAccess -Path $Path)) {
            throw "Evidence access controls could not be hardened: $($script:LastEvidenceAccessError)"
        }
        return $true
    } catch {
        $script:LastEvidenceWriteError = $_.Exception.Message
        if ($null -ne $stream) { $stream.Dispose() }
        Remove-Item -LiteralPath $temporaryPath -Force -ErrorAction SilentlyContinue
        return $false
    }
}

function Write-ProtectedAtomicJson {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseShouldProcessForStateChangingFunctions", "", Justification = "Serializes and atomically replaces an explicitly supplied protected JSON evidence file.")]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter(Mandatory = $true)]
        [AllowNull()][object]$Data,
        [int]$Depth = 24,
        [AllowNull()][scriptblock]$DataValidationScript = $null,
        [bool]$KeepLastKnownGood = $true
    )

    try {
        if ($null -ne $DataValidationScript) {
            $dataValidation = & $DataValidationScript $Data
            if (($dataValidation -is [bool] -and -not $dataValidation) -or
                ($dataValidation.PSObject.Properties["Valid"] -and
                    -not [bool]$dataValidation.Valid)) {
                $reason = if ($dataValidation.PSObject.Properties["Reason"]) {
                    [string]$dataValidation.Reason
                } else {
                    "JSON data validation failed"
                }
                throw $reason
            }
        }
        $json = $Data | ConvertTo-Json -Depth $Depth
        $validationScript = {
            param([string]$CandidatePath)
            try {
                $candidateObject = [IO.File]::ReadAllText($CandidatePath) |
                    ConvertFrom-Json -ErrorAction Stop
                return ($null -ne $candidateObject)
            } catch {
                return [PSCustomObject]@{ Valid = $false; Reason = $_.Exception.Message }
            }
        }
        return Write-ProtectedAtomicFile -Path $Path -Content $json `
            -ValidationScript $validationScript -KeepLastKnownGood:$KeepLastKnownGood
    } catch {
        $script:LastEvidenceWriteError = $_.Exception.Message
        return $false
    }
}

function Add-ProtectedEvidenceLine {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Line
    )

    $stream = $null
    try {
        $directory = Split-Path -Parent $Path
        if (-not (New-ProtectedDirectory -Path $directory)) {
            throw "Evidence directory could not be protected"
        }
        $wasCreated = -not (Test-Path -LiteralPath $Path)
        $bytes = (New-Object Text.UTF8Encoding($false)).GetBytes(
            "$(Protect-EvidenceText -Text $Line)$([Environment]::NewLine)"
        )
        $stream = [IO.FileStream]::new(
            $Path,
            [IO.FileMode]::Append,
            [IO.FileAccess]::Write,
            [IO.FileShare]::ReadWrite,
            4096,
            [IO.FileOptions]::WriteThrough
        )
        $stream.Write($bytes, 0, $bytes.Length)
        $stream.Flush($true)
        $stream.Dispose()
        $stream = $null
        if ($wasCreated -and -not (Set-EvidencePathAccess -Path $Path)) {
            throw "Appended evidence could not be protected"
        }
        return $true
    } catch {
        if ($null -ne $stream) { $stream.Dispose() }
        $script:LastEvidenceWriteError = $_.Exception.Message
        return $false
    }
}

function Move-EvidenceToQuarantine {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseShouldProcessForStateChangingFunctions", "", Justification = "Moves an explicitly supplied corrupt evidence file beside its source for recovery inspection.")]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [string]$Reason = ""
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return "" }
    $directory = Split-Path -Parent $Path
    $extension = [IO.Path]::GetExtension($Path)
    $baseName = [IO.Path]::GetFileNameWithoutExtension($Path)
    $quarantineName = "{0}.corrupt.{1}.{2}{3}" -f @(
        $baseName,
        (Get-Date -Format "yyyyMMddTHHmmss"),
        ([guid]::NewGuid().ToString("N")),
        $extension
    )
    $quarantinePath = Join-Path $directory $quarantineName
    try {
        Move-Item -LiteralPath $Path -Destination $quarantinePath -ErrorAction Stop
        if (-not (Set-EvidencePathAccess -Path $quarantinePath)) {
            throw "Quarantined evidence could not be protected"
        }
        $script:LastEvidenceQuarantineReason = $Reason
        return $quarantinePath
    } catch {
        $script:LastEvidenceWriteError = $_.Exception.Message
        return ""
    }
}

function Read-ProtectedJsonFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [AllowNull()][scriptblock]$MigrationScript = $null,
        [AllowNull()][scriptblock]$ValidationScript = $null
    )

    $failures = [System.Collections.ArrayList]::new()
    foreach ($candidate in @($Path, "$Path.previous")) {
        if (-not (Test-Path -LiteralPath $candidate -PathType Leaf)) { continue }
        try {
            $access = Test-EvidencePathAccess -Path $candidate
            if (-not $access.Valid -and (Test-CurrentIdentityIsAdministrator)) {
                $writeSafety = Test-EvidencePathWriteSafety -Path $candidate `
                    -AllowCurrentIdentity
                if (-not $writeSafety.Valid) { throw $writeSafety.Reason }
                [void](Set-EvidencePathAccess -Path $candidate)
                $access = Test-EvidencePathAccess -Path $candidate
            }
            if (-not $access.Valid) { throw $access.Reason }
            $data = ConvertTo-Hashtable -InputObject (
                [IO.File]::ReadAllText($candidate) | ConvertFrom-Json -ErrorAction Stop
            )
            if ($null -ne $MigrationScript) {
                $data = & $MigrationScript $data
            }
            if ($null -ne $ValidationScript) {
                $validation = & $ValidationScript $data
                if (($validation -is [bool] -and -not $validation) -or
                    ($validation.PSObject.Properties["Valid"] -and -not [bool]$validation.Valid)) {
                    $reason = if ($validation.PSObject.Properties["Reason"]) {
                        [string]$validation.Reason
                    } else {
                        "Evidence validation failed"
                    }
                    throw $reason
                }
            }

            $recovered = $candidate -ne $Path
            if ($recovered -and -not (Write-ProtectedAtomicJson -Path $Path -Data $data `
                -DataValidationScript $ValidationScript)) {
                throw "Last-known-good payload was valid but could not be restored: $($script:LastEvidenceWriteError)"
            }
            return [PSCustomObject]@{
                Success = $true; Data = $data; Recovered = $recovered
                Source = $candidate; Error = ""
            }
        } catch {
            [void]$failures.Add("$candidate`: $($_.Exception.Message)")
            [void](Move-EvidenceToQuarantine -Path $candidate -Reason $_.Exception.Message)
        }
    }
    return [PSCustomObject]@{
        Success = $false; Data = $null; Recovered = $false; Source = ""
        Error = $failures -join "; "
    }
}

function Protect-EvidenceFile {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseShouldProcessForStateChangingFunctions", "", Justification = "ACL-hardens and redacts an explicitly supplied text evidence file.")]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [AllowEmptyCollection()][string[]]$SensitiveValues = @()
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $false }
    if (-not (Set-EvidencePathAccess -Path $Path)) { return $false }
    $extension = [IO.Path]::GetExtension($Path).ToLowerInvariant()
    if ($extension -notin @(".log", ".txt", ".xml", ".json", ".csv", ".html", ".htm")) {
        return $true
    }
    try {
        $file = Get-Item -LiteralPath $Path -ErrorAction Stop
        if ($file.Length -gt 32MB) { return $true }
        $content = [IO.File]::ReadAllText($Path)
        $safeContent = Protect-EvidenceText -Text $content -SensitiveValues $SensitiveValues
        if ($safeContent -ne $content) {
            return Write-ProtectedAtomicFile -Path $Path -Content $safeContent `
                -KeepLastKnownGood:$false
        }
        return $true
    } catch {
        $script:LastEvidenceWriteError = $_.Exception.Message
        return $false
    }
}

function Protect-EvidenceTree {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseShouldProcessForStateChangingFunctions", "", Justification = "Recursively ACL-hardens a product-owned evidence directory and redacts supported text artifacts.")]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [AllowEmptyCollection()][string[]]$SensitiveValues = @()
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Container)) { return $false }
    $succeeded = Set-EvidencePathAccess -Path $Path
    foreach ($item in @(Get-ChildItem -LiteralPath $Path -Force -Recurse -ErrorAction SilentlyContinue)) {
        if ($item.PSIsContainer) {
            if (-not (Set-EvidencePathAccess -Path $item.FullName)) { $succeeded = $false }
        } elseif (-not (Protect-EvidenceFile -Path $item.FullName -SensitiveValues $SensitiveValues)) {
            $succeeded = $false
        }
    }
    return $succeeded
}

function Test-HttpsWebhookEndpoint {
    param(
        [AllowNull()][string]$Endpoint
    )

    if ([string]::IsNullOrWhiteSpace($Endpoint)) {
        return [PSCustomObject]@{ Valid = $false; Reason = "Webhook endpoint is empty" }
    }
    if ($Endpoint.Length -gt 4096 -or $Endpoint -match "[`r`n`0]") {
        return [PSCustomObject]@{
            Valid = $false
            Reason = "Webhook endpoint length or characters are invalid"
        }
    }
    $uri = $null
    if (-not [uri]::TryCreate($Endpoint, [UriKind]::Absolute, [ref]$uri) -or
        $uri.Scheme -ne [Uri]::UriSchemeHttps -or
        [string]::IsNullOrWhiteSpace($uri.DnsSafeHost)) {
        return [PSCustomObject]@{
            Valid = $false
            Reason = "Webhook endpoint must be an absolute HTTPS URL"
        }
    }
    if (-not [string]::IsNullOrWhiteSpace($uri.UserInfo)) {
        return [PSCustomObject]@{
            Valid = $false
            Reason = "Webhook endpoint must not contain URI user information"
        }
    }
    if (-not [string]::IsNullOrWhiteSpace($uri.Fragment)) {
        return [PSCustomObject]@{
            Valid = $false
            Reason = "Webhook endpoint must not contain a fragment"
        }
    }
    return [PSCustomObject]@{ Valid = $true; Reason = ""; Uri = $uri }
}

function Resolve-WebhookSecretReference {
    param(
        [AllowNull()][string]$Reference = $WebhookSecretReference
    )

    if ([string]::IsNullOrWhiteSpace($Reference)) {
        return [PSCustomObject]@{
            Success = $true
            Url = ""
            Source = "None"
            Error = ""
        }
    }

    $endpoint = ""
    $source = ""
    try {
        if ($Reference -match "(?i)^env:(?<VariableName>[A-Z_][A-Z0-9_]*)$") {
            $variableName = [string]$Matches.VariableName
            foreach ($target in @(
                [EnvironmentVariableTarget]::Process,
                [EnvironmentVariableTarget]::Machine,
                [EnvironmentVariableTarget]::User
            )) {
                $endpoint = [Environment]::GetEnvironmentVariable($variableName, $target)
                if (-not [string]::IsNullOrWhiteSpace($endpoint)) { break }
            }
            if ([string]::IsNullOrWhiteSpace($endpoint)) {
                throw "Webhook environment reference is not set"
            }
            $source = "Environment"
        } elseif ($Reference -match "(?i)^file:(?<ConfigPath>[A-Z]:[\\/].+)$") {
            $configPath = [IO.Path]::GetFullPath([string]$Matches.ConfigPath)
            if ([IO.Path]::GetExtension($configPath) -ne ".json" -or
                -not (Test-Path -LiteralPath $configPath -PathType Leaf)) {
                throw "Webhook config reference must name an existing JSON file"
            }
            $access = Test-EvidencePathAccess -Path $configPath
            if (-not $access.Valid) {
                throw "Webhook config ACL is not trusted: $($access.Reason)"
            }
            $config = ConvertTo-Hashtable -InputObject (
                [IO.File]::ReadAllText($configPath) | ConvertFrom-Json -ErrorAction Stop
            )
            if ($config -isnot [System.Collections.IDictionary] -or
                [int]$config.schema_version -ne 1 -or
                -not $config.Contains("webhook_url")) {
                throw "Webhook config must use schema_version 1 and webhook_url"
            }
            $endpoint = [string]$config.webhook_url
            $source = "ProtectedFile"
        } else {
            throw "Webhook secret reference format is invalid"
        }

        $validation = Test-HttpsWebhookEndpoint -Endpoint $endpoint
        if (-not $validation.Valid) { throw $validation.Reason }
        Add-SensitiveEvidenceValue -Value $endpoint
        return [PSCustomObject]@{
            Success = $true
            Url = $endpoint
            Source = $source
            Error = ""
        }
    } catch {
        return [PSCustomObject]@{
            Success = $false
            Url = ""
            Source = $source
            Error = Protect-EvidenceText -Text $_.Exception.Message
        }
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

    $safeMessage = Protect-EvidenceText -Text $Message
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $safeMessage"
    [void](Add-ProtectedEvidenceLine -Path $script:LogFile -Line $logEntry)

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
    $displayMsg = "$($prefixes[$Level])$safeMessage"
    if ($DryRun -and $Level -notin @("HEADER", "DEBUG")) {
        $displayMsg = "[DRY RUN] $displayMsg"
    }

    $consoleColor = Get-ConsoleColorCapability
    if ($consoleColor.SupportsColor) {
        Write-Host $displayMsg -ForegroundColor $colors[$Level]
    } else {
        Write-Host $displayMsg
    }

    # Track warnings and errors
    if ($Level -eq "WARNING") { [void]$script:Warnings.Add($safeMessage) }
    if ($Level -eq "ERROR") { [void]$script:Errors.Add($safeMessage) }
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
        foreach ($path in @(
            $script:DataPath,
            $script:MutationJournalDirectory,
            $script:WebhookDeliveryDirectory,
            $LogPath,
            (Join-Path $script:DataPath "DriverBackups")
        )) {
            if (-not (New-ProtectedDirectory -Path $path)) {
                throw "Could not initialize protected evidence directory '$path': $($script:LastEvidenceAccessError)"
            }
        }
        Add-SensitiveEvidenceValue -Value $WebhookUrl

        try {
            Start-Transcript -Path $script:TranscriptFile -Force -ErrorAction Stop | Out-Null
            $script:TranscriptStarted = $true
            if (-not (Set-EvidencePathAccess -Path $script:TranscriptFile)) {
                throw "Transcript access controls could not be hardened"
            }
        } catch {
            if ($script:TranscriptStarted) {
                try {
                    Stop-Transcript | Out-Null
                } catch {
                    [Diagnostics.Debug]::WriteLine($_.Exception.Message)
                }
            }
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
            $_.Name -in @("OEM", "WindowsUpdate", "Winget", "PackageManagers", "DriverBackup", "DriverRollback", "WindowsUpdateRepair", "Cleanup", "Continuation") -and
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
        Dependencies     = @($script:AcquisitionProvenance)
        MutationRecovery = @($script:MutationEvidence)
        Capabilities     = $script:CapabilityAssessment
        Retention        = $script:RetentionResult
        DependencyReadiness = $script:DependencyReadiness
        DownloadPolicy   = $script:DownloadPolicy
        WingetScopes     = @($script:WingetScopeResults)
        RolloutDecision  = $script:RolloutDecision
        WindowsUpdatePolicy = $script:WindowsUpdatePolicy
        MaintenanceDecision = $script:MaintenanceDecision
        PowerPlanState = $script:PowerPlanState
        Health = [ordered]@{
            PreRun = $script:PreHealthCheck
            PostRun = $script:PostHealthCheck
            Regression = $script:HealthRegression
        }
        Metrics = @{}
        EvidenceDelivery = @{}
    }
}

# ============================================================================
# LOCK FILE MANAGEMENT
# ============================================================================

function Test-LockDocument {
    param(
        [AllowNull()][object]$Lock
    )

    if ($Lock -isnot [System.Collections.IDictionary] -or
        [int]$Lock.SchemaVersion -ne $script:LockSchemaVersion) {
        return [PSCustomObject]@{ Valid = $false; Reason = "Lock schema is invalid" }
    }
    $lockPid = 0
    $lockStarted = [datetime]::MinValue
    if (-not [int]::TryParse([string]$Lock.PID, [ref]$lockPid) -or $lockPid -lt 1 -or
        -not [datetime]::TryParse([string]$Lock.StartTime, [ref]$lockStarted)) {
        return [PSCustomObject]@{ Valid = $false; Reason = "Lock process metadata is invalid" }
    }
    return [PSCustomObject]@{ Valid = $true; Reason = "" }
}

function Test-LockFile {
    if (Test-Path -LiteralPath $script:LockFile) {
        $lockRead = Read-ProtectedJsonFile -Path $script:LockFile `
            -ValidationScript ${function:Test-LockDocument}
        if ($lockRead.Success) {
            $lockContent = $lockRead.Data
            # Check if the process is still running
            $process = Get-Process -Id $lockContent.PID -ErrorAction SilentlyContinue
            if ($process -and $process.ProcessName -in @("powershell", "pwsh")) {
                # Check if lock is stale (older than 4 hours)
                $lockTime = [DateTime]::Parse($lockContent.StartTime)
                if ((Get-Date) - $lockTime -lt [TimeSpan]::FromHours(4)) {
                    return $true  # Lock is valid
                }
            }
        }
        # Stale lock - remove it
        Remove-Item -LiteralPath $script:LockFile -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath "$($script:LockFile).previous" -Force -ErrorAction SilentlyContinue
    }
    return $false
}

function New-LockFile {
    $lockData = [ordered]@{
        SchemaVersion = $script:LockSchemaVersion
        RunId = $script:RunId
        PID = $PID
        StartTime = (Get-Date).ToString("o")
        Computer = $env:COMPUTERNAME
    }
    if (-not (Write-ProtectedAtomicJson -Path $script:LockFile -Data $lockData `
        -DataValidationScript ${function:Test-LockDocument})) {
        throw "Lock file could not be committed: $($script:LastEvidenceWriteError)"
    }
}

function Remove-LockFile {
    Remove-Item -LiteralPath $script:LockFile -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath "$($script:LockFile).previous" -Force -ErrorAction SilentlyContinue
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
        RollbackDrivers    = $RollbackDrivers.IsPresent
        ShowHistory        = $false
        WebhookSecretReference = [string]$WebhookSecretReference
        HistoryCount       = [int]$HistoryCount
        MaxRetries         = [int]$MaxRetries
        MaxUpdatePasses    = [int]$MaxUpdatePasses
        MinDiskSpaceGB     = [int]$MinDiskSpaceGB
        MinFirmwareChargePercent = [int]$MinFirmwareChargePercent
        LogPath            = [string]$LogPath
        LogRetentionDays   = [int]$LogRetentionDays
        EvidenceMaxSizeMB  = [int]$EvidenceMaxSizeMB
        RedactionMode      = [string]$RedactionMode
        Reboot             = $Reboot.IsPresent
        Force              = $Force.IsPresent
        Offline            = $Offline.IsPresent
        DependencyCachePath = [string]$DependencyCachePath
        SourceTimeoutSeconds = [int]$SourceTimeoutSeconds
        AllowMeteredNetwork = $AllowMeteredNetwork.IsPresent
        PolicyPath          = [string]$PolicyPath
        RolloutPolicyPath   = [string]$RolloutPolicyPath
        FeatureDeferralDays = [int]$FeatureDeferralDays
        SecurityOnly        = $SecurityOnly.IsPresent
        PreStage            = $PreStage.IsPresent
        Interactive         = $Interactive.IsPresent
    }
}

function Get-ContinuationParameterName {
    return @(
        "SkipOEM", "SkipWindows", "SkipWinget", "IncludeBIOS", "BypassWSUS",
        "RepairWindowsUpdate", "CleanupAfter", "ResetComponentBase", "ContinueAfterReboot", "DryRun",
        "BackupDrivers", "RollbackDrivers", "ShowHistory", "WebhookSecretReference", "HistoryCount", "MaxRetries",
        "MaxUpdatePasses", "MinDiskSpaceGB", "MinFirmwareChargePercent", "LogPath",
        "LogRetentionDays", "EvidenceMaxSizeMB", "RedactionMode", "Reboot", "Force",
        "Offline", "DependencyCachePath", "SourceTimeoutSeconds", "AllowMeteredNetwork",
        "PolicyPath", "RolloutPolicyPath", "FeatureDeferralDays", "SecurityOnly", "PreStage", "Interactive"
    )
}

function Convert-ContinuationStateSchema {
    param(
        [AllowNull()][object]$State
    )

    $migrated = ConvertTo-Hashtable -InputObject $State
    if ($migrated -isnot [System.Collections.IDictionary]) { return $migrated }
    $schemaVersion = [int]$migrated.SchemaVersion
    $migrationSourceSchema = $schemaVersion
    if ($schemaVersion -eq 3) {
        if ($migrated.Parameters -is [System.Collections.IDictionary]) {
            if (-not $migrated.Parameters.Contains("EvidenceMaxSizeMB")) {
                $migrated.Parameters["EvidenceMaxSizeMB"] = 512
            }
            if (-not $migrated.Parameters.Contains("RedactionMode")) {
                $migrated.Parameters["RedactionMode"] = "SecretsAndSerials"
            }
        }
        $migrated["SchemaVersion"] = 4
        $schemaVersion = 4
    }
    if ($schemaVersion -eq 4) {
        $legacyWebhookUrl = ""
        if ($migrated.Parameters -is [System.Collections.IDictionary]) {
            if ($migrated.Parameters.Contains("WebhookUrl")) {
                $legacyWebhookUrl = [string]$migrated.Parameters.WebhookUrl
                [void]$migrated.Parameters.Remove("WebhookUrl")
            }
            if (-not $migrated.Parameters.Contains("WebhookSecretReference")) {
                $migrated.Parameters["WebhookSecretReference"] = ""
            }
        }
        if (-not $migrated.Contains("ResolvedWebhookUrl")) {
            $migrated["ResolvedWebhookUrl"] = $legacyWebhookUrl
        }
        $migrated["SchemaVersion"] = 5
        $schemaVersion = 5
        $migrated["_MigrationSourceSchema"] = $migrationSourceSchema
    }
    if ($schemaVersion -eq 5 -and $migrated.Parameters -is [System.Collections.IDictionary]) {
        $defaults = [ordered]@{
            Offline = $false
            DependencyCachePath = "C:\ProgramData\SystemUpdatePro\Cache"
            SourceTimeoutSeconds = 30
            AllowMeteredNetwork = $false
            PolicyPath = ""
            RolloutPolicyPath = ""
        }
        foreach ($default in $defaults.GetEnumerator()) {
            if (-not $migrated.Parameters.Contains($default.Key)) {
                $migrated.Parameters[$default.Key] = $default.Value
            }
        }
        if (-not $migrated.Parameters.Contains("FeatureDeferralDays")) {
            $migrated.Parameters["FeatureDeferralDays"] = 0
        }
        if (-not $migrated.Parameters.Contains("SecurityOnly")) {
            $migrated.Parameters["SecurityOnly"] = $false
        }
        if (-not $migrated.Parameters.Contains("PreStage")) {
            $migrated.Parameters["PreStage"] = $false
        }
        $migrated["SchemaVersion"] = 6
        $schemaVersion = 6
        $migrated["_MigrationSourceSchema"] = $migrationSourceSchema
    }
    if ($schemaVersion -eq 6 -and $migrated.Parameters -is [System.Collections.IDictionary]) {
        if (-not $migrated.Parameters.Contains("Interactive")) {
            $migrated.Parameters["Interactive"] = $false
        }
        $migrated["SchemaVersion"] = 7
        $schemaVersion = 7
        $migrated["_MigrationSourceSchema"] = $migrationSourceSchema
    }
    if ($schemaVersion -eq 7 -and $migrated.Parameters -is [System.Collections.IDictionary]) {
        if (-not $migrated.Parameters.Contains("RollbackDrivers")) {
            $migrated.Parameters["RollbackDrivers"] = $false
        }
        $migrated["SchemaVersion"] = 8
        $migrated["_MigrationSourceSchema"] = $migrationSourceSchema
    }
    return $migrated
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
    if ([string]$State.StageCursor -notin @("WindowsUpdate", "Winget", "PackageManagers", "Cleanup", "Complete")) {
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
    if (-not $State.Contains("ResolvedWebhookUrl")) {
        return & $failure "Resolved webhook continuation value is missing"
    }
    if (-not [string]::IsNullOrWhiteSpace([string]$State.ResolvedWebhookUrl)) {
        $webhookValidation = Test-HttpsWebhookEndpoint -Endpoint ([string]$State.ResolvedWebhookUrl)
        if (-not $webhookValidation.Valid) {
            return & $failure "Resolved webhook continuation value is invalid"
        }
    }

    return [PSCustomObject]@{ Valid = $true; Reason = "" }
}

function Set-ContinuationStateAccess {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseShouldProcessForStateChangingFunctions", "", Justification = "Hardens the run-owned state file against non-administrator modification.")]
    param([string]$Path = $script:StateFile)

    $result = Set-EvidencePathAccess -Path $Path
    if (-not $result) { $script:LastStateAccessError = $script:LastEvidenceAccessError }
    return $result
}

function Test-ContinuationStateAccess {
    param([string]$Path = $script:StateFile)

    return Test-EvidencePathAccess -Path $Path
}

function Save-State {
    param([System.Collections.IDictionary]$State)

    try {
        $State["LastUpdatedAt"] = (Get-Date).ToString("o")
        if (-not (Write-ProtectedAtomicJson -Path $script:StateFile -Data $State -Depth 20 `
            -DataValidationScript ${function:Test-ContinuationState})) {
            throw $script:LastEvidenceWriteError
        }
        return $true
    } catch {
        $script:LastStateError = $_.Exception.Message
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

    $quarantinePath = Move-EvidenceToQuarantine -Path $script:StateFile -Reason $Reason
    if ($quarantinePath) {
        try {
            Write-Log "Quarantined invalid continuation state: $Reason ($quarantinePath)" "WARNING"
        } catch {
            [Diagnostics.Debug]::WriteLine($_.Exception.Message)
        }
    }
    return $quarantinePath
}

function Get-State {
    $read = Read-ProtectedJsonFile -Path $script:StateFile `
        -MigrationScript ${function:Convert-ContinuationStateSchema} `
        -ValidationScript ${function:Test-ContinuationState}
    if (-not $read.Success) { return @{} }
    if ($read.Data.Contains("_MigrationSourceSchema")) {
        [void]$read.Data.Remove("_MigrationSourceSchema")
        [void](Write-ProtectedAtomicJson -Path $script:StateFile -Data $read.Data -Depth 20 `
            -DataValidationScript ${function:Test-ContinuationState})
    }
    return $read.Data
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
        ResolvedWebhookUrl = [string]$WebhookUrl
        Parameters     = Get-EffectiveRunParameter
        StageResults   = @($script:StageResults)
        AcquisitionProvenance = @($script:AcquisitionProvenance)
        MutationEvidence = @($script:MutationEvidence)
        Capabilities   = $script:CapabilityAssessment
        DependencyReadiness = $script:DependencyReadiness
        DownloadPolicy = $script:DownloadPolicy
        WingetScopes = @($script:WingetScopeResults)
        RolloutDecision = $script:RolloutDecision
        WindowsUpdatePolicy = $script:WindowsUpdatePolicy
        MaintenanceDecision = $script:MaintenanceDecision
        PowerPlanState = $script:PowerPlanState
        PreHealthCheck = $script:PreHealthCheck
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
        "BackupDrivers", "RollbackDrivers", "ShowHistory", "Reboot", "Force", "SecurityOnly", "PreStage", "Interactive"
    )
    $integerNames = @(
        "HistoryCount", "MaxRetries", "MaxUpdatePasses", "MinDiskSpaceGB",
        "MinFirmwareChargePercent", "LogRetentionDays", "EvidenceMaxSizeMB", "FeatureDeferralDays"
    )

    foreach ($name in $switchNames) {
        Set-Variable -Name $name -Scope Script -Value ([switch][bool]$State.Parameters[$name])
    }
    foreach ($name in $integerNames) {
        Set-Variable -Name $name -Scope Script -Value ([int]$State.Parameters[$name])
    }
    Set-Variable -Name "Offline" -Scope Script -Value ([switch][bool]$State.Parameters.Offline)
    Set-Variable -Name "AllowMeteredNetwork" -Scope Script -Value ([switch][bool]$State.Parameters.AllowMeteredNetwork)
    Set-Variable -Name "SourceTimeoutSeconds" -Scope Script -Value ([int]$State.Parameters.SourceTimeoutSeconds)
    Set-Variable -Name "DependencyCachePath" -Scope Script -Value ([string]$State.Parameters.DependencyCachePath)
    Set-Variable -Name "PolicyPath" -Scope Script -Value ([string]$State.Parameters.PolicyPath)
    Set-Variable -Name "RolloutPolicyPath" -Scope Script -Value ([string]$State.Parameters.RolloutPolicyPath)
    Set-Variable -Name "WebhookSecretReference" -Scope Script `
        -Value ([string]$State.Parameters.WebhookSecretReference)
    Set-Variable -Name "WebhookUrl" -Scope Script -Value ([string]$State.ResolvedWebhookUrl)
    Add-SensitiveEvidenceValue -Value ([string]$State.ResolvedWebhookUrl)
    Set-Variable -Name "LogPath" -Scope Script -Value ([string]$State.Parameters.LogPath)
    Set-Variable -Name "RedactionMode" -Scope Script -Value ([string]$State.Parameters.RedactionMode)

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
    $script:DependencyReadiness = if ($State.Contains("DependencyReadiness")) {
        ConvertTo-Hashtable -InputObject $State.DependencyReadiness
    } else { $null }
    $script:DownloadPolicy = if ($State.Contains("DownloadPolicy")) {
        ConvertTo-Hashtable -InputObject $State.DownloadPolicy
    } else { $null }
    $script:WingetScopeResults = if ($State.Contains("WingetScopes")) {
        @($State.WingetScopes | ForEach-Object { ConvertTo-Hashtable -InputObject $_ })
    } else { @() }
    $script:RolloutDecision = if ($State.Contains("RolloutDecision")) {
        ConvertTo-Hashtable -InputObject $State.RolloutDecision
    } else { $null }
    $script:WindowsUpdatePolicy = if ($State.Contains("WindowsUpdatePolicy")) {
        ConvertTo-Hashtable -InputObject $State.WindowsUpdatePolicy
    } else { $null }
    $script:MaintenanceDecision = if ($State.Contains("MaintenanceDecision")) {
        ConvertTo-Hashtable -InputObject $State.MaintenanceDecision
    } else { $null }
    $script:PowerPlanState = if ($State.Contains("PowerPlanState")) {
        ConvertTo-Hashtable -InputObject $State.PowerPlanState
    } else { $null }
    $script:PreHealthCheck = if ($State.Contains("PreHealthCheck")) {
        ConvertTo-Hashtable -InputObject $State.PreHealthCheck
    } else { $null }
    $script:AcquisitionProvenance = [System.Collections.ArrayList]::new()
    foreach ($dependency in @($State.AcquisitionProvenance)) {
        [void]$script:AcquisitionProvenance.Add([PSCustomObject]$dependency)
    }
    $script:MutationEvidence = [System.Collections.ArrayList]::new()
    foreach ($mutation in @($State.MutationEvidence)) {
        [void]$script:MutationEvidence.Add([PSCustomObject]$mutation)
    }
    if ($State.Contains("Capabilities") -and $null -ne $State.Capabilities) {
        $script:CapabilityAssessment = ConvertTo-Hashtable -InputObject $State.Capabilities
    }

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
        [ValidateSet("WindowsUpdate", "Winget", "PackageManagers", "Cleanup", "Complete")]
        [string]$StageCursor
    )

    if (-not $script:ContinuationActive -or $null -eq $script:ContinuationState) { return $true }

    $script:ContinuationState.StageCursor = $StageCursor
    $script:ContinuationState.StageResults = @($script:StageResults)
    $script:ContinuationState.AcquisitionProvenance = @($script:AcquisitionProvenance)
    $script:ContinuationState.MutationEvidence = @($script:MutationEvidence)
    $script:ContinuationState.Capabilities = $script:CapabilityAssessment
    $script:ContinuationState.DependencyReadiness = $script:DependencyReadiness
    $script:ContinuationState.DownloadPolicy = $script:DownloadPolicy
    $script:ContinuationState.WingetScopes = @($script:WingetScopeResults)
    $script:ContinuationState.RolloutDecision = $script:RolloutDecision
    $script:ContinuationState.WindowsUpdatePolicy = $script:WindowsUpdatePolicy
    $script:ContinuationState.MaintenanceDecision = $script:MaintenanceDecision
    $script:ContinuationState.PowerPlanState = $script:PowerPlanState
    $script:ContinuationState.PreHealthCheck = $script:PreHealthCheck
    $script:ContinuationState.Errors = @($script:Errors)
    $script:ContinuationState.Warnings = @($script:Warnings)
    $script:ContinuationState.Parameters = Get-EffectiveRunParameter
    return Save-State -State $script:ContinuationState
}

function Test-ShouldRunContinuationStage {
    param(
        [ValidateSet("WindowsUpdate", "Winget", "PackageManagers", "Cleanup")]
        [string]$Stage
    )

    if (-not $script:ContinuationActive) { return $true }
    $order = @("WindowsUpdate", "Winget", "PackageManagers", "Cleanup", "Complete")
    $cursorIndex = [array]::IndexOf($order, $script:ResumeStageCursor)
    $stageIndex = [array]::IndexOf($order, $Stage)
    return ($stageIndex -ge $cursorIndex)
}

# ============================================================================
# PRIVILEGED MUTATION JOURNAL
# ============================================================================

function Get-MutationJournalPath {
    param([string]$RunId = $script:RunId)

    $parsedRunId = [guid]::Empty
    if (-not [guid]::TryParse($RunId, [ref]$parsedRunId)) {
        throw "Mutation journal run ID is invalid"
    }
    return Join-Path $script:MutationJournalDirectory "$($parsedRunId.ToString()).json"
}

function New-MutationJournal {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseShouldProcessForStateChangingFunctions", "", Justification = "Creates an in-memory recovery journal only.")]
    param([string]$RunId = $script:RunId)

    return [ordered]@{
        SchemaVersion = $script:MutationJournalSchemaVersion
        RunId          = $RunId
        Status         = "Open"
        CreatedAt      = (Get-Date).ToUniversalTime().ToString("o")
        LastUpdatedAt  = (Get-Date).ToUniversalTime().ToString("o")
        Entries        = @()
    }
}

function Test-MutationJournal {
    param([AllowNull()][object]$Journal)

    if ($Journal -isnot [System.Collections.IDictionary]) {
        return [PSCustomObject]@{ Valid = $false; Reason = "Journal root is not an object" }
    }
    if ([int]$Journal.SchemaVersion -ne $script:MutationJournalSchemaVersion) {
        return [PSCustomObject]@{ Valid = $false; Reason = "Unsupported mutation journal schema" }
    }
    $parsedRunId = [guid]::Empty
    if (-not [guid]::TryParse([string]$Journal.RunId, [ref]$parsedRunId)) {
        return [PSCustomObject]@{ Valid = $false; Reason = "Mutation journal run ID is invalid" }
    }
    if ([string]$Journal.Status -notin @("Open", "AwaitingContinuation", "Restoring", "RecoveryFailed")) {
        return [PSCustomObject]@{ Valid = $false; Reason = "Mutation journal status is invalid" }
    }
    foreach ($entry in @($Journal.Entries)) {
        if ($entry -isnot [System.Collections.IDictionary]) {
            return [PSCustomObject]@{ Valid = $false; Reason = "Mutation journal entry is not an object" }
        }
        $parsedEntryId = [guid]::Empty
        if (-not [guid]::TryParse([string]$entry.Id, [ref]$parsedEntryId)) {
            return [PSCustomObject]@{ Valid = $false; Reason = "Mutation journal entry ID is invalid" }
        }
        if ([string]$entry.Type -notin @("RegistryValue", "Service", "DirectoryRename", "ScheduledTask")) {
            return [PSCustomObject]@{ Valid = $false; Reason = "Mutation journal entry type is invalid" }
        }
        if ([string]$entry.State -notin @("Planned", "Applied", "Committing", "Restored", "Committed", "RecoveryFailed")) {
            return [PSCustomObject]@{ Valid = $false; Reason = "Mutation journal entry state is invalid" }
        }
        if ([int]$entry.Sequence -lt 1 -or [string]::IsNullOrWhiteSpace([string]$entry.Target) -or
            [string]::IsNullOrWhiteSpace([string]$entry.Scope) -or
            $entry.Before -isnot [System.Collections.IDictionary]) {
            return [PSCustomObject]@{ Valid = $false; Reason = "Mutation journal entry contract is incomplete" }
        }
    }
    return [PSCustomObject]@{ Valid = $true; Reason = "" }
}

function Save-MutationJournal {
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.IDictionary]$Journal
    )

    $journalPath = Get-MutationJournalPath -RunId ([string]$Journal.RunId)
    try {
        $Journal["LastUpdatedAt"] = (Get-Date).ToUniversalTime().ToString("o")
        $validation = Test-MutationJournal -Journal $Journal
        if (-not $validation.Valid) { throw $validation.Reason }
        if (-not (Write-ProtectedAtomicJson -Path $journalPath -Data $Journal -Depth 24 `
            -DataValidationScript ${function:Test-MutationJournal})) {
            throw $script:LastEvidenceWriteError
        }
        return $true
    } catch {
        $script:LastMutationJournalError = $_.Exception.Message
        return $false
    }
}

function Get-MutationJournal {
    param([string]$RunId = $script:RunId)

    $journalPath = Get-MutationJournalPath -RunId $RunId
    $read = Read-ProtectedJsonFile -Path $journalPath `
        -ValidationScript ${function:Test-MutationJournal}
    if ($read.Success) { return $read.Data }
    $script:LastMutationJournalError = $read.Error
    return $null
}

function Get-OrCreateMutationJournal {
    if ($null -ne $script:MutationJournal -and
        [string]$script:MutationJournal.RunId -eq [string]$script:RunId) {
        return $script:MutationJournal
    }

    $script:MutationJournal = Get-MutationJournal -RunId $script:RunId
    if ($null -eq $script:MutationJournal) {
        $journalPath = Get-MutationJournalPath -RunId $script:RunId
        if (Test-Path -LiteralPath $journalPath) {
            throw "Existing mutation journal could not be validated: $($script:LastMutationJournalError)"
        }
        $script:MutationJournal = New-MutationJournal -RunId $script:RunId
        if (-not (Save-MutationJournal -Journal $script:MutationJournal)) {
            throw "Mutation journal could not be created: $($script:LastMutationJournalError)"
        }
    }
    return $script:MutationJournal
}

function Add-MutationJournalEntry {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseShouldProcessForStateChangingFunctions", "", Justification = "Persists a before-image before an authorized privileged mutation.")]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("RegistryValue", "Service", "DirectoryRename", "ScheduledTask")]
        [string]$Type,
        [Parameter(Mandatory = $true)]
        [string]$Target,
        [Parameter(Mandatory = $true)]
        [System.Collections.IDictionary]$Before,
        [Parameter(Mandatory = $true)]
        [string]$RecoveryAction,
        [string]$Scope = "Run",
        [bool]$RestoreOnFinalize = $true,
        [AllowNull()][System.Collections.IDictionary]$Metadata = $null
    )

    $journal = Get-OrCreateMutationJournal
    foreach ($existing in @($journal.Entries)) {
        if ([string]$existing.Type -eq $Type -and [string]$existing.Target -eq $Target -and
            [string]$existing.Scope -eq $Scope -and
            [string]$existing.State -notin @("Restored", "Committed")) {
            return [string]$existing.Id
        }
    }

    $entry = [ordered]@{
        Id                = [guid]::NewGuid().ToString()
        Sequence          = @($journal.Entries).Count + 1
        Type              = $Type
        Target            = $Target
        Scope             = $Scope
        State             = "Planned"
        RestoreOnFinalize = $RestoreOnFinalize
        RecoveryAction    = $RecoveryAction
        Before            = $Before
        Metadata          = $(if ($null -eq $Metadata) { [ordered]@{} } else { $Metadata })
        CreatedAt         = (Get-Date).ToUniversalTime().ToString("o")
        AppliedAt         = ""
        CommitStartedAt   = ""
        RecoveredAt       = ""
        Error             = ""
    }
    $journal.Entries = @($journal.Entries) + @($entry)
    if (-not (Save-MutationJournal -Journal $journal)) {
        $journal.Entries = @($journal.Entries | Where-Object { [string]$_.Id -ne [string]$entry.Id })
        throw "Before-image for '$Target' could not be journaled: $($script:LastMutationJournalError)"
    }

    [void]$script:MutationEvidence.Add([PSCustomObject][ordered]@{
        EntryId = $entry.Id; Scope = $Scope; Target = $Target; Action = "Journaled"
        Timestamp = (Get-Date).ToUniversalTime().ToString("o"); Detail = $RecoveryAction
    })
    return [string]$entry.Id
}

function Set-MutationJournalEntryState {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseShouldProcessForStateChangingFunctions", "", Justification = "Persists mutation transaction state.")]
    param(
        [Parameter(Mandatory = $true)]
        [string]$EntryId,
        [Parameter(Mandatory = $true)]
        [ValidateSet("Applied", "Committing", "Restored", "Committed", "RecoveryFailed")]
        [string]$State,
        [string]$ErrorMessage = ""
    )

    $journal = $script:MutationJournal
    if ($null -eq $journal) {
        $journal = Get-OrCreateMutationJournal
    }
    $entry = @($journal.Entries | Where-Object { [string]$_.Id -eq $EntryId } | Select-Object -First 1)
    if ($entry.Count -ne 1) { throw "Mutation journal entry '$EntryId' was not found" }
    $entry = $entry[0]
    $entry.State = $State
    $entry.Error = $ErrorMessage
    if ($State -eq "Applied") {
        $entry.AppliedAt = (Get-Date).ToUniversalTime().ToString("o")
    } elseif ($State -eq "Committing") {
        $entry.CommitStartedAt = (Get-Date).ToUniversalTime().ToString("o")
    } elseif ($State -in @("Restored", "Committed")) {
        $entry.RecoveredAt = (Get-Date).ToUniversalTime().ToString("o")
    }
    if (-not (Save-MutationJournal -Journal $journal)) {
        throw "Mutation journal state '$State' could not be persisted: $($script:LastMutationJournalError)"
    }
    return $true
}

function Test-MutationDirectoryContract {
    param(
        [Parameter(Mandatory = $true)]
        [string]$OriginalPath,
        [Parameter(Mandatory = $true)]
        [string]$BackupPath
    )

    $windowsRoot = [IO.Path]::GetFullPath([string]$script:WindowsRoot).TrimEnd("\", "/")
    $original = [IO.Path]::GetFullPath($OriginalPath).TrimEnd("\", "/")
    $backup = [IO.Path]::GetFullPath($BackupPath).TrimEnd("\", "/")
    $allowedOriginals = @(
        [IO.Path]::GetFullPath((Join-Path $windowsRoot "SoftwareDistribution")).TrimEnd("\", "/"),
        [IO.Path]::GetFullPath((Join-Path $windowsRoot "System32\catroot2")).TrimEnd("\", "/")
    )
    return (
        $original -in $allowedOriginals -and
        $backup.StartsWith("$original.SystemUpdatePro.", [StringComparison]::OrdinalIgnoreCase) -and
        $backup.EndsWith(".bak", [StringComparison]::OrdinalIgnoreCase)
    )
}

function Restore-DirectoryRenameSnapshot {
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.IDictionary]$Snapshot
    )

    $originalPath = [string]$Snapshot.OriginalPath
    $backupPath = [string]$Snapshot.BackupPath
    if (-not (Test-MutationDirectoryContract -OriginalPath $originalPath -BackupPath $backupPath)) {
        return $false
    }

    try {
        foreach ($serviceName in @($Snapshot.Services)) {
            if ([string]::IsNullOrWhiteSpace([string]$serviceName)) { continue }
            $service = Get-Service -Name ([string]$serviceName) -ErrorAction SilentlyContinue
            if ($service -and $service.Status -ne "Stopped") {
                Stop-Service -Name ([string]$serviceName) -Force -ErrorAction Stop
            }
        }

        if ([bool]$Snapshot.Exists) {
            if (-not (Test-Path -LiteralPath $backupPath -PathType Container)) {
                # A planned entry can be recovered before its rename occurred.
                return (Test-Path -LiteralPath $originalPath -PathType Container)
            }
            if (Test-Path -LiteralPath $originalPath) {
                Remove-Item -LiteralPath $originalPath -Recurse -Force -ErrorAction Stop
            }
            Move-Item -LiteralPath $backupPath -Destination $originalPath -ErrorAction Stop
            return (
                (Test-Path -LiteralPath $originalPath -PathType Container) -and
                -not (Test-Path -LiteralPath $backupPath)
            )
        }

        if (Test-Path -LiteralPath $originalPath) {
            Remove-Item -LiteralPath $originalPath -Recurse -Force -ErrorAction Stop
        }
        if (Test-Path -LiteralPath $backupPath) {
            Remove-Item -LiteralPath $backupPath -Recurse -Force -ErrorAction Stop
        }
        return (-not (Test-Path -LiteralPath $originalPath) -and -not (Test-Path -LiteralPath $backupPath))
    } catch {
        return $false
    }
}

function Invoke-JournaledDirectoryReset {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseShouldProcessForStateChangingFunctions", "", Justification = "The durable before-image makes this cache reset reversible.")]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [string[]]$Services,
        [string]$Scope = "WindowsUpdateRepair"
    )

    $resolvedPath = [IO.Path]::GetFullPath($Path).TrimEnd("\", "/")
    $backupPath = "$resolvedPath.SystemUpdatePro.$($script:RunId).bak"
    if (-not (Test-MutationDirectoryContract -OriginalPath $resolvedPath -BackupPath $backupPath)) {
        throw "Directory reset target is outside the approved Windows Update cache paths"
    }
    if (Test-Path -LiteralPath $backupPath) {
        throw "Recovery backup already exists: $backupPath"
    }

    $snapshot = [ordered]@{
        OriginalPath = $resolvedPath
        BackupPath   = $backupPath
        Exists       = (Test-Path -LiteralPath $resolvedPath -PathType Container)
        Services     = @($Services)
    }
    $entryId = Add-MutationJournalEntry -Type "DirectoryRename" -Target $resolvedPath `
        -Before $snapshot -RecoveryAction "Restore the original Windows Update cache directory" `
        -Scope $Scope -RestoreOnFinalize $false

    try {
        if ([bool]$snapshot.Exists) {
            Move-Item -LiteralPath $resolvedPath -Destination $backupPath -ErrorAction Stop
        }
        New-Item -ItemType Directory -Path $resolvedPath -Force -ErrorAction Stop | Out-Null
        if (-not (Test-Path -LiteralPath $resolvedPath -PathType Container) -or
            ([bool]$snapshot.Exists -and -not (Test-Path -LiteralPath $backupPath -PathType Container))) {
            throw "Directory reset verification failed for '$resolvedPath'"
        }
        [void](Set-MutationJournalEntryState -EntryId $entryId -State "Applied")
        return $entryId
    } catch {
        [void](Restore-MutationJournalScope -Scope $Scope)
        throw
    }
}

function Complete-DirectoryRenameSnapshot {
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.IDictionary]$Snapshot
    )

    $originalPath = [string]$Snapshot.OriginalPath
    $backupPath = [string]$Snapshot.BackupPath
    if (-not (Test-MutationDirectoryContract -OriginalPath $originalPath -BackupPath $backupPath)) {
        return $false
    }
    try {
        if (-not (Test-Path -LiteralPath $originalPath -PathType Container)) { return $false }
        if (Test-Path -LiteralPath $backupPath) {
            Remove-Item -LiteralPath $backupPath -Recurse -Force -ErrorAction Stop
        }
        return (-not (Test-Path -LiteralPath $backupPath))
    } catch {
        return $false
    }
}

function Restore-MutationJournalEntry {
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.IDictionary]$Entry
    )

    if ([string]$Entry.State -in @("Restored", "Committed")) { return $true }
    $success = switch ([string]$Entry.Type) {
        "RegistryValue"  { Restore-RegistryValueSnapshot -Snapshot $Entry.Before }
        "Service"        { Restore-ServiceSnapshot -Snapshot $Entry.Before }
        "DirectoryRename" { Restore-DirectoryRenameSnapshot -Snapshot $Entry.Before }
        "ScheduledTask"  { Restore-ScheduledTaskSnapshot -Snapshot $Entry.Before }
        default          { $false }
    }

    if ($success) {
        [void](Set-MutationJournalEntryState -EntryId ([string]$Entry.Id) -State "Restored")
        [void]$script:MutationEvidence.Add([PSCustomObject][ordered]@{
            EntryId = $Entry.Id; Scope = $Entry.Scope; Target = $Entry.Target; Action = "Restored"
            Timestamp = (Get-Date).ToUniversalTime().ToString("o"); Detail = $Entry.RecoveryAction
        })
        return $true
    }

    [void](Set-MutationJournalEntryState -EntryId ([string]$Entry.Id) -State "RecoveryFailed" `
        -ErrorMessage "Recovery verification failed")
    return $false
}

function Complete-MutationJournalEntry {
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.IDictionary]$Entry
    )

    if ([bool]$Entry.RestoreOnFinalize) {
        return Restore-MutationJournalEntry -Entry $Entry
    }
    if ([string]$Entry.State -in @("Restored", "Committed")) { return $true }

    $success = if ([string]$Entry.Type -eq "DirectoryRename") {
        if ([string]$Entry.State -ne "Committing") {
            [void](Set-MutationJournalEntryState -EntryId ([string]$Entry.Id) -State "Committing")
        }
        Complete-DirectoryRenameSnapshot -Snapshot $Entry.Before
    } else {
        $true
    }
    if (-not $success) {
        [void](Set-MutationJournalEntryState -EntryId ([string]$Entry.Id) -State "RecoveryFailed" `
            -ErrorMessage "Commit verification failed")
        return $false
    }

    [void](Set-MutationJournalEntryState -EntryId ([string]$Entry.Id) -State "Committed")
    [void]$script:MutationEvidence.Add([PSCustomObject][ordered]@{
        EntryId = $Entry.Id; Scope = $Entry.Scope; Target = $Entry.Target; Action = "Committed"
        Timestamp = (Get-Date).ToUniversalTime().ToString("o"); Detail = $Entry.RecoveryAction
    })
    return $true
}

function Restore-MutationJournalScope {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Scope
    )

    if ($null -eq $script:MutationJournal) {
        $script:MutationJournal = Get-MutationJournal -RunId $script:RunId
    }
    if ($null -eq $script:MutationJournal) { return $true }

    $success = $true
    $entries = @($script:MutationJournal.Entries | Where-Object {
        [string]$_.Scope -eq $Scope -and [string]$_.State -notin @("Restored", "Committed")
    } | Sort-Object -Property { [int]$_['Sequence'] } -Descending)
    foreach ($entry in $entries) {
        if (-not (Restore-MutationJournalEntry -Entry $entry)) { $success = $false }
    }
    return $success
}

function Complete-MutationJournal {
    param(
        [AllowEmptyCollection()][string[]]$PreserveScopes = @()
    )

    if ($null -eq $script:MutationJournal) {
        $script:MutationJournal = Get-MutationJournal -RunId $script:RunId
    }
    if ($null -eq $script:MutationJournal) { return $true }

    $success = $true
    $entries = @($script:MutationJournal.Entries |
        Sort-Object -Property { [int]$_['Sequence'] } -Descending)
    foreach ($entry in $entries) {
        if ([string]$entry.Scope -in $PreserveScopes) { continue }
        if (-not (Complete-MutationJournalEntry -Entry $entry)) { $success = $false }
    }

    $remaining = @($script:MutationJournal.Entries | Where-Object {
        [string]$_.State -notin @("Restored", "Committed")
    })
    if ($success -and $remaining.Count -eq 0) {
        $journalPath = Get-MutationJournalPath -RunId ([string]$script:MutationJournal.RunId)
        Remove-Item -LiteralPath $journalPath -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath "$journalPath.previous" -Force -ErrorAction SilentlyContinue
        $script:MutationJournal = $null
        return (-not (Test-Path -LiteralPath $journalPath))
    }

    $script:MutationJournal.Status = $(if ($success) { "AwaitingContinuation" } else { "RecoveryFailed" })
    if (-not (Save-MutationJournal -Journal $script:MutationJournal)) { return $false }
    return $success
}

function Invoke-UnfinishedMutationRecovery {
    param([string]$ExcludeRunId = "")

    $result = [ordered]@{ Attempted = 0; Recovered = 0; Failed = 0; Messages = @() }
    if (-not (Test-Path -LiteralPath $script:MutationJournalDirectory)) { return $result }

    foreach ($file in @(Get-ChildItem -LiteralPath $script:MutationJournalDirectory -Filter "*.json" -File -ErrorAction SilentlyContinue)) {
        try {
            $journalRead = Read-ProtectedJsonFile -Path $file.FullName `
                -ValidationScript ${function:Test-MutationJournal}
            if (-not $journalRead.Success) { throw $journalRead.Error }
            $journalObject = $journalRead.Data
            if ([string]$journalObject.RunId -ne [IO.Path]::GetFileNameWithoutExtension($file.Name)) {
                throw "Mutation journal filename does not match its run ID"
            }
            if (-not [string]::IsNullOrWhiteSpace($ExcludeRunId) -and
                [string]$journalObject.RunId -eq $ExcludeRunId) {
                continue
            }

            $result.Attempted++
            $script:MutationJournal = $journalObject
            $script:MutationJournal.Status = "Restoring"
            if (-not (Save-MutationJournal -Journal $script:MutationJournal)) {
                throw "Could not claim mutation journal for recovery"
            }

            $journalRecovered = $true
            foreach ($entry in @($script:MutationJournal.Entries |
                Where-Object { [string]$_.State -notin @("Restored", "Committed") } |
                Sort-Object -Property { [int]$_['Sequence'] } -Descending)) {
                # An interrupted run always rolls back, even when a successful
                # terminal run would have committed a reversible cache swap.
                # A durable Committing state is the sole exception: recovery
                # completes that already-decided cleanup instead of reviving a
                # partially deleted backup.
                $entryRecovered = if ([string]$entry.State -eq "Committing") {
                    Complete-MutationJournalEntry -Entry $entry
                } else {
                    Restore-MutationJournalEntry -Entry $entry
                }
                if (-not $entryRecovered) {
                    $journalRecovered = $false
                }
            }

            if ($journalRecovered) {
                Remove-Item -LiteralPath $file.FullName -Force -ErrorAction Stop
                Remove-Item -LiteralPath "$($file.FullName).previous" -Force -ErrorAction SilentlyContinue
                $result.Recovered++
                $result.Messages += "Recovered unfinished run $($journalObject.RunId)"
            } else {
                $script:MutationJournal.Status = "RecoveryFailed"
                [void](Save-MutationJournal -Journal $script:MutationJournal)
                $result.Failed++
                $result.Messages += "Recovery verification failed for run $($journalObject.RunId)"
            }
        } catch {
            $result.Failed++
            $result.Messages += "Could not recover '$($file.Name)': $($_.Exception.Message)"
        } finally {
            $script:MutationJournal = $null
        }
    }
    return $result
}

# ============================================================================
# LOG ROTATION
# ============================================================================

function Get-EvidenceArtifactSize {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    try {
        $item = Get-Item -LiteralPath $Path -Force -ErrorAction Stop
        if (-not $item.PSIsContainer) {
            return [PSCustomObject]@{ Bytes = [long]$item.Length; Files = 1 }
        }
        $files = @(Get-ChildItem -LiteralPath $Path -File -Force -Recurse -ErrorAction SilentlyContinue)
        $bytes = [long](($files | Measure-Object -Property Length -Sum).Sum)
        return [PSCustomObject]@{ Bytes = $bytes; Files = $files.Count }
    } catch {
        return [PSCustomObject]@{ Bytes = 0L; Files = 0 }
    }
}

function Get-EvidenceRetentionCandidate {
    $candidates = [System.Collections.ArrayList]::new()
    $excludedPaths = @(
        [IO.Path]::GetFullPath([string]$script:LogFile),
        [IO.Path]::GetFullPath([string]$script:TranscriptFile),
        [IO.Path]::GetFullPath([string]$script:StateFile),
        [IO.Path]::GetFullPath([string]$script:HistoryFile),
        [IO.Path]::GetFullPath("$($script:StateFile).previous"),
        [IO.Path]::GetFullPath("$($script:HistoryFile).previous")
    )

    if (Test-Path -LiteralPath $LogPath -PathType Container) {
        foreach ($item in @(Get-ChildItem -LiteralPath $LogPath -Force -ErrorAction SilentlyContinue)) {
            $owned = if ($item.PSIsContainer) {
                $item.Name -match "^HPIA_\d{8}_\d{6}$"
            } else {
                $item.Name -match "^(SystemUpdatePro_(?:Transcript_|Report_)?|DCU_(?:Scan|Apply)_).*\.(?:log|html)$"
            }
            if (-not $owned -or [IO.Path]::GetFullPath($item.FullName) -in $excludedPaths) { continue }
            $size = Get-EvidenceArtifactSize -Path $item.FullName
            [void]$candidates.Add([PSCustomObject]@{
                Path = $item.FullName
                Kind = $(if ($item.PSIsContainer) { "Directory" } else { "File" })
                LastWriteTime = $item.LastWriteTime
                Bytes = [long]$size.Bytes
                Files = [int]$size.Files
                Category = $(if ($item.PSIsContainer) { "VendorReport" } else { "LogOrReport" })
            })
        }
    }

    $driverBackupRoot = Join-Path $script:DataPath "DriverBackups"
    if (Test-Path -LiteralPath $driverBackupRoot -PathType Container) {
        foreach ($directory in @(Get-ChildItem -LiteralPath $driverBackupRoot -Directory -Force `
            -ErrorAction SilentlyContinue | Where-Object { $_.Name -match "^Backup_\d{8}_\d{6}$" })) {
            $size = Get-EvidenceArtifactSize -Path $directory.FullName
            [void]$candidates.Add([PSCustomObject]@{
                Path = $directory.FullName
                Kind = "Directory"
                LastWriteTime = $directory.LastWriteTime
                Bytes = [long]$size.Bytes
                Files = [int]$size.Files
                Category = "DriverBackup"
            })
        }
    }

    $bundleRoot = Join-Path $script:DataPath "Bundles"
    if (Test-Path -LiteralPath $bundleRoot -PathType Container) {
        foreach ($file in @(Get-ChildItem -LiteralPath $bundleRoot -File -Force `
            -Filter "SystemUpdatePro_Diagnostic_*.zip" -ErrorAction SilentlyContinue)) {
            [void]$candidates.Add([PSCustomObject]@{
                Path = $file.FullName
                Kind = "File"
                LastWriteTime = $file.LastWriteTime
                Bytes = [long]$file.Length
                Files = 1
                Category = "DiagnosticBundle"
            })
        }
    }

    if (Test-Path -LiteralPath $script:WebhookDeliveryDirectory -PathType Container) {
        foreach ($file in @(Get-ChildItem -LiteralPath $script:WebhookDeliveryDirectory `
            -File -Force -ErrorAction SilentlyContinue | Where-Object {
                $_.Name -match "^[A-Za-z0-9._-]+\.json(?:\.previous)?$"
            })) {
            [void]$candidates.Add([PSCustomObject]@{
                Path = $file.FullName
                Kind = "File"
                LastWriteTime = $file.LastWriteTime
                Bytes = [long]$file.Length
                Files = 1
                Category = "WebhookDelivery"
            })
        }
    }

    foreach ($root in @($script:DataPath, $script:MutationJournalDirectory)) {
        if (-not (Test-Path -LiteralPath $root -PathType Container)) { continue }
        foreach ($file in @(Get-ChildItem -LiteralPath $root -File -Force -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match "\.corrupt\.\d{8}T\d{6}\.[a-f0-9]+" })) {
            [void]$candidates.Add([PSCustomObject]@{
                Path = $file.FullName
                Kind = "File"
                LastWriteTime = $file.LastWriteTime
                Bytes = [long]$file.Length
                Files = 1
                Category = "Quarantine"
            })
        }
    }
    return @($candidates)
}

function Remove-EvidenceRetentionCandidate {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseShouldProcessForStateChangingFunctions", "", Justification = "Deletes only a candidate constructed from product-owned filename and directory contracts.")]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Candidate,
        [Parameter(Mandatory = $true)]
        [System.Collections.IDictionary]$Result,
        [string]$Reason
    )

    try {
        if ($Candidate.Kind -eq "Directory") {
            Remove-Item -LiteralPath $Candidate.Path -Recurse -Force -ErrorAction Stop
            $Result["DeletedDirectories"] = [int]$Result.DeletedDirectories + 1
        } else {
            Remove-Item -LiteralPath $Candidate.Path -Force -ErrorAction Stop
        }
        $Result["DeletedFiles"] = [int]$Result.DeletedFiles + [int]$Candidate.Files
        $Result["BytesFreed"] = [long]$Result.BytesFreed + [long]$Candidate.Bytes
        $Result["Items"] = @($Result.Items) + @([ordered]@{
            Path = [string]$Candidate.Path
            Category = [string]$Candidate.Category
            Reason = $Reason
            Files = [int]$Candidate.Files
            Bytes = [long]$Candidate.Bytes
        })
        return $true
    } catch {
        $Result["Errors"] = @($Result.Errors) + @(
            "$(Protect-EvidenceText -Text $Candidate.Path): $($_.Exception.Message)"
        )
        return $false
    }
}

function Invoke-EvidenceRetention {
    param(
        [int]$RetentionDays = $LogRetentionDays,
        [int]$MaximumSizeMB = $EvidenceMaxSizeMB
    )

    $result = [ordered]@{
        SchemaVersion = 1
        EvaluatedAt = (Get-Date).ToUniversalTime().ToString("o")
        RetentionDays = [math]::Max(1, $RetentionDays)
        MaximumSizeMB = [math]::Max(10, $MaximumSizeMB)
        DeletedFiles = 0
        DeletedDirectories = 0
        BytesFreed = 0L
        RemainingBytes = 0L
        Items = @()
        Errors = @()
    }
    $cutoff = (Get-Date).AddDays(-[int]$result.RetentionDays)
    $candidates = @(Get-EvidenceRetentionCandidate)
    $removedPaths = [System.Collections.ArrayList]::new()

    foreach ($candidate in @($candidates | Where-Object {
        $_.LastWriteTime -lt $cutoff
    } | Sort-Object LastWriteTime)) {
        if (Remove-EvidenceRetentionCandidate -Candidate $candidate -Result $result -Reason "Age") {
            [void]$removedPaths.Add([string]$candidate.Path)
        }
    }

    $remainingBackups = @($candidates | Where-Object {
        $_.Category -eq "DriverBackup" -and [string]$_.Path -notin @($removedPaths)
    } | Sort-Object LastWriteTime -Descending)
    foreach ($candidate in @($remainingBackups | Select-Object -Skip 3)) {
        if (Remove-EvidenceRetentionCandidate -Candidate $candidate -Result $result -Reason "BackupCount") {
            [void]$removedPaths.Add([string]$candidate.Path)
        }
    }

    $remaining = @($candidates | Where-Object {
        [string]$_.Path -notin @($removedPaths) -and (Test-Path -LiteralPath $_.Path)
    })
    $maximumBytes = [long]$result.MaximumSizeMB * 1MB
    $remainingBytes = [long](($remaining | Measure-Object -Property Bytes -Sum).Sum)
    foreach ($candidate in @($remaining | Sort-Object LastWriteTime)) {
        if ($remainingBytes -le $maximumBytes) { break }
        if (Remove-EvidenceRetentionCandidate -Candidate $candidate -Result $result -Reason "Size") {
            $remainingBytes -= [long]$candidate.Bytes
            [void]$removedPaths.Add([string]$candidate.Path)
        }
    }
    $result["RemainingBytes"] = [math]::Max(0L, $remainingBytes)
    $script:RetentionResult = $result

    $deletedArtifacts = [int]$result.DeletedFiles + [int]$result.DeletedDirectories
    if ($deletedArtifacts -gt 0 -or @($result.Errors).Count -gt 0) {
        Write-Log (
            "Evidence retention removed $($result.DeletedFiles) files and " +
            "$($result.DeletedDirectories) directories, freed $($result.BytesFreed) bytes, " +
            "and recorded $(@($result.Errors).Count) errors"
        ) $(if (@($result.Errors).Count -gt 0) { "WARNING" } else { "DEBUG" })
    }
    return $result
}

function Invoke-LogRotation {
    param([int]$RetentionDays = 30)

    return Invoke-EvidenceRetention -RetentionDays $RetentionDays `
        -MaximumSizeMB $EvidenceMaxSizeMB
}

# ============================================================================
# UPDATE HISTORY TRACKING
# ============================================================================

function Convert-HistorySchema {
    param(
        [AllowNull()][object]$History
    )

    $converted = ConvertTo-Hashtable -InputObject $History
    if ($converted -is [System.Collections.IDictionary] -and
        $converted.Contains("entries") -and [int]$converted.schema_version -eq 2) {
        return $converted
    }

    $legacyEntries = if ($null -eq $converted) {
        @()
    } elseif ($converted -is [System.Collections.IDictionary] -and $converted.Contains("entries")) {
        @($converted.entries)
    } else {
        @($converted)
    }
    $migratedEntries = [System.Collections.ArrayList]::new()
    foreach ($legacyEntry in $legacyEntries) {
        $entry = ConvertTo-Hashtable -InputObject $legacyEntry
        if ($entry -isnot [System.Collections.IDictionary]) { continue }
        if ([string]::IsNullOrWhiteSpace([string]$entry.run_id)) {
            $entry["run_id"] = [guid]::NewGuid().ToString()
        }
        if (-not $entry.Contains("schema_version")) { $entry["schema_version"] = 1 }
        if (-not $entry.Contains("stages")) { $entry["stages"] = @() }
        if (-not $entry.Contains("dependencies")) { $entry["dependencies"] = @() }
        if (-not $entry.Contains("mutation_recovery")) { $entry["mutation_recovery"] = @() }
        if (-not $entry.Contains("capabilities")) { $entry["capabilities"] = $null }
        if (-not $entry.Contains("evidence_delivery")) { $entry["evidence_delivery"] = [ordered]@{} }
        [void]$migratedEntries.Add($entry)
    }

    return [ordered]@{
        schema_version = $script:HistorySchemaVersion
        _migration_source_schema = $(if ($converted -is [System.Collections.IDictionary] -and
            $converted.Contains("schema_version")) { [int]$converted.schema_version } else { 0 })
        last_updated_at = (Get-Date).ToUniversalTime().ToString("o")
        entries = @($migratedEntries)
    }
}

function Test-HistoryDocument {
    param(
        [AllowNull()][object]$History
    )

    if ($History -isnot [System.Collections.IDictionary]) {
        return [PSCustomObject]@{ Valid = $false; Reason = "History root is not an object" }
    }
    if ([int]$History.schema_version -ne $script:HistorySchemaVersion) {
        return [PSCustomObject]@{ Valid = $false; Reason = "Unsupported history schema version" }
    }
    if (-not $History.Contains("entries")) {
        return [PSCustomObject]@{ Valid = $false; Reason = "History entries are missing" }
    }
    foreach ($entry in @($History.entries)) {
        if ($entry -isnot [System.Collections.IDictionary]) {
            return [PSCustomObject]@{ Valid = $false; Reason = "History entry is not an object" }
        }
        $parsedRunId = [guid]::Empty
        if (-not [guid]::TryParse([string]$entry.run_id, [ref]$parsedRunId)) {
            return [PSCustomObject]@{ Valid = $false; Reason = "History entry run ID is invalid" }
        }
    }
    return [PSCustomObject]@{ Valid = $true; Reason = "" }
}

function Read-UpdateHistory {
    $read = Read-ProtectedJsonFile -Path $script:HistoryFile `
        -MigrationScript ${function:Convert-HistorySchema} `
        -ValidationScript ${function:Test-HistoryDocument}
    if (-not $read.Success) {
        $script:LastHistoryError = $read.Error
        return $null
    }
    if ($read.Data.Contains("_migration_source_schema")) {
        [void]$read.Data.Remove("_migration_source_schema")
        $read.Data = Protect-EvidenceObject -InputObject $read.Data
        [void](Write-ProtectedAtomicJson -Path $script:HistoryFile -Data $read.Data -Depth 24 `
            -DataValidationScript ${function:Test-HistoryDocument})
        [void](Write-ProtectedAtomicJson -Path "$($script:HistoryFile).previous" `
            -Data $read.Data -Depth 24 -DataValidationScript ${function:Test-HistoryDocument} `
            -KeepLastKnownGood:$false)
    }
    return $read.Data
}

function Save-UpdateHistory {
    param(
        [hashtable]$RunData
    )

    try {
        $document = Read-UpdateHistory
        if ($null -eq $document) {
            $document = Convert-HistorySchema -History @()
        }
        $history = @($document.entries)

        # A reboot continuation retains the logical run ID. Replace the prior
        # segment so history contains one durable record per logical run.
        $history = @($history | Where-Object { [string]$_.run_id -ne [string]$RunData.RunId })

        $entry = Protect-EvidenceObject -InputObject ([ordered]@{
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
            dependencies      = @($RunData.Dependencies)
            mutation_recovery = @($RunData.MutationRecovery)
            capabilities      = $RunData.Capabilities
            retention         = $RunData.Retention
            dependency_readiness = $RunData.DependencyReadiness
            download_policy   = $RunData.DownloadPolicy
            winget_scopes     = @($RunData.WingetScopes)
            rollout_decision  = $RunData.RolloutDecision
            windows_update_policy = $RunData.WindowsUpdatePolicy
            maintenance_decision = $RunData.MaintenanceDecision
            power_plan_state = $RunData.PowerPlanState
            health = $RunData.Health
            metrics = $RunData.Metrics
            evidence_delivery = $RunData.EvidenceDelivery
            parameters        = Protect-EvidenceObject -InputObject (
                Get-EffectiveRunParameter
            )
        })

        # Prepend new entry, keep last 100 runs
        $history = @($entry) + @($history)
        if ($history.Count -gt 100) {
            $history = $history[0..99]
        }

        $document["schema_version"] = $script:HistorySchemaVersion
        $document["last_updated_at"] = (Get-Date).ToUniversalTime().ToString("o")
        $document["entries"] = @($history)
        if ($document.Contains("_migration_source_schema")) {
            [void]$document.Remove("_migration_source_schema")
        }
        if (-not (Write-ProtectedAtomicJson -Path $script:HistoryFile -Data $document -Depth 24 `
            -DataValidationScript ${function:Test-HistoryDocument})) {
            throw $script:LastEvidenceWriteError
        }
        return $true
    } catch {
        $script:LastHistoryError = $_.Exception.Message
        return $false
    }
}

function Show-UpdateHistory {
    param([int]$Count = 10)

    if (-not (Test-Path -LiteralPath $script:HistoryFile) -and
        -not (Test-Path -LiteralPath "$($script:HistoryFile).previous")) {
        Write-Host "No update history found." -ForegroundColor Yellow
        return
    }

    try {
        $history = Read-UpdateHistory
        if ($null -eq $history -or @($history.entries).Count -eq 0) {
            Write-Host "No update history found." -ForegroundColor Yellow
            return
        }

        $entries = @(@($history.entries) | Select-Object -First $Count)

        Write-Host ""
        Write-Host "  ================================================================" -ForegroundColor Cyan
        Write-Host "    SystemUpdatePro - Update History (Last $(@($entries).Count) Runs)" -ForegroundColor Cyan
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
# SYSTEM RESTORE AND DRIVER ROLLBACK
# ============================================================================

function Get-RestorePointThrottleMinutes {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseSingularNouns", "", Justification = "Uses the Windows System Restore policy's minute-based throttle terminology.")]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingEmptyCatchBlock", "", Justification = "Registry policy is optional; an unavailable policy falls back to the documented default.")]
    param(
        [string]$RegistryPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore"
    )

    $defaultMinutes = 1440
    try {
        $configured = [int](Get-ItemProperty -LiteralPath $RegistryPath `
            -Name "SystemRestorePointCreationFrequency" -ErrorAction Stop).SystemRestorePointCreationFrequency
        if ($configured -ge 0 -and $configured -le 525600) {
            return $configured
        }
    } catch { }
    return $defaultMinutes
}

function Invoke-RestorePointIfNeeded {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseShouldProcessForStateChangingFunctions", "", Justification = "Creates one throttled System Restore point through the Windows restore-point provider.")]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingEmptyCatchBlock", "", Justification = "Optional restore-point telemetry failures must not prevent the guarded fallback path.")]
    param(
        [string]$RunId = [string]$script:RunId,
        [bool]$DryRunMode = [bool]$script:DryRun
    )

    $markerPath = Join-Path $script:DataPath "restore_point.json"
    $throttleMinutes = Get-RestorePointThrottleMinutes
    $result = [ordered]@{
        SchemaVersion = 1
        Success = $false
        Status = "Failed"
        Created = $false
        Attempted = 0
        Installed = 0
        Failed = 0
        Skipped = 0
        DryRun = $DryRunMode
        ThrottleMinutes = $throttleMinutes
        CreatedAt = ""
        Path = $markerPath
        Reason = ""
        Evidence = @()
    }

    $lastCreatedAt = $null
    if (Test-Path -LiteralPath $markerPath -PathType Leaf) {
        try {
            $marker = try {
                ConvertTo-Hashtable -InputObject (Get-Content -LiteralPath $markerPath -Raw | ConvertFrom-Json -ErrorAction Stop)
            } catch {
                $markerRead = Read-ProtectedJsonFile -Path $markerPath
                if ($markerRead.Success) { $markerRead.Data } else { $null }
            }
            $markerCreatedValue = if ($marker -is [System.Collections.IDictionary]) {
                $marker["CreatedAt"]
            } else { $marker.CreatedAt }
            if ($markerCreatedValue -is [datetime]) {
                $markerDate = [datetime]$markerCreatedValue
                $lastCreatedAt = if ($markerDate.Kind -eq [DateTimeKind]::Unspecified) {
                    [datetime]::SpecifyKind($markerDate, [DateTimeKind]::Utc)
                } else { $markerDate.ToUniversalTime() }
            } elseif (-not [string]::IsNullOrWhiteSpace([string]$markerCreatedValue)) {
                $lastCreatedAt = [datetimeoffset]::Parse(
                    [string]$markerCreatedValue,
                    [Globalization.CultureInfo]::InvariantCulture,
                    [Globalization.DateTimeStyles]::RoundtripKind
                ).UtcDateTime
            }
        } catch { }
    }
    if ($null -eq $lastCreatedAt) {
        try {
            $restorePointCommand = Get-Command -Name "Get-ComputerRestorePoint" -ErrorAction SilentlyContinue
            if ($null -ne $restorePointCommand) {
                foreach ($restorePoint in @(Get-ComputerRestorePoint -ErrorAction Stop)) {
                    try {
                        $candidate = [datetime]$restorePoint.CreationTime
                        if ($null -eq $lastCreatedAt -or $candidate -gt $lastCreatedAt) {
                            $lastCreatedAt = $candidate.ToUniversalTime()
                        }
                    } catch { }
                }
            }
        } catch { }
    }

    if ($null -ne $lastCreatedAt -and $throttleMinutes -gt 0) {
        $ageMinutes = ((Get-Date).ToUniversalTime() - $lastCreatedAt).TotalMinutes
        if ($ageMinutes -ge 0 -and $ageMinutes -lt $throttleMinutes) {
            $result.Success = $true
            $result.Status = "Skipped"
            $result.Skipped = 1
            $result.CreatedAt = $lastCreatedAt.ToString("o")
            $result.Reason = "System Restore point creation is throttled for another $([math]::Ceiling($throttleMinutes - $ageMinutes)) minute(s)"
            $result.Evidence = @("last-created:$($lastCreatedAt.ToString('o'))", "throttle-minutes:$throttleMinutes")
            return [PSCustomObject]$result
        }
    }

    if ($DryRunMode) {
        $result.Success = $true
        $result.Status = "Skipped"
        $result.Skipped = 1
        $result.Reason = "Dry run would create a System Restore point before update work"
        $result.Evidence = @("throttle-minutes:$throttleMinutes", "marker:$markerPath")
        return [PSCustomObject]$result
    }

    $checkpointCommand = Get-Command -Name "Checkpoint-Computer" -ErrorAction SilentlyContinue
    if ($null -eq $checkpointCommand) {
        $result.Success = $true
        $result.Status = "Skipped"
        $result.Skipped = 1
        $result.Reason = "System Restore point provider is unavailable on this Windows installation"
        return [PSCustomObject]$result
    }

    try {
        Checkpoint-Computer -Description "SystemUpdatePro $RunId" `
            -RestorePointType "MODIFY_SETTINGS" -ErrorAction Stop
        $createdAt = (Get-Date).ToUniversalTime()
        $markerData = [ordered]@{
            SchemaVersion = 1
            RunId = $RunId
            CreatedAt = $createdAt.ToString("o")
            ThrottleMinutes = $throttleMinutes
        }
        if (-not (Write-ProtectedAtomicJson -Path $markerPath -Data $markerData -Depth 8)) {
            throw "Restore point marker could not be persisted: $script:LastEvidenceWriteError"
        }
        $result.Success = $true
        $result.Status = "Succeeded"
        $result.Created = $true
        $result.Attempted = 1
        $result.CreatedAt = $createdAt.ToString("o")
        $result.Reason = "System Restore point created before update work"
        $result.Evidence = @("marker:$markerPath", "throttle-minutes:$throttleMinutes")
    } catch {
        $result.Failed = 1
        $result.Reason = "System Restore point creation failed: $($_.Exception.Message)"
    }
    return [PSCustomObject]$result
}

function Get-LatestDriverBackupPath {
    param(
        [string]$BackupRoot = (Join-Path $script:DataPath "DriverBackups")
    )

    try {
        if (-not (Test-Path -LiteralPath $BackupRoot -PathType Container)) { return "" }
        $candidate = @(Get-ChildItem -LiteralPath $BackupRoot -Directory -Force -ErrorAction Stop |
            Where-Object { $_.Name -match "^Backup_\d{8}_\d{6}$" } |
            Sort-Object LastWriteTime -Descending | Select-Object -First 1)
        if ($candidate.Count -eq 0) { return "" }
        return [string]$candidate[0].FullName
    } catch {
        return ""
    }
}

function Invoke-DriverRollback {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseShouldProcessForStateChangingFunctions", "", Justification = "Restores only a protected, explicitly selected driver backup through DISM.")]
    param(
        [string]$BackupPath = "",
        [bool]$DryRunMode = [bool]$script:DryRun
    )

    $result = [ordered]@{
        SchemaVersion = 1
        Success = $false
        Status = "Failed"
        Path = ""
        Command = ""
        Attempted = 0
        Installed = 0
        Failed = 0
        Skipped = 0
        RebootRequired = $false
        DryRun = $DryRunMode
        Reason = ""
        Evidence = @()
    }
    $approvedRoot = [IO.Path]::GetFullPath((Join-Path $script:DataPath "DriverBackups")).TrimEnd("\", "/")
    try {
        $selectedPath = if ([string]::IsNullOrWhiteSpace($BackupPath)) {
            Get-LatestDriverBackupPath -BackupRoot $approvedRoot
        } else {
            [IO.Path]::GetFullPath($BackupPath).TrimEnd("\", "/")
        }
        if ([string]::IsNullOrWhiteSpace($selectedPath)) {
            throw "No protected driver backup is available"
        }
        if (-not $selectedPath.StartsWith("$approvedRoot$([IO.Path]::DirectorySeparatorChar)", [StringComparison]::OrdinalIgnoreCase)) {
            throw "Driver rollback path must remain under the protected DriverBackups directory"
        }
        if (-not (Test-Path -LiteralPath $selectedPath -PathType Container)) {
            throw "Driver backup path does not exist: $selectedPath"
        }
        $infFiles = @(Get-ChildItem -LiteralPath $selectedPath -Recurse -File -Filter "*.inf" -ErrorAction Stop)
        if ($infFiles.Count -eq 0) {
            throw "Driver backup contains no INF packages"
        }
        $result.Path = $selectedPath
        $result.Attempted = $infFiles.Count
        $result.Command = "dism.exe /Online /Add-Driver /Driver:$selectedPath /Recurse"
        if ($DryRunMode) {
            $result.Success = $true
            $result.Status = "Skipped"
            $result.Skipped = 1
            $result.Reason = "Dry run would restore $($infFiles.Count) driver package(s) from the newest protected backup"
            $result.Evidence = @("driver-count:$($infFiles.Count)", "command:$($result.Command)")
            return [PSCustomObject]$result
        }

        $dismCommand = Get-Command -Name "dism.exe" -ErrorAction SilentlyContinue | Select-Object -First 1
        $dismPath = if ($null -ne $dismCommand) { [string]$dismCommand.Source } else { Join-Path $env:SystemRoot "System32\dism.exe" }
        if (-not (Test-Path -LiteralPath $dismPath -PathType Leaf)) {
            throw "DISM executable was not found"
        }
        $commandResult = Invoke-CapturedCommand -FilePath $dismPath -ArgumentList @(
            "/Online", "/Add-Driver", "/Driver:$selectedPath", "/Recurse"
        ) -TimeoutSeconds 900
        $exitCode = [int]$commandResult.ExitCode
        $result.RebootRequired = ($exitCode -eq 3010 -or [string]$commandResult.StandardOutput -match "(?i)restart|reboot")
        $result.Success = ($exitCode -in @(0, 3010))
        $result.Status = if ($result.Success) { "Succeeded" } else { "Failed" }
        if ($result.Success) {
            $result.Reason = "DISM restored $($infFiles.Count) driver package(s) from the protected backup"
        } else {
            $result.Failed = 1
            $detail = if (-not [string]::IsNullOrWhiteSpace([string]$commandResult.StandardError)) {
                [string]$commandResult.StandardError
            } else { "exit code $exitCode" }
            $result.Reason = "DISM driver rollback failed: $detail"
        }
        $result.Evidence = @("exit-code:$exitCode", "driver-count:$($infFiles.Count)", "command:$($result.Command)")
    } catch {
        $result.Failed = [math]::Max(1, [int]$result.Failed)
        $result.Reason = "Driver rollback failed: $($_.Exception.Message)"
    }
    return [PSCustomObject]$result
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
        if ($DryRun) {
            Write-Log "Would export drivers via DISM /Online /Export-Driver (or Export-WindowsDriver fallback) to $backupDir" "INFO"
            # Count installed third-party drivers for dry-run info
            try {
                $driverCount = @(Get-WindowsDriver -Online -ErrorAction SilentlyContinue | Where-Object { $_.OriginalFileName -notmatch '\\windows\\' }).Count
                Write-Log "Found $driverCount third-party drivers that would be backed up" "INFO"
            } catch {
                Write-Log "Driver enumeration not available in this context" "DEBUG"
            }
            return $backupDir
        }

        if (-not (New-ProtectedDirectory -Path $backupDir)) {
            throw "Driver backup directory could not be protected"
        }

        $dismCommand = Get-Command -Name "dism.exe" -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($null -ne $dismCommand) {
            $dismPath = [string]$dismCommand.Source
            $exportResult = Invoke-CapturedCommand -FilePath $dismPath -ArgumentList @(
                "/Online", "/Export-Driver", "/Destination:$backupDir"
            ) -TimeoutSeconds 900
            if ([int]$exportResult.ExitCode -notin @(0, 3010)) {
                throw "DISM driver export failed with exit code $($exportResult.ExitCode)"
            }
        } else {
            Export-WindowsDriver -Online -Destination $backupDir -ErrorAction Stop | Out-Null
        }

        $driverFiles = @(Get-ChildItem -Path $backupDir -Recurse -File -ErrorAction SilentlyContinue)
        $totalSizeMB = [math]::Round(($driverFiles | Measure-Object -Property Length -Sum).Sum / 1MB, 1)
        $driverFolders = @(Get-ChildItem -Path $backupDir -Directory -ErrorAction SilentlyContinue).Count

        Write-Log "Backed up $driverFolders drivers ($totalSizeMB MB) to $backupDir" "SUCCESS"

        [void](Invoke-EvidenceRetention -RetentionDays $LogRetentionDays `
            -MaximumSizeMB $EvidenceMaxSizeMB)

        return $backupDir
    } catch {
        Write-Log "Driver backup failed: $($_.Exception.Message)" "WARNING"
        return $null
    }
}

# ============================================================================
# SYSTEM HEALTH CHECKS
# ============================================================================

function Invoke-SystemHealthCheck {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseShouldProcessForStateChangingFunctions", "", Justification = "Runs read-only DISM, SFC, and bounded CBS evidence checks.")]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingEmptyCatchBlock", "", Justification = "Unavailable diagnostic evidence is represented as Unknown in the health contract.")]
    param(
        [ValidateSet("PreRun", "PostRun")][string]$Phase = "PreRun",
        [ValidateRange(1, 900)][int]$TimeoutSeconds = 120
    )

    $commands = [System.Collections.ArrayList]::new()
    foreach ($definition in @(
        @{ Name = "DISM"; File = "dism.exe"; Arguments = @("/Online", "/Cleanup-Image", "/CheckHealth") },
        @{ Name = "SFC"; File = "sfc.exe"; Arguments = @("/verifyonly") }
    )) {
        $commandInfo = $null
        try { $commandInfo = Get-Command -Name $definition.File -ErrorAction SilentlyContinue | Select-Object -First 1 } catch { }
        $commandPath = if ($null -ne $commandInfo) {
            if (-not [string]::IsNullOrWhiteSpace([string]$commandInfo.Source)) {
                [string]$commandInfo.Source
            } else { [string]$commandInfo.Path }
        } else {
            Join-Path $script:WindowsRoot ("System32\{0}" -f $definition.File)
        }
        if ([string]::IsNullOrWhiteSpace($commandPath) -or
            -not (Test-Path -LiteralPath $commandPath -PathType Leaf)) {
            [void]$commands.Add([PSCustomObject][ordered]@{
                Name = $definition.Name; File = $definition.File; Status = "Unknown"
                Success = $false; ExitCode = -1; Output = ""; Error = "Executable unavailable"
            })
            continue
        }
        try {
            $commandResult = Invoke-CapturedCommand -FilePath $commandPath `
                -ArgumentList $definition.Arguments -TimeoutSeconds $TimeoutSeconds
            $exitCode = -1
            try { $exitCode = [int]$commandResult.ExitCode } catch { }
            $output = "$(Protect-EvidenceText -Text $commandResult.StandardOutput)`r`n$(Protect-EvidenceText -Text $commandResult.StandardError)"
            $status = if ($exitCode -eq 0) { "Healthy" } else { "Degraded" }
            if ($definition.Name -eq "SFC" -and $output -match "(?i)(found corrupt files|could not repair|integrity violations were found|repair service failed)") {
                $status = "Degraded"
            }
            [void]$commands.Add([PSCustomObject][ordered]@{
                Name = $definition.Name; File = $definition.File; Status = $status
                Success = ($status -eq "Healthy"); ExitCode = $exitCode
                Output = $output.Trim(); Error = [string]$commandResult.Error
            })
        } catch {
            [void]$commands.Add([PSCustomObject][ordered]@{
                Name = $definition.Name; File = $definition.File; Status = "Unknown"
                Success = $false; ExitCode = -1; Output = ""
                Error = Protect-EvidenceText -Text $_.Exception.Message
            })
        }
    }

    $cbsPath = Join-Path $script:WindowsRoot "Logs\CBS\CBS.log"
    $cbs = [ordered]@{
        Path = $cbsPath; Status = "Unknown"; ErrorCount = 0; Truncated = $false; Reason = "CBS.log was not available"
    }
    if (Test-Path -LiteralPath $cbsPath -PathType Leaf) {
        $cbsRead = Read-BoundedEvidenceText -Path $cbsPath -MaximumBytes 1048576
        if ($cbsRead.Success) {
            $cbsText = [string]$cbsRead.Text
            $cbsMatches = @([regex]::Matches($cbsText, "(?im)(cannot repair|repair.*failed|error 0x[0-9a-f]+|corrupt component)"))
            $cbs.Status = if ($cbsMatches.Count -gt 0) { "Degraded" } else { "Healthy" }
            $cbs.ErrorCount = $cbsMatches.Count
            $cbs.Truncated = [bool]$cbsRead.Truncated
            $cbs.Reason = if ($cbsMatches.Count -gt 0) {
                "CBS.log contains $($cbsMatches.Count) bounded repair-error indicator(s)"
            } else { "CBS.log contains no bounded repair-error indicators" }
        } else {
            $cbs.Reason = "CBS.log could not be read: $($cbsRead.Error)"
        }
    }

    $degraded = @($commands | Where-Object { $_.Status -eq "Degraded" }).Count -gt 0 -or
        $cbs.Status -eq "Degraded"
    $unknown = @($commands | Where-Object { $_.Status -eq "Unknown" }).Count -gt 0 -or
        $cbs.Status -eq "Unknown"
    $overallStatus = if ($degraded) { "Degraded" } elseif ($unknown) { "Unknown" } else { "Healthy" }
    $reason = switch ($overallStatus) {
        "Healthy" { "$Phase health checks completed without a detected servicing-integrity problem" }
        "Degraded" { "$Phase health checks detected a servicing-integrity problem" }
        default { "$Phase health checks were incomplete because one or more evidence sources were unavailable" }
    }
    return [PSCustomObject][ordered]@{
        SchemaVersion = 1
        Phase = $Phase
        Status = $overallStatus
        Success = ($overallStatus -ne "Degraded")
        Commands = @($commands)
        CBS = [PSCustomObject]$cbs
        Attempted = @($commands | Where-Object { $_.Status -ne "Unknown" }).Count
        Failed = @($commands | Where-Object { $_.Status -eq "Degraded" }).Count + [int]($cbs.Status -eq "Degraded")
        Skipped = @($commands | Where-Object { $_.Status -eq "Unknown" }).Count + [int]($cbs.Status -eq "Unknown")
        Reason = $reason
        Evidence = @(
            $commands | ForEach-Object { "$($_.Name):$($_.Status):exit=$($_.ExitCode)" }
            "CBS:$($cbs.Status):errors=$($cbs.ErrorCount):truncated=$($cbs.Truncated)"
        )
    }
}

function Compare-SystemHealthRegression {
    param(
        [AllowNull()][object]$Before,
        [AllowNull()][object]$After
    )

    $beforeStatus = [string](Get-ResultValue -Result $Before -Names @("Status") -Default "Unknown")
    $afterStatus = [string](Get-ResultValue -Result $After -Names @("Status") -Default "Unknown")
    $regressed = $beforeStatus -eq "Healthy" -and $afterStatus -eq "Degraded"
    $comparisonStatus = if ($regressed) { "Regressed" } elseif ($beforeStatus -eq "Healthy" -and $afterStatus -eq "Unknown") {
        "Indeterminate"
    } elseif ($afterStatus -eq "Healthy") { "Healthy" } else { "NotComparable" }
    return [PSCustomObject][ordered]@{
        SchemaVersion = 1
        Before = $beforeStatus
        After = $afterStatus
        Regressed = $regressed
        Status = $comparisonStatus
        Reason = if ($regressed) {
            "Post-run servicing health regressed from Healthy to Degraded"
        } elseif ($comparisonStatus -eq "Indeterminate") {
            "Post-run health could not be compared because the post-run check was incomplete"
        } else { "No servicing health regression was established" }
    }
}

# ============================================================================
# PARALLEL READ-ONLY PLANNING
# ============================================================================

function Get-ParallelUpdateSafety {
    param(
        [bool]$DryRunMode = [bool]$script:DryRun,
        [bool]$SkipOEMMode = [bool]$SkipOEM,
        [bool]$SkipWindowsMode = [bool]$SkipWindows,
        [bool]$IncludeFirmware = [bool]$IncludeBIOS,
        [bool]$Repair = [bool]$RepairWindowsUpdate,
        [bool]$Bypass = [bool]$BypassWSUS,
        [bool]$Cleanup = [bool]$CleanupAfter,
        [bool]$ResetBase = [bool]$ResetComponentBase,
        [bool]$Backup = [bool]$BackupDrivers,
        [bool]$Rollback = [bool]$RollbackDrivers,
        [bool]$PreStageMode = [bool]$PreStage,
        [bool]$ContinueMode = [bool]$ContinueAfterReboot
    )

    $reasons = [System.Collections.ArrayList]::new()
    if (-not $DryRunMode) { [void]$reasons.Add("parallel execution is limited to non-installing dry-run plans") }
    if ($SkipOEMMode -or $SkipWindowsMode) { [void]$reasons.Add("both OEM and Windows Update plans must be enabled") }
    if ($IncludeFirmware) { [void]$reasons.Add("firmware applicability and installation remain serial") }
    if ($Repair -or $Bypass) { [void]$reasons.Add("servicing repair or WSUS policy mutation was requested") }
    if ($Cleanup -or $ResetBase) { [void]$reasons.Add("component cleanup must remain serial") }
    if ($Backup -or $Rollback) { [void]$reasons.Add("driver backup or rollback was requested") }
    if ($PreStageMode) { [void]$reasons.Add("pre-stage plan persistence must remain serial") }
    if ($ContinueMode) { [void]$reasons.Add("reboot continuation coordination must remain serial") }
    return [PSCustomObject][ordered]@{
        SchemaVersion = 1
        Allowed = ($reasons.Count -eq 0)
        Mode = if ($reasons.Count -eq 0) { "ReadOnlyParallelPlan" } else { "Serial" }
        Reasons = @($reasons)
        Reason = if ($reasons.Count -eq 0) {
            "OEM and Windows Update discovery can run concurrently without installation or persistent mutation"
        } else { $reasons -join "; " }
    }
}

function Invoke-ParallelReadOnlyPlans {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseShouldProcessForStateChangingFunctions", "", Justification = "Runs two explicitly supplied read-only plan scriptblocks in isolated runspaces.")]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseSingularNouns", "", Justification = "The function executes the two plan contracts as a single parallel operation.")]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingEmptyCatchBlock", "", Justification = "Runspace cleanup and optional function serialization failures are isolated before returning the plan contract.")]
    param(
        [Parameter(Mandatory = $true)][scriptblock]$OEMPlan,
        [Parameter(Mandatory = $true)][scriptblock]$WindowsPlan
    )

    $functionDefinitions = [ordered]@{}
    foreach ($functionCommand in @(Get-Command -CommandType Function -ErrorAction SilentlyContinue)) {
        if ([string]$functionCommand.Name -match "^(prompt|TabExpansion|TabExpansion2)$") { continue }
        try { $functionDefinitions[[string]$functionCommand.Name] = [string]$functionCommand.Definition } catch { }
    }
    $variableNames = @(
        "DataPath", "WindowsRoot", "ProductName", "RunId", "DryRun", "SkipOEM", "SkipWindows", "SkipWinget",
        "IncludeBIOS", "BypassWSUS", "RepairWindowsUpdate", "CleanupAfter", "ResetComponentBase",
        "BackupDrivers", "RollbackDrivers", "MaxRetries", "MaxUpdatePasses", "MinDiskSpaceGB",
        "MinFirmwareChargePercent", "LogPath", "LogRetentionDays", "EvidenceMaxSizeMB", "RedactionMode",
        "Offline", "DependencyCachePath", "SourceTimeoutSeconds", "AllowMeteredNetwork", "PolicyPath",
        "RolloutPolicyPath", "FeatureDeferralDays", "SecurityOnly", "PreStage", "Interactive",
        "CurrentSystemInfo", "FirmwarePrerequisites", "CapabilityAssessment", "AcquisitionManifest",
        "AcquisitionManifestVersion", "DependencyReadiness", "DownloadPolicy", "WindowsUpdatePolicy",
        "PackagePolicy", "RolloutPolicy", "RolloutDecision", "PSModuleInstallRoot", "HPIAInstallRoot",
        "MutationJournalDirectory", "MutationJournal", "MutationEvidence", "ProtectedEvidenceDirectories",
        "SensitiveEvidenceValues", "WebhookUrl", "WebhookSecretReference", "MaxContinuationAttempts",
        "StageResults", "Errors", "Warnings", "OEMUpdates", "WindowsUpdates", "WingetUpdates",
        "RunStartedAt", "EventLogSource"
    )
    $variableValues = [ordered]@{}
    foreach ($name in $variableNames) {
        $variable = Get-Variable -Scope Script -Name $name -ErrorAction SilentlyContinue
        if ($null -ne $variable) { $variableValues[$name] = $variable.Value }
    }

    $workerScript = {
        param([string]$ActionText, [System.Collections.IDictionary]$Definitions, [System.Collections.IDictionary]$Variables)
        foreach ($definition in $Definitions.GetEnumerator()) {
            try { . ([scriptblock]::Create([string]$definition.Value)) } catch { }
        }
        foreach ($variable in $Variables.GetEnumerator()) {
            Set-Variable -Name ([string]$variable.Key) -Scope Script -Value $variable.Value -Force
        }
        $script:DryRun = [switch]$true
        Set-Variable -Name "DryRun" -Scope Script -Value ([switch]$true) -Force
        try {
            $value = & ([scriptblock]::Create($ActionText))
            return [PSCustomObject]@{ Success = $true; Value = $value; Error = "" }
        } catch {
            return [PSCustomObject]@{ Success = $false; Value = $null; Error = (Protect-EvidenceText -Text $_.Exception.Message) }
        }
    }

    $pool = $null
    $jobs = @()
    try {
        $pool = [runspacefactory]::CreateRunspacePool(1, 2)
        $pool.Open()
        foreach ($plan in @(
            @{ Name = "OEM"; Action = $OEMPlan },
            @{ Name = "Windows"; Action = $WindowsPlan }
        )) {
            $powerShell = [powershell]::Create()
            $powerShell.RunspacePool = $pool
            [void]$powerShell.AddScript($workerScript.ToString()).AddArgument(
                ([string]$plan.Action.ToString())
            ).AddArgument($functionDefinitions).AddArgument($variableValues)
            $jobs += [PSCustomObject]@{
                Name = $plan.Name; PowerShell = $powerShell; Handle = $powerShell.BeginInvoke()
            }
        }
        foreach ($job in $jobs) {
            $job.Handle.AsyncWaitHandle.WaitOne() | Out-Null
        }
        $results = [ordered]@{}
        $errors = [System.Collections.ArrayList]::new()
        foreach ($job in $jobs) {
            try {
                $output = @($job.PowerShell.EndInvoke($job.Handle))
                $workerResult = $output | Select-Object -Last 1
                $results[$job.Name] = $workerResult
                if ($null -eq $workerResult -or -not [bool]$workerResult.Success) {
                    [void]$errors.Add("$($job.Name) plan failed: $([string]$workerResult.Error)")
                }
            } catch {
                [void]$errors.Add("$($job.Name) plan runspace failed: $($_.Exception.Message)")
                $results[$job.Name] = [PSCustomObject]@{ Success = $false; Value = $null; Error = $_.Exception.Message }
            }
        }
        return [PSCustomObject][ordered]@{
            SchemaVersion = 1
            Success = ($errors.Count -eq 0)
            Parallel = $true
            OEM = $results["OEM"]
            Windows = $results["Windows"]
            Errors = @($errors)
            Reason = if ($errors.Count -eq 0) { "Read-only OEM and Windows Update plans completed concurrently" } else { $errors -join "; " }
        }
    } catch {
        $serialResults = [ordered]@{}
        $serialErrors = [System.Collections.ArrayList]::new()
        foreach ($plan in @(
            @{ Name = "OEM"; Action = $OEMPlan },
            @{ Name = "Windows"; Action = $WindowsPlan }
        )) {
            try { $serialResults[$plan.Name] = [PSCustomObject]@{ Success = $true; Value = (& $plan.Action); Error = "" } }
            catch {
                [void]$serialErrors.Add("$($plan.Name) serial fallback failed: $($_.Exception.Message)")
                $serialResults[$plan.Name] = [PSCustomObject]@{ Success = $false; Value = $null; Error = $_.Exception.Message }
            }
        }
        return [PSCustomObject][ordered]@{
            SchemaVersion = 1
            Success = ($serialErrors.Count -eq 0)
            Parallel = $false
            OEM = $serialResults["OEM"]
            Windows = $serialResults["Windows"]
            Errors = @($serialErrors)
            Reason = if ($serialErrors.Count -eq 0) {
                "Parallel runspace setup was unavailable; read-only plans completed serially"
            } else { $serialErrors -join "; " }
        }
    } finally {
        foreach ($job in $jobs) {
            try { $job.PowerShell.Dispose() } catch { }
        }
        if ($null -ne $pool) {
            try { $pool.Close(); $pool.Dispose() } catch { }
        }
    }
}

function Invoke-ParallelOEMWindowsUpdatePlan {
    param(
        [Parameter(Mandatory = $true)][object]$SystemInfo,
        [int]$MaxPasses = [int]$MaxUpdatePasses
    )

    $script:CurrentSystemInfo = $SystemInfo
    $script:MaxUpdatePasses = $MaxPasses
    $oemPlan = {
        $sysInfo = $script:CurrentSystemInfo
        $manufacturer = [string]$sysInfo.Manufacturer
        $additional = Invoke-AdditionalOEMUpdate -SystemInfo $sysInfo -IncludeGPU
        $result = $null
        if ($manufacturer -match "DELL|ALIENWARE") {
            $result = Invoke-DellUpdate -IncludeBIOS:$false -SystemInfo $sysInfo -FirmwarePrerequisites $script:FirmwarePrerequisites
        } elseif ($manufacturer -match "LENOVO") {
            $result = Invoke-LenovoUpdate -IncludeBIOS:$false -SystemInfo $sysInfo -FirmwarePrerequisites $script:FirmwarePrerequisites
        } elseif ($manufacturer -match "HP|HEWLETT") {
            $result = Invoke-HPUpdate -IncludeBIOS:$false -SystemInfo $sysInfo -FirmwarePrerequisites $script:FirmwarePrerequisites
        } else {
            $result = $additional
        }
        if ($null -ne $result -and $result -ne $additional -and @($additional.Items).Count -gt 0) {
            $result.Items = @($result.Items) + @($additional.Items)
            $result.Available = [int]$result.Available + [int]$additional.Available
            $result.Attempted = [int]$result.Attempted + [int]$additional.Attempted
            $result.Failed = [int]$result.Failed + [int]$additional.Failed
            $result.Skipped = [int]$result.Skipped + [int]$additional.Skipped
            $result.RebootRequired = [bool]$result.RebootRequired -or [bool]$additional.RebootRequired
        }
        [ordered]@{ Result = $result; ItemNames = @($result.Items | ForEach-Object { [string]$_.Name }) }
    }
    $windowsPlan = {
        $result = Invoke-WindowsUpdate -MaxPasses $script:MaxUpdatePasses
        [ordered]@{ Result = $result; ItemNames = @($result.Items | ForEach-Object { [string]$_.Name }) }
    }
    return Invoke-ParallelReadOnlyPlans -OEMPlan $oemPlan -WindowsPlan $windowsPlan
}

# ============================================================================
# SYSTEM CHECKS
# ============================================================================

function Get-MaintenanceWindowPolicy {
    param([string]$Path = [string]$script:PolicyPath)

    $policy = [ordered]@{
        SchemaVersion = 1; Enabled = $false; Start = ""; End = ""; Days = @()
        Source = "NotConfigured"; Reason = "No maintenance window is configured"
        IntuneDetected = $false
    }
    try {
        $document = Get-PolicyDocument -Path $Path
        $window = if ($document.Contains("maintenance_window") -and
            $document.maintenance_window -is [System.Collections.IDictionary]) {
            $document.maintenance_window
        } else { $null }
        if ($null -ne $window) {
            $policy.Enabled = if ($window.Contains("enabled")) { [bool]$window.enabled } else { $true }
            $policy.Start = [string](Get-ResultValue -Result $window -Names @("start", "Start") -Default "")
            $policy.End = [string](Get-ResultValue -Result $window -Names @("end", "End") -Default "")
            $policy.Days = @($window.days | ForEach-Object { [string]$_ })
            $policy.Source = "Policy"
            $policy.Reason = "Protected maintenance-window policy"
            return [PSCustomObject]$policy
        }
    } catch {
        $policy.Source = "PolicyError"
        $policy.Reason = Protect-EvidenceText -Text $_.Exception.Message
        return [PSCustomObject]$policy
    }

    foreach ($registryPath in @(
        "HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device\Update",
        "HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device\UpdateSchedule",
        "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"
    )) {
        try {
            if (-not (Test-Path -LiteralPath $registryPath)) { continue }
            $values = Get-ItemProperty -LiteralPath $registryPath -ErrorAction SilentlyContinue
            $start = Get-ResultValue -Result $values -Names @("MaintenanceWindowStart", "ActiveHoursStart", "UpdateWindowStart") -Default $null
            $end = Get-ResultValue -Result $values -Names @("MaintenanceWindowEnd", "ActiveHoursEnd", "UpdateWindowEnd") -Default $null
            if ($null -ne $start -and $null -ne $end) {
                $policy.Enabled = $true
                $policy.Start = [string]$start
                $policy.End = [string]$end
                $policy.Source = "Intune"
                $policy.IntuneDetected = $true
                $policy.Reason = "Maintenance window detected from $registryPath"
                return [PSCustomObject]$policy
            }
        } catch { continue }
    }
    return [PSCustomObject]$policy
}

function Test-MaintenanceWindow {
    param(
        [AllowNull()][object]$Policy = $null,
        [datetime]$Now = (Get-Date)
    )

    if ($null -eq $Policy) { $Policy = Get-MaintenanceWindowPolicy }
    if (-not [bool]$Policy.Enabled) {
        return [PSCustomObject][ordered]@{
            SchemaVersion = 1; Allowed = $true; Status = "NotConfigured"; Source = [string]$Policy.Source
            Reason = [string]$Policy.Reason; EvaluatedAt = $Now.ToString("o"); NextWindow = ""
        }
    }
    $start = [TimeSpan]::Zero
    $end = [TimeSpan]::Zero
    if (-not [TimeSpan]::TryParse([string]$Policy.Start, [ref]$start) -or
        -not [TimeSpan]::TryParse([string]$Policy.End, [ref]$end)) {
        return [PSCustomObject][ordered]@{
            SchemaVersion = 1; Allowed = $false; Status = "Unknown"; Source = [string]$Policy.Source
            Reason = "Maintenance window start/end is invalid"; EvaluatedAt = $Now.ToString("o"); NextWindow = ""
        }
    }
    $days = @($Policy.Days | ForEach-Object { [string]$_ } | Where-Object { $_ })
    $dayAllowed = $days.Count -eq 0 -or @($days | Where-Object {
        $_ -eq $Now.DayOfWeek.ToString() -or $_ -eq ([int]$Now.DayOfWeek).ToString()
    }).Count -gt 0
    $current = $Now.TimeOfDay
    $timeAllowed = if ($start -eq $end) { $true } elseif ($start -lt $end) {
        $current -ge $start -and $current -lt $end
    } else {
        $current -ge $start -or $current -lt $end
    }
    if ($dayAllowed -and $timeAllowed) {
        return [PSCustomObject][ordered]@{
            SchemaVersion = 1; Allowed = $true; Status = "Allowed"; Source = [string]$Policy.Source
            Reason = "Current time is inside the configured maintenance window"
            EvaluatedAt = $Now.ToString("o"); NextWindow = ""
        }
    }
    $next = $Now.Date.AddDays(1).Add($start)
    if ($dayAllowed -and $start -gt $current) { $next = $Now.Date.Add($start) }
    return [PSCustomObject][ordered]@{
        SchemaVersion = 1; Allowed = $false; Status = "Blocked"; Source = [string]$Policy.Source
        Reason = "Current time is outside the configured maintenance window"
        EvaluatedAt = $Now.ToString("o"); NextWindow = $next.ToString("o")
    }
}

function Get-StaggeredRebootDecision {
    param(
        [AllowNull()][object]$Policy = $null,
        [string]$DeviceIdentity = [string]$env:COMPUTERNAME,
        [datetime]$Now = (Get-Date)
    )

    if ($null -eq $Policy) {
        try {
            $document = Get-PolicyDocument -Path ([string]$script:PolicyPath)
            $Policy = if ($document.Contains("cluster") -and $document.cluster -is [System.Collections.IDictionary]) {
                $document.cluster
            } elseif ($document.Contains("cluster_reboot") -and $document.cluster_reboot -is [System.Collections.IDictionary]) {
                $document.cluster_reboot
            } else { $null }
        } catch { $Policy = $null }
    }
    if ($null -eq $Policy -or (-not $Policy.Contains("enabled") -and -not $Policy.Contains("max_concurrent"))) {
        return [PSCustomObject][ordered]@{
            SchemaVersion = 1; Enabled = $false; Allowed = $true; Status = "NotConfigured"
            Group = ""; Slot = 0; MaxConcurrent = 0; ActiveLeases = 0; Reason = "Cluster reboot coordination is not configured"
            NextEligibleAt = ""; LeasePath = ""
        }
    }
    $enabled = if ($Policy.Contains("enabled")) { [bool]$Policy.enabled } else { $true }
    if (-not $enabled) {
        return [PSCustomObject][ordered]@{ SchemaVersion = 1; Enabled = $false; Allowed = $true; Status = "Disabled"; Group = ""; Slot = 0; MaxConcurrent = 0; ActiveLeases = 0; Reason = "Cluster reboot coordination is disabled"; NextEligibleAt = ""; LeasePath = "" }
    }
    $group = [string](Get-ResultValue -Result $Policy -Names @("group", "Group") -Default "default")
    $maxConcurrent = 1
    [void][int]::TryParse([string](Get-ResultValue -Result $Policy -Names @("max_concurrent", "MaxConcurrent") -Default 1), [ref]$maxConcurrent)
    $maxConcurrent = [math]::Max(1, [math]::Min(1000, $maxConcurrent))
    $cohort = Get-EndpointCohort -DeviceIdentity "$group|$DeviceIdentity" -Cohort "cluster"
    $slot = [int]$cohort % $maxConcurrent
    $leasePath = [string](Get-ResultValue -Result $Policy -Names @("coordination_path", "CoordinationPath") -Default "")
    $activeLeases = @()
    if (-not [string]::IsNullOrWhiteSpace($leasePath) -and (Test-Path -LiteralPath $leasePath -PathType Leaf)) {
        try {
            $activeLeases = @((Get-Content -LiteralPath $leasePath -Raw | ConvertFrom-Json -ErrorAction Stop).Leases | Where-Object {
                $expires = [datetime]::MinValue
                [datetime]::TryParse([string]$_.ExpiresAt, [ref]$expires) -and $expires -gt $Now
            })
        } catch { $activeLeases = @() }
    }
    if ($activeLeases.Count -ge $maxConcurrent) {
        return [PSCustomObject][ordered]@{
            SchemaVersion = 1; Enabled = $true; Allowed = $false; Status = "Blocked"
            Group = $group; Slot = $slot; MaxConcurrent = $maxConcurrent; ActiveLeases = $activeLeases.Count
            Reason = "Cluster reboot concurrency limit is already in use"
            NextEligibleAt = $Now.AddMinutes(5).ToString("o"); LeasePath = $leasePath
        }
    }
    return [PSCustomObject][ordered]@{
        SchemaVersion = 1; Enabled = $true; Allowed = $true; Status = "Allowed"
        Group = $group; Slot = $slot; MaxConcurrent = $maxConcurrent; ActiveLeases = $activeLeases.Count
        Reason = "Cluster reboot slot is available"; NextEligibleAt = ""; LeasePath = $leasePath
    }
}

function Save-StaggeredRebootLease {
    param(
        [Parameter(Mandatory = $true)][object]$Decision,
        [int]$LeaseMinutes = 30,
        [bool]$DryRunMode = [bool]$script:DryRun
    )
    if (-not [bool]$Decision.Enabled -or [string]$Decision.Status -eq "NotConfigured") {
        return [PSCustomObject]@{ Success = $true; Persisted = $false; Path = ""; Reason = "Coordination is not configured" }
    }
    if (-not [bool]$Decision.Allowed) {
        return [PSCustomObject]@{ Success = $false; Persisted = $false; Path = [string]$Decision.LeasePath; Reason = [string]$Decision.Reason }
    }
    if ($DryRunMode -or [string]::IsNullOrWhiteSpace([string]$Decision.LeasePath)) {
        return [PSCustomObject]@{ Success = $true; Persisted = $false; Path = [string]$Decision.LeasePath; Reason = "Dry run or no shared lease path; reboot decision was recorded in memory" }
    }
    try {
        $directory = Split-Path -Parent ([string]$Decision.LeasePath)
        if (-not (New-ProtectedDirectory -Path $directory)) { throw "Cluster lease directory could not be protected" }
        $leases = @()
        if (Test-Path -LiteralPath $Decision.LeasePath -PathType Leaf) {
            try { $leases = @((Get-Content -LiteralPath $Decision.LeasePath -Raw | ConvertFrom-Json -ErrorAction Stop).Leases) } catch { $leases = @() }
        }
        $now = (Get-Date).ToUniversalTime()
        $leases = @($leases | Where-Object {
            $expires = [datetime]::MinValue
            [datetime]::TryParse([string]$_.ExpiresAt, [ref]$expires) -and $expires.ToUniversalTime() -gt $now
        })
        $leases += [ordered]@{ RunId = [string]$script:RunId; ComputerName = [string]$env:COMPUTERNAME; ExpiresAt = $now.AddMinutes([math]::Max(1, $LeaseMinutes)).ToString("o") }
        $document = [ordered]@{ SchemaVersion = 1; Group = [string]$Decision.Group; Leases = @($leases) }
        if (-not (Write-ProtectedAtomicJson -Path ([string]$Decision.LeasePath) -Data $document -Depth 12)) { throw $script:LastEvidenceWriteError }
        return [PSCustomObject]@{ Success = $true; Persisted = $true; Path = [string]$Decision.LeasePath; Reason = "Cluster reboot lease persisted" }
    } catch {
        return [PSCustomObject]@{ Success = $false; Persisted = $false; Path = [string]$Decision.LeasePath; Reason = "Cluster reboot lease failed: $($_.Exception.Message)" }
    }
}

function Get-PowerPlanState {
    $result = [ordered]@{ SchemaVersion = 1; Success = $false; ActiveScheme = ""; ActiveName = ""; Reason = "" }
    try {
        $output = @(& powercfg.exe /getactivescheme 2>&1)
        $line = (@($output) | Where-Object { [string]$_ -match "[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}" } | Select-Object -First 1)
        if ($line -match "(?<Guid>[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})\s*\((?<Name>[^\)]+)\)") {
            $result.ActiveScheme = $Matches.Guid
            $result.ActiveName = $Matches.Name
            $result.Success = $true
            $result.Reason = "Active power scheme detected"
        } else { $result.Reason = "powercfg did not return an active scheme" }
    } catch { $result.Reason = "Power scheme query failed: $($_.Exception.Message)" }
    return [PSCustomObject]$result
}

function Set-HighPerformancePowerPlan {
    param([bool]$DryRunMode = [bool]$script:DryRun)
    $previous = Get-PowerPlanState
    $highPerformanceGuid = "8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c"
    $result = [ordered]@{
        SchemaVersion = 1; Success = $previous.Success; Changed = $false; DryRun = $DryRunMode
        PreviousScheme = [string]$previous.ActiveScheme; AppliedScheme = $highPerformanceGuid; Reason = [string]$previous.Reason
    }
    if (-not $previous.Success) { return [PSCustomObject]$result }
    if ($previous.ActiveScheme -eq $highPerformanceGuid) {
        $result.Reason = "High performance power scheme is already active"
        return [PSCustomObject]$result
    }
    $result.Changed = $true
    if ($DryRunMode) { $result.Reason = "Would temporarily activate the High performance power scheme"; return [PSCustomObject]$result }
    try {
        $process = Start-Process -FilePath "powercfg.exe" -ArgumentList @("/setactive", $highPerformanceGuid) `
            -Wait -NoNewWindow -PassThru -ErrorAction Stop
        $result.Success = ($process.ExitCode -eq 0)
        $result.Reason = if ($result.Success) { "High performance power scheme activated" } else { "powercfg /setactive exited with $($process.ExitCode)" }
    } catch {
        $result.Success = $false; $result.Reason = "Power scheme activation failed: $($_.Exception.Message)"
    }
    return [PSCustomObject]$result
}

function Restore-PowerPlan {
    param(
        [AllowNull()][object]$State = $script:PowerPlanState,
        [bool]$DryRunMode = [bool]$script:DryRun
    )
    if ($null -eq $State -or -not [bool]$State.Changed -or [string]::IsNullOrWhiteSpace([string]$State.PreviousScheme)) {
        return [PSCustomObject]@{ Success = $true; Restored = $false; Reason = "No temporary power-plan change requires restoration" }
    }
    if ($DryRunMode) {
        return [PSCustomObject]@{ Success = $true; Restored = $false; Reason = "Dry run did not change the power plan" }
    }
    try {
        $process = Start-Process -FilePath "powercfg.exe" -ArgumentList @("/setactive", [string]$State.PreviousScheme) `
            -Wait -NoNewWindow -PassThru -ErrorAction Stop
        return [PSCustomObject]@{ Success = ($process.ExitCode -eq 0); Restored = ($process.ExitCode -eq 0); Reason = if ($process.ExitCode -eq 0) { "Original power scheme restored" } else { "powercfg restore exited with $($process.ExitCode)" } }
    } catch {
        return [PSCustomObject]@{ Success = $false; Restored = $false; Reason = "Power scheme restoration failed: $($_.Exception.Message)" }
    }
}

function Get-DryRunMutationSnapshot {
    $files = [ordered]@{}
    foreach ($path in @($script:StateFile, $script:LockFile, (Join-Path $script:DataPath "WindowsUpdatePrestage.json"), (Join-Path $script:DataPath "WindowsPolicySnapshot.json"), (Join-Path $script:DataPath "restore_point.json"))) {
        $key = [string]$path
        if (Test-Path -LiteralPath $path -PathType Leaf) {
            try { $files[$key] = [ordered]@{ Exists = $true; Length = [long](Get-Item -LiteralPath $path).Length; Hash = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash } } catch { $files[$key] = [ordered]@{ Exists = $true; Length = -1; Hash = "error" } }
        } else { $files[$key] = [ordered]@{ Exists = $false; Length = 0; Hash = "" } }
    }
    $registry = @()
    foreach ($target in @(
        @{ Path = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"; Name = "UseWUServer" }
    )) {
        try { $registry += Get-RegistryValueSnapshot -Path $target.Path -Name $target.Name } catch { $registry += [ordered]@{ Path = $target.Path; Name = $target.Name; Exists = $false; Error = $_.Exception.Message } }
    }
    return [PSCustomObject][ordered]@{ SchemaVersion = 1; Files = $files; Registry = @($registry) }
}

function Compare-DryRunMutationSnapshot {
    param(
        [Parameter(Mandatory = $true)][object]$Before,
        [Parameter(Mandatory = $true)][object]$After
    )
    $changes = [System.Collections.ArrayList]::new()
    foreach ($path in @($Before.Files.Keys)) {
        $beforeJson = ($Before.Files[$path] | ConvertTo-Json -Depth 8 -Compress)
        $afterJson = ($After.Files[$path] | ConvertTo-Json -Depth 8 -Compress)
        if ($beforeJson -ne $afterJson) { [void]$changes.Add("file:$path") }
    }
    $beforeRegistry = @($Before.Registry | ConvertTo-Json -Depth 12 -Compress)
    $afterRegistry = @($After.Registry | ConvertTo-Json -Depth 12 -Compress)
    if (($beforeRegistry -join "") -ne ($afterRegistry -join "")) { [void]$changes.Add("registry:WindowsUpdate") }
    return [PSCustomObject]@{ SchemaVersion = 1; Changed = ($changes.Count -gt 0); Changes = @($changes); Reason = if ($changes.Count) { "Dry-run persistent mutation detected: $($changes -join ', ')" } else { "No tracked persistent system mutation detected" } }
}

function Get-ConsoleColorCapability {
    $windowsPlatform = [Environment]::OSVersion.Platform -eq [PlatformID]::Win32NT
    $supportsColor = $false
    try { $supportsColor = $null -ne $Host.UI.RawUI -and $null -ne $Host.UI.RawUI.ForegroundColor } catch { $supportsColor = $false }
    return [PSCustomObject]@{ SupportsColor = $supportsColor; SupportsVirtualTerminal = ($windowsPlatform -and [int]$PSVersionTable.PSVersion.Major -ge 7); LegacyConsole = ($windowsPlatform -and -not $supportsColor) }
}

function Write-StageProgress {
    param(
        [Parameter(Mandatory = $true)][string]$Stage,
        [ValidateRange(0, 100)][int]$PercentComplete,
        [string]$Status = ""
    )
    try {
        Write-Progress -Id 1 -Activity "SystemUpdatePro" -Status "$Stage $Status" -PercentComplete $PercentComplete
        if ($PercentComplete -ge 100) { Write-Progress -Id 1 -Activity "SystemUpdatePro" -Completed }
    } catch { }
}

function Test-InternetConnection {
    param([string[]]$Endpoints = @())

    $endpoints = if ($Endpoints.Count -gt 0) {
        @($Endpoints)
    } else {
        @(
            "https://www.microsoft.com",
            "https://download.microsoft.com"
        )
    }

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

    try {
        $systemDrive = [string]$env:SystemDrive
        if ([string]::IsNullOrWhiteSpace($systemDrive)) {
            throw "The system drive environment variable is empty"
        }

        $disk = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='$systemDrive'" -ErrorAction Stop
        if ($null -eq $disk -or $null -eq $disk.FreeSpace) {
            throw "CIM did not return free-space data for $systemDrive"
        }

        $freeGB = [math]::Round($disk.FreeSpace / 1GB, 2)
        return @{
            Status = $(if ($freeGB -ge $MinGB) { "Ready" } else { "Blocked" })
            Sufficient = ($freeGB -ge $MinGB)
            FreeGB = $freeGB
            RequiredGB = $MinGB
            Message = $(if ($freeGB -ge $MinGB) {
                "$freeGB GB free on $systemDrive"
            } else {
                "Only $freeGB GB is free on $systemDrive; free at least $MinGB GB before firmware or update installation"
            })
        }
    } catch {
        return @{
            Status = "Unknown"
            Sufficient = $false
            FreeGB = $null
            RequiredGB = $MinGB
            Message = "Disk space could not be verified: $($_.Exception.Message). Repair CIM/WMI and rerun"
        }
    }
}

function Get-SystemPowerStatus {
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        $power = [System.Windows.Forms.SystemInformation]::PowerStatus
        return [PSCustomObject]@{
            PowerLineStatus = [string]$power.PowerLineStatus
            BatteryStatus = [string]$power.BatteryChargeStatus
            BatteryLifePercent = [double]$power.BatteryLifePercent
        }
    } catch {
        throw "Windows power status query failed: $($_.Exception.Message)"
    }
}

function Test-BatteryPower {
    param([int]$MinimumChargePercent = 50)

    try {
        $power = Get-SystemPowerStatus
        $lineStatus = [string]$power.PowerLineStatus
        $batteryStatus = [string]$power.BatteryStatus
        $hasBattery = $batteryStatus -notmatch "(?i)NoSystemBattery"
        $chargePercent = $null
        $lifePercent = [double]$power.BatteryLifePercent
        if ($hasBattery -and $lifePercent -ge 0 -and $lifePercent -le 1) {
            $chargePercent = [int][math]::Round($lifePercent * 100)
        }

        if ($lineStatus -eq "Unknown" -or [string]::IsNullOrWhiteSpace($lineStatus)) {
            return @{
                Status = "Unknown"; HasBattery = $hasBattery; OnBattery = $false; OnACPower = $false
                ChargePercent = $chargePercent
                RequiredChargePercent = $MinimumChargePercent
                Message = "AC power state is unknown. Connect a verified AC adapter and repair the Windows power-status provider"
            }
        }

        if ($lineStatus -eq "Offline") {
            return @{
                Status = "Blocked"; HasBattery = $hasBattery; OnBattery = $true; OnACPower = $false
                ChargePercent = $chargePercent
                RequiredChargePercent = $MinimumChargePercent
                Message = "The system is running on battery power. Connect a verified AC adapter before firmware installation"
            }
        }

        if ($lineStatus -ne "Online") {
            return @{
                Status = "Unknown"; HasBattery = $hasBattery; OnBattery = $false; OnACPower = $false
                ChargePercent = $chargePercent
                RequiredChargePercent = $MinimumChargePercent
                Message = "Unexpected AC power state '$lineStatus'. Verify the adapter and rerun"
            }
        }

        if (-not $hasBattery) {
            return @{
                Status = "Ready"; HasBattery = $false; OnBattery = $false; OnACPower = $true
                ChargePercent = $null
                RequiredChargePercent = $MinimumChargePercent
                Message = "Line power is online and Windows reports no system battery"
            }
        }

        if ($batteryStatus -match "(?i)Unknown" -or $null -eq $chargePercent) {
            return @{
                Status = "Unknown"; HasBattery = $true; OnBattery = $false; OnACPower = $true
                ChargePercent = $chargePercent
                RequiredChargePercent = $MinimumChargePercent
                Message = "Battery charge could not be verified. Keep AC connected, repair the battery/power provider, and rerun"
            }
        }

        if ($chargePercent -lt $MinimumChargePercent) {
            return @{
                Status = "Blocked"; HasBattery = $true; OnBattery = $false; OnACPower = $true
                ChargePercent = $chargePercent
                RequiredChargePercent = $MinimumChargePercent
                Message = "Battery charge is $chargePercent%; charge to at least $MinimumChargePercent% while connected to AC before firmware installation"
            }
        }

        return @{
            Status = "Ready"; HasBattery = $true; OnBattery = $false; OnACPower = $true
            ChargePercent = $chargePercent
            RequiredChargePercent = $MinimumChargePercent
            Message = "AC power is online and battery charge is $chargePercent%"
        }
    } catch {
        return @{
            Status = "Unknown"; HasBattery = $null; OnBattery = $false; OnACPower = $false
            ChargePercent = $null
            RequiredChargePercent = $MinimumChargePercent
            Message = "$($_.Exception.Message). Connect AC, repair the power-status provider, and rerun"
        }
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
        if (-not (Get-Command Get-BitLockerVolume -ErrorAction SilentlyContinue)) {
            throw "Get-BitLockerVolume is unavailable"
        }

        $bl = Get-BitLockerVolume -MountPoint $env:SystemDrive -ErrorAction Stop
        if ($null -eq $bl) { throw "No BitLocker volume was returned for $env:SystemDrive" }

        $protectionStatus = [string]$bl.ProtectionStatus
        $volumeStatus = [string]$bl.VolumeStatus
        $encryptionMethod = [string]$bl.EncryptionMethod
        $canSuspend = [bool](Get-Command Suspend-BitLocker -ErrorAction SilentlyContinue) -and
            [bool](Get-Command Resume-BitLocker -ErrorAction SilentlyContinue)

        if ($protectionStatus -notin @("On", "Off")) {
            throw "Unexpected protection state '$protectionStatus'"
        }

        $encryptionActive = $encryptionMethod -notin @("", "None", "0") -and $volumeStatus -ne "FullyDecrypted"
        $protectionOn = $protectionStatus -eq "On"
        $suspended = $encryptionActive -and -not $protectionOn
        $state = if ($protectionOn) { "Protected" } elseif ($suspended) { "Suspended" } else { "Disabled" }

        return @{
            Status = "Ready"
            State = $state
            Enabled = $encryptionActive
            ProtectionOn = $protectionOn
            Suspended = $suspended
            CanSuspend = $canSuspend
            Method = $encryptionMethod
            Message = $(if ($protectionOn) {
                "BitLocker protection is active; firmware requires a verified provider auto-suspend path"
            } elseif ($suspended) {
                "BitLocker is already suspended"
            } else {
                "BitLocker is disabled on the operating-system volume"
            })
        }
    } catch {
        return @{
            Status = "Unknown"; State = "Unknown"; Enabled = $null; ProtectionOn = $null
            Suspended = $null; CanSuspend = $false; Method = "Unknown"
            Message = "BitLocker state could not be verified: $($_.Exception.Message). Run Get-BitLockerVolume for $env:SystemDrive and resolve the error"
        }
    }
}

function Test-OEMUpdateIsFirmware {
    param([AllowNull()][object]$Update)

    if ($null -eq $Update) { return $false }
    $identity = @(
        [string](Get-ResultValue -Result $Update -Names @("Title", "Name") -Default ""),
        [string](Get-ResultValue -Result $Update -Names @("Category") -Default ""),
        [string](Get-ResultValue -Result $Update -Names @("Type") -Default "")
    ) -join " "
    return $identity -match "(?i)(^|\W)(BIOS|UEFI|firmware)(\W|$)"
}

function Test-FirmwareReadiness {
    param(
        [switch]$Requested,
        [Parameter(Mandatory = $true)]
        [ValidateSet("Dell", "Lenovo", "HP")]
        [string]$Provider,
        [AllowNull()][object]$SystemInfo,
        [ValidateSet("Ready", "Blocked", "Unknown")]
        [string]$ToolState = "Unknown",
        [string]$ToolMessage = "",
        [bool]$SupportsBitLockerAutoSuspend = $false,
        [AllowNull()][object]$Prerequisites = $null
    )

    if (-not $Requested) {
        return [PSCustomObject][ordered]@{
            Requested = $false; Safe = $false; Status = "NotRequested"; Provider = $Provider
            Model = [string](Get-ResultValue -Result $SystemInfo -Names @("Model") -Default "")
            BitLockerProtectionOn = $false; RequiresBitLockerSuspension = $false
            Message = "Firmware and BIOS updates were not requested; use -IncludeBIOS after reviewing the safety prerequisites"
            Checks = @()
        }
    }

    $disk = Get-ResultValue -Result $Prerequisites -Names @("Disk") -Default $null
    $power = Get-ResultValue -Result $Prerequisites -Names @("Power") -Default $null
    $bitLocker = Get-ResultValue -Result $Prerequisites -Names @("BitLocker") -Default $null
    if ($null -eq $disk) { $disk = Test-DiskSpace -MinGB $MinDiskSpaceGB }
    if ($null -eq $power) { $power = Test-BatteryPower -MinimumChargePercent $MinFirmwareChargePercent }
    if ($null -eq $bitLocker) { $bitLocker = Test-BitLockerEnabled }

    $checks = [System.Collections.ArrayList]::new()
    $manufacturer = [string](Get-ResultValue -Result $SystemInfo -Names @("Manufacturer") -Default "")
    $model = [string](Get-ResultValue -Result $SystemInfo -Names @("Model") -Default "")
    $manufacturerPattern = if ($Provider -eq "Dell") {
        "DELL|ALIENWARE"
    } elseif ($Provider -eq "Lenovo") {
        "LENOVO"
    } else {
        "HP|HEWLETT"
    }

    if ([string]::IsNullOrWhiteSpace($manufacturer) -or [string]::IsNullOrWhiteSpace($model)) {
        [void]$checks.Add([PSCustomObject]@{
            Name = "Model"; Status = "Unknown"
            Message = "Manufacturer/model inventory is incomplete. Repair CIM/WMI (Win32_ComputerSystem) and rerun"
        })
    } elseif ($manufacturer -notmatch $manufacturerPattern) {
        [void]$checks.Add([PSCustomObject]@{
            Name = "Model"; Status = "Blocked"
            Message = "Model '$manufacturer $model' does not match the $Provider adapter. Use the matching OEM updater"
        })
    } else {
        [void]$checks.Add([PSCustomObject]@{
            Name = "Model"; Status = "Ready"; Message = "Detected $manufacturer $model"
        })
    }

    $effectiveToolMessage = if ([string]::IsNullOrWhiteSpace($ToolMessage)) {
        if ($ToolState -eq "Ready") {
            "$Provider tooling completed an applicability scan for this model"
        } else {
            "$Provider tooling did not prove model support. Repair or install the OEM tool and rerun its applicability scan"
        }
    } else {
        $ToolMessage
    }
    [void]$checks.Add([PSCustomObject]@{ Name = "Tool"; Status = $ToolState; Message = $effectiveToolMessage })

    $diskStatus = [string](Get-ResultValue -Result $disk -Names @("Status") -Default "Unknown")
    $diskMessage = [string](Get-ResultValue -Result $disk -Names @("Message") -Default "Disk readiness is unknown; repair CIM/WMI and rerun")
    [void]$checks.Add([PSCustomObject]@{ Name = "Disk"; Status = $diskStatus; Message = $diskMessage })

    $powerStatus = [string](Get-ResultValue -Result $power -Names @("Status") -Default "Unknown")
    $powerMessage = [string](Get-ResultValue -Result $power -Names @("Message") -Default "Power readiness is unknown; connect AC and rerun")
    [void]$checks.Add([PSCustomObject]@{ Name = "Power"; Status = $powerStatus; Message = $powerMessage })

    $bitLockerStatus = [string](Get-ResultValue -Result $bitLocker -Names @("Status") -Default "Unknown")
    $bitLockerProtectionOn = [bool](Get-ResultValue -Result $bitLocker -Names @("ProtectionOn") -Default $false)
    if ($bitLockerStatus -ne "Ready") {
        [void]$checks.Add([PSCustomObject]@{
            Name = "BitLocker"; Status = "Unknown"
            Message = [string](Get-ResultValue -Result $bitLocker -Names @("Message") -Default "BitLocker state is unknown; verify it and rerun")
        })
    } elseif ($bitLockerProtectionOn -and -not $SupportsBitLockerAutoSuspend) {
        [void]$checks.Add([PSCustomObject]@{
            Name = "BitLocker"; Status = "Blocked"
            Message = "$Provider has no verified automatic BitLocker suspension path. Suspend protection for one reboot, verify recovery-key escrow, and rerun"
        })
    } elseif ($bitLockerProtectionOn) {
        [void]$checks.Add([PSCustomObject]@{
            Name = "BitLocker"; Status = "Ready"
            Message = "$Provider supports automatic one-update BitLocker suspension"
        })
    } else {
        [void]$checks.Add([PSCustomObject]@{
            Name = "BitLocker"; Status = "Ready"
            Message = [string](Get-ResultValue -Result $bitLocker -Names @("Message") -Default "BitLocker does not block firmware")
        })
    }

    $unknownChecks = @($checks | Where-Object { $_.Status -eq "Unknown" })
    $blockedChecks = @($checks | Where-Object { $_.Status -eq "Blocked" })
    $status = if ($unknownChecks.Count -gt 0) {
        "Unknown"
    } elseif ($blockedChecks.Count -gt 0) {
        "Blocked"
    } else {
        "Ready"
    }
    $reasons = @($checks | Where-Object { $_.Status -ne "Ready" } | ForEach-Object { "$($_.Name): $($_.Message)" })
    $message = if ($status -eq "Ready") {
        "$Provider firmware prerequisites are known-ready for $model"
    } else {
        "$Provider firmware safety is $($status.ToLowerInvariant()): $($reasons -join '; '). -Force cannot override firmware safety"
    }

    return [PSCustomObject][ordered]@{
        Requested = $true
        Safe = ($status -eq "Ready")
        Status = $status
        Provider = $Provider
        Model = $model
        BitLockerProtectionOn = $bitLockerProtectionOn
        RequiresBitLockerSuspension = ($status -eq "Ready" -and $bitLockerProtectionOn)
        Message = $message
        Checks = @($checks)
    }
}

function Get-FirmwareUpdatePolicy {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Readiness
    )

    $includeFirmware = [bool](Get-ResultValue -Result $Readiness -Names @("Requested") -Default $false) -and
        [string](Get-ResultValue -Result $Readiness -Names @("Status") -Default "Unknown") -eq "Ready"
    return [PSCustomObject][ordered]@{
        IncludeFirmware = $includeFirmware
        DellUpdateTypes = $(if ($includeFirmware) {
            @("bios", "firmware", "driver", "application", "others")
        } else {
            @("driver", "application", "others")
        })
        HPCategories = $(if ($includeFirmware) { @("Drivers", "Firmware", "BIOS") } else { @("Drivers") })
        LenovoExcludeFirmware = (-not $includeFirmware)
        DellAutoSuspendBitLocker = ($includeFirmware -and
            [string](Get-ResultValue -Result $Readiness -Names @("Provider") -Default "") -eq "Dell" -and
            [bool](Get-ResultValue -Result $Readiness -Names @("RequiresBitLockerSuspension") -Default $false))
    }
}

function New-FirmwareReadinessItem {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Readiness
    )

    $readinessStatus = [string](Get-ResultValue -Result $Readiness -Names @("Status") -Default "Unknown")
    $itemStatus = if ($readinessStatus -eq "Ready") {
        "Succeeded"
    } elseif ($readinessStatus -eq "NotRequested") {
        "Skipped"
    } elseif ($readinessStatus -eq "Blocked") {
        "Blocked"
    } else {
        "Unknown"
    }
    return New-UpdateItemResult -Name "Firmware safety" -Status $itemStatus `
        -Message ([string](Get-ResultValue -Result $Readiness -Names @("Message") -Default "Firmware readiness was not reported"))
}

function Test-MeteredConnection {
    return ((Get-NetworkCostState).Status -eq "Metered")
}

function Get-NetworkCostState {
    try {
        $cost = [Windows.Networking.Connectivity.NetworkInformation, Windows, ContentType=WindowsRuntime]::GetInternetConnectionProfile().GetConnectionCost()
        $costType = [string]$cost.NetworkCostType
        if ($costType -eq "Unrestricted") {
            return [PSCustomObject]@{
                Status = "Unrestricted"; CostType = $costType; Known = $true
                Reason = "Connection is unrestricted"
            }
        }
        return [PSCustomObject]@{
            Status = "Metered"; CostType = $costType; Known = $true
            Reason = "Connection cost is $costType"
        }
    } catch {
        return [PSCustomObject]@{
            Status = "Unknown"; CostType = ""; Known = $false
            Reason = "Network cost could not be determined: $($_.Exception.Message)"
        }
    }
}

function Get-DownloadPolicy {
    param(
        [Parameter(Mandatory = $true)]
        [object]$NetworkCost,
        [bool]$AllowOverride = $false,
        [bool]$DryRunMode = $false
    )

    $status = "Allowed"
    $reason = "Network cost policy permits provider downloads"
    $override = $false
    if ([string]$NetworkCost.Status -eq "Metered") {
        if ($AllowOverride) {
            $override = $true
            $reason = "Metered-network download policy explicitly overridden by the operator"
        } else {
            $status = "Blocked"
            $reason = "Known metered connection blocks provider downloads; use -AllowMeteredNetwork for an audited override"
        }
    } elseif ([string]$NetworkCost.Status -eq "Unknown") {
        $status = "AllowedWithWarning"
        $reason = "Network cost is unknown; provider downloads remain allowed and the uncertainty is recorded"
    }
    return [PSCustomObject][ordered]@{
        SchemaVersion = 1
        Status = $status
        Allowed = ($status -ne "Blocked")
        Deferred = ($status -eq "Blocked")
        DryRun = $DryRunMode
        AuditedOverride = $override
        NetworkCost = $NetworkCost
        Reason = $reason
        EvaluatedAt = (Get-Date).ToUniversalTime().ToString("o")
    }
}

function Test-DownloadAllowed {
    if ($null -eq $script:DownloadPolicy) { return $true }
    return [bool]$script:DownloadPolicy.Allowed
}

function Get-SystemInfo {
    $cs = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction SilentlyContinue
    $bios = Get-CimInstance -ClassName Win32_BIOS -ErrorAction SilentlyContinue
    $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction SilentlyContinue
    $cpu = Get-CimInstance -ClassName Win32_Processor -ErrorAction SilentlyContinue | Select-Object -First 1
    $currentVersion = Get-ItemProperty -LiteralPath "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" `
        -ErrorAction SilentlyContinue
    $productType = if ($null -ne $os.ProductType) { [int]$os.ProductType } else { 0 }
    $installationType = [string]$currentVersion.InstallationType
    if ([string]::IsNullOrWhiteSpace($installationType) -and $productType -eq 1) {
        $installationType = "Client"
    }
    $runContext = try {
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        if ($identity.IsSystem) { "System" } else { "AdministratorUser" }
    } catch {
        "Unknown"
    }

    return @{
        Manufacturer = $cs.Manufacturer
        Model = $cs.Model
        SerialNumber = $bios.SerialNumber
        BIOSVersion = $bios.SMBIOSBIOSVersion
        BIOSDate = $bios.ReleaseDate
        OSName = $os.Caption
        OSVersion = $os.Version
        OSBuild = $os.BuildNumber
        ProductType = $productType
        OperatingSystemSKU = $os.OperatingSystemSKU
        InstallationType = $installationType
        EditionID = $currentVersion.EditionID
        DisplayVersion = $currentVersion.DisplayVersion
        Architecture = Get-SystemArchitecture
        PowerShellVersion = $PSVersionTable.PSVersion.ToString()
        PowerShellEdition = [string]$PSVersionTable.PSEdition
        ExecutionContext = $runContext
        IsServer = ($productType -in @(2, 3))
        IsServerCore = ($installationType -match "(?i)Core")
        TotalRAM = [math]::Round($cs.TotalPhysicalMemory / 1GB, 1)
        Processor = $cpu.Name
    }
}

# ============================================================================
# PLATFORM AND PROVIDER CAPABILITY CONTRACT
# ============================================================================

function Get-CapabilityMatrix {
    $manifest = Get-AcquisitionManifest
    return [ordered]@{
        SchemaVersion = $script:CapabilitySchemaVersion
        Platform = [ordered]@{
            MinimumBuild = 14393
            MinimumPowerShellVersion = "5.1"
            Architectures = @("x86", "x64", "arm64")
            ProductTypes = @(1, 2, 3)
            InstallationTypes = @("Client", "Server", "Server Core")
            PowerShellEditions = @("Desktop", "Core")
            UnsupportedEditionPatterns = @("(?i)Nano")
        }
        Providers = [ordered]@{
            WindowsUpdate = [ordered]@{
                DisplayName = "Windows Update"
                ClientMinimumBuild = 14393
                ServerMinimumBuild = 14393
                SupportsServer = $true
                SupportsServerCore = $true
                Architectures = @("x86", "x64", "arm64")
                Contexts = @("AdministratorUser", "System")
                BootstrapContexts = @("AdministratorUser", "System")
                DependencyName = "PSWindowsUpdate"
                VersionPolicy = "BuiltInFallback"
                MinimumVersion = [string]$manifest.PSWindowsUpdate.MinimumVersion
                AcquisitionVersion = [string]$manifest.PSWindowsUpdate.ExactVersion
                ManufacturerPattern = ""
            }
            WindowsServicing = [ordered]@{
                DisplayName = "Windows servicing"
                ClientMinimumBuild = 14393
                ServerMinimumBuild = 14393
                SupportsServer = $true
                SupportsServerCore = $true
                Architectures = @("x86", "x64", "arm64")
                Contexts = @("AdministratorUser", "System")
                BootstrapContexts = @()
                DependencyName = "DISM/SFC"
                VersionPolicy = "Inbox"
                MinimumVersion = "Inbox"
                AcquisitionVersion = ""
                ManufacturerPattern = ""
            }
            Winget = [ordered]@{
                DisplayName = "WinGet CLI"
                ClientMinimumBuild = 17763
                ServerMinimumBuild = 26100
                SupportsServer = $true
                SupportsServerCore = $false
                Architectures = @("x86", "x64", "arm64")
                Contexts = @("AdministratorUser")
                BootstrapContexts = @("AdministratorUser")
                DependencyName = "WinGet"
                VersionPolicy = "VerifiedAcquisition"
                MinimumVersion = [string]$manifest.WinGet.MinimumVersion
                AcquisitionVersion = [string]$manifest.WinGet.ExactVersion
                ManufacturerPattern = ""
            }
            Dell = [ordered]@{
                DisplayName = "Dell Command Update"
                ClientMinimumBuild = 14393
                ServerMinimumBuild = 0
                SupportsServer = $false
                SupportsServerCore = $false
                Architectures = @($manifest.DellCommandUpdate.Architectures)
                Contexts = @("AdministratorUser", "System")
                BootstrapContexts = @("AdministratorUser")
                DependencyName = "DellCommandUpdate"
                VersionPolicy = "VerifiedAcquisition"
                MinimumVersion = [string]$manifest.DellCommandUpdate.MinimumVersion
                AcquisitionVersion = [string]$manifest.DellCommandUpdate.ExactVersion
                AdditionalMinimumVersions = [ordered]@{
                    InventoryCollector = [string]$manifest.DellCommandUpdate.InventoryCollectorMinimum
                }
                ManufacturerPattern = "(?i)DELL|ALIENWARE"
            }
            Lenovo = [ordered]@{
                DisplayName = "Lenovo LSUClient"
                ClientMinimumBuild = 14393
                ServerMinimumBuild = 0
                SupportsServer = $false
                SupportsServerCore = $false
                Architectures = @("x64")
                Contexts = @("AdministratorUser", "System")
                BootstrapContexts = @("AdministratorUser", "System")
                DependencyName = "LSUClient"
                VersionPolicy = "VerifiedAcquisition"
                MinimumVersion = [string]$manifest.LSUClient.MinimumVersion
                AcquisitionVersion = [string]$manifest.LSUClient.ExactVersion
                ManufacturerPattern = "(?i)LENOVO"
            }
            HP = [ordered]@{
                DisplayName = "HP Image Assistant"
                ClientMinimumBuild = 17763
                ServerMinimumBuild = 0
                SupportsServer = $false
                SupportsServerCore = $false
                Architectures = @($manifest.HPIA.Architectures)
                Contexts = @("AdministratorUser", "System")
                BootstrapContexts = @("AdministratorUser", "System")
                DependencyName = "HPIA"
                VersionPolicy = "VerifiedAcquisition"
                MinimumVersion = [string]$manifest.HPIA.InstalledMinimumVersion
                AcquisitionVersion = [string]$manifest.HPIA.ExactVersion
                ManufacturerPattern = "(?i)HP|HEWLETT"
            }
        }
    }
}

function Get-ProviderVersionInventory {
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.IDictionary]$SystemInfo,
        [AllowNull()][System.Collections.IDictionary]$Matrix = $null
    )

    if ($null -eq $Matrix) { $Matrix = Get-CapabilityMatrix }
    $inventory = [ordered]@{}

    $psWindowsUpdate = Test-InstalledPSModule -ModuleName "PSWindowsUpdate"
    $inventory.WindowsUpdate = if ($psWindowsUpdate.Valid) {
        [ordered]@{
            Status = "Ready"; Version = [string]$psWindowsUpdate.Version
            Source = "PSWindowsUpdate"; Reason = "Verified pinned PSWindowsUpdate payload is installed"
        }
    } else {
        [ordered]@{
            Status = "BuiltInFallback"; Version = [string]$SystemInfo.OSVersion
            Source = "Windows Update Agent"
            Reason = "Verified PSWindowsUpdate is not installed; the inbox WUA fallback remains available"
        }
    }

    $windowsRoot = if ([string]::IsNullOrWhiteSpace([string]$script:WindowsRoot)) {
        [string]$env:SystemRoot
    } else {
        [string]$script:WindowsRoot
    }
    $dismPath = Join-Path $windowsRoot "System32\dism.exe"
    $sfcPath = Join-Path $windowsRoot "System32\sfc.exe"
    $servicingReady = (
        (Test-Path -LiteralPath $dismPath -PathType Leaf) -and
        (Test-Path -LiteralPath $sfcPath -PathType Leaf)
    )
    $inventory.WindowsServicing = [ordered]@{
        Status = $(if ($servicingReady) { "Ready" } else { "Unavailable" })
        Version = [string]$SystemInfo.OSVersion
        Source = "Windows inbox"
        Reason = $(if ($servicingReady) {
            "Inbox DISM and SFC are present"
        } else {
            "Inbox DISM or SFC is missing"
        })
    }

    $wingetEvidence = Get-WingetTrustEvidence
    $inventory.Winget = if ($wingetEvidence.Valid) {
        [ordered]@{
            Status = "Ready"; Version = [string]$wingetEvidence.Version
            Source = "Microsoft Desktop App Installer"
            Reason = "Installed WinGet meets the pinned version and publisher policy"
        }
    } else {
        [ordered]@{
            Status = "AcquisitionRequired"; Version = ""; Source = "Pinned acquisition manifest"
            Reason = [string]$wingetEvidence.Reason
        }
    }

    $manufacturer = [string]$SystemInfo.Manufacturer
    $dellSpec = Get-AcquisitionManifestEntry -Name "DellCommandUpdate"
    if ($manufacturer -match [string]$Matrix.Providers.Dell.ManufacturerPattern) {
        $dellEvidence = $null
        foreach ($candidate in @(
            "${env:ProgramFiles}\Dell\CommandUpdate\dcu-cli.exe",
            "${env:ProgramFiles(x86)}\Dell\CommandUpdate\dcu-cli.exe"
        )) {
            $candidateEvidence = Get-InstalledExecutableEvidence -Path $candidate `
                -MinimumVersion $dellSpec.MinimumVersion -PublisherPattern $dellSpec.PublisherPattern
            if ($candidateEvidence.Valid) {
                $dellEvidence = $candidateEvidence
                break
            }
        }
        $inventoryEvidence = Get-DellInventoryCollectorEvidence
        $inventory.Dell = if ($null -ne $dellEvidence -and $inventoryEvidence.Valid) {
            [ordered]@{
                Status = "Ready"; Version = [string]$dellEvidence.Version
                Source = "Dell Command Update"
                Reason = "Dell Command Update and Inventory Collector meet the minimum version policy"
            }
        } else {
            $dellReasons = @()
            if ($null -eq $dellEvidence) { $dellReasons += "Dell Command Update is missing or unverified" }
            if (-not $inventoryEvidence.Valid) {
                $dellReasons += "Inventory Collector: $($inventoryEvidence.Reason)"
            }
            [ordered]@{
                Status = "AcquisitionRequired"; Version = ""; Source = "Pinned acquisition manifest"
                Reason = $dellReasons -join "; "
            }
        }
    } else {
        $inventory.Dell = [ordered]@{
            Status = "NotApplicable"; Version = ""; Source = ""
            Reason = "Manufacturer does not use the Dell adapter"
        }
    }

    if ($manufacturer -match [string]$Matrix.Providers.Lenovo.ManufacturerPattern) {
        $lenovoEvidence = Test-InstalledPSModule -ModuleName "LSUClient"
        $inventory.Lenovo = if ($lenovoEvidence.Valid) {
            [ordered]@{
                Status = "Ready"; Version = [string]$lenovoEvidence.Version
                Source = "LSUClient"; Reason = "Verified pinned LSUClient payload is installed"
            }
        } else {
            [ordered]@{
                Status = "AcquisitionRequired"; Version = ""; Source = "Pinned acquisition manifest"
                Reason = [string]$lenovoEvidence.Reason
            }
        }
    } else {
        $inventory.Lenovo = [ordered]@{
            Status = "NotApplicable"; Version = ""; Source = ""
            Reason = "Manufacturer does not use the Lenovo adapter"
        }
    }

    if ($manufacturer -match [string]$Matrix.Providers.HP.ManufacturerPattern) {
        $hpSpec = Get-AcquisitionManifestEntry -Name "HPIA"
        $hpEvidence = $null
        foreach ($searchPath in @(
            $script:HPIAInstallRoot,
            "C:\SWSetup\SP*",
            "${env:ProgramFiles}\HP\HPIA"
        )) {
            foreach ($candidate in @(Get-ChildItem -Path $searchPath -Filter "HPImageAssistant.exe" `
                -File -Recurse -ErrorAction SilentlyContinue)) {
                $candidateEvidence = Get-InstalledExecutableEvidence -Path $candidate.FullName `
                    -MinimumVersion $hpSpec.InstalledMinimumVersion `
                    -PublisherPattern $hpSpec.InstalledPublisherPattern
                if ($candidateEvidence.Valid) {
                    $hpEvidence = $candidateEvidence
                    break
                }
            }
            if ($null -ne $hpEvidence) { break }
        }
        $inventory.HP = if ($null -ne $hpEvidence) {
            [ordered]@{
                Status = "Ready"; Version = [string]$hpEvidence.Version
                Source = "HP Image Assistant"
                Reason = "Installed HP Image Assistant meets the minimum version and publisher policy"
            }
        } else {
            [ordered]@{
                Status = "AcquisitionRequired"; Version = ""; Source = "Pinned acquisition manifest"
                Reason = "HP Image Assistant is missing or does not meet the verified version policy"
            }
        }
    } else {
        $inventory.HP = [ordered]@{
            Status = "NotApplicable"; Version = ""; Source = ""
            Reason = "Manufacturer does not use the HP adapter"
        }
    }

    return $inventory
}

function Get-PlatformCapability {
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.IDictionary]$SystemInfo,
        [AllowNull()][System.Collections.IDictionary]$Matrix = $null
    )

    if ($null -eq $Matrix) { $Matrix = Get-CapabilityMatrix }
    $unknown = [System.Collections.ArrayList]::new()
    $unsupported = [System.Collections.ArrayList]::new()

    $build = 0
    if (-not [int]::TryParse([string]$SystemInfo.OSBuild, [ref]$build)) {
        [void]$unknown.Add("OS build could not be determined")
    } elseif ($build -lt [int]$Matrix.Platform.MinimumBuild) {
        [void]$unsupported.Add(
            "OS build $build is below the supported floor $($Matrix.Platform.MinimumBuild)"
        )
    }

    $productType = 0
    if (-not [int]::TryParse([string]$SystemInfo.ProductType, [ref]$productType) -or
        $productType -notin @($Matrix.Platform.ProductTypes)) {
        [void]$unknown.Add("Windows product type could not be classified")
    }

    $installationType = [string]$SystemInfo.InstallationType
    if ([string]::IsNullOrWhiteSpace($installationType)) {
        [void]$unknown.Add("Windows installation type could not be determined")
    } elseif ($installationType -notin @($Matrix.Platform.InstallationTypes)) {
        [void]$unsupported.Add("Windows installation type '$installationType' is unsupported")
    }

    $editionId = [string]$SystemInfo.EditionID
    $operatingSystemSKU = 0
    [void][int]::TryParse([string]$SystemInfo.OperatingSystemSKU, [ref]$operatingSystemSKU)
    if ([string]::IsNullOrWhiteSpace($editionId) -and $operatingSystemSKU -le 0) {
        [void]$unknown.Add("Windows edition could not be determined")
    } elseif (@($Matrix.Platform.UnsupportedEditionPatterns | Where-Object {
        -not [string]::IsNullOrWhiteSpace($editionId) -and $editionId -match [string]$_
    }).Count -gt 0) {
        [void]$unsupported.Add("Windows edition '$editionId' is unsupported")
    }

    $architecture = ([string]$SystemInfo.Architecture).ToLowerInvariant()
    if ([string]::IsNullOrWhiteSpace($architecture) -or $architecture -eq "unknown") {
        [void]$unknown.Add("Processor architecture could not be determined")
    } elseif ($architecture -notin @($Matrix.Platform.Architectures)) {
        [void]$unsupported.Add("Architecture '$architecture' is unsupported")
    }

    $powerShellVersion = ConvertTo-SafeVersion -Value ([string]$SystemInfo.PowerShellVersion)
    if ($null -eq $powerShellVersion) {
        [void]$unknown.Add("PowerShell version could not be determined")
    } elseif (-not (Test-VersionAtLeast -Version $powerShellVersion `
        -MinimumVersion ([string]$Matrix.Platform.MinimumPowerShellVersion))) {
        [void]$unsupported.Add(
            "PowerShell $powerShellVersion is below $($Matrix.Platform.MinimumPowerShellVersion)"
        )
    }

    $powerShellEdition = [string]$SystemInfo.PowerShellEdition
    if ($powerShellEdition -notin @($Matrix.Platform.PowerShellEditions)) {
        [void]$unsupported.Add("PowerShell edition '$powerShellEdition' is unsupported")
    }

    $runContext = [string]$SystemInfo.ExecutionContext
    if ($runContext -notin @("AdministratorUser", "System")) {
        [void]$unknown.Add("Execution context could not be classified")
    }

    $status = if ($unknown.Count -gt 0) {
        "Unknown"
    } elseif ($unsupported.Count -gt 0) {
        "Unsupported"
    } else {
        "Ready"
    }
    $reasons = @($unknown) + @($unsupported)
    return [ordered]@{
        Status = $status
        Supported = ($status -eq "Ready")
        Reason = $(if ($reasons.Count) { $reasons -join "; " } else { "Platform capability checks passed" })
        OSBuild = $build
        ProductType = $productType
        OperatingSystemSKU = $operatingSystemSKU
        InstallationType = $installationType
        EditionID = $editionId
        IsServer = ($productType -in @(2, 3))
        IsServerCore = ($installationType -match "(?i)Core")
        Architecture = $architecture
        PowerShellVersion = [string]$powerShellVersion
        PowerShellEdition = $powerShellEdition
        ExecutionContext = $runContext
    }
}

function Get-ProviderCapability {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("WindowsUpdate", "WindowsServicing", "Winget", "Dell", "Lenovo", "HP")]
        [string]$Provider,
        [Parameter(Mandatory = $true)]
        [System.Collections.IDictionary]$SystemInfo,
        [AllowNull()][System.Collections.IDictionary]$Platform = $null,
        [AllowNull()][System.Collections.IDictionary]$Matrix = $null,
        [AllowNull()][System.Collections.IDictionary]$VersionState = $null
    )

    if ($null -eq $Matrix) { $Matrix = Get-CapabilityMatrix }
    if ($null -eq $Platform) {
        $Platform = Get-PlatformCapability -SystemInfo $SystemInfo -Matrix $Matrix
    }
    $spec = $Matrix.Providers[$Provider]
    $reasons = [System.Collections.ArrayList]::new()
    if ($null -eq $VersionState) {
        $VersionState = [ordered]@{
            Status = "Unknown"; Version = ""; Source = ""
            Reason = "Provider version was not assessed"
        }
    }

    if (-not [bool]$Platform.Supported) {
        [void]$reasons.Add("Platform is $($Platform.Status): $($Platform.Reason)")
    } else {
        $buildFloor = if ([bool]$Platform.IsServer) {
            [int]$spec.ServerMinimumBuild
        } else {
            [int]$spec.ClientMinimumBuild
        }
        if ([bool]$Platform.IsServer -and -not [bool]$spec.SupportsServer) {
            [void]$reasons.Add("$($spec.DisplayName) is not supported on Windows Server")
        } elseif ([int]$Platform.OSBuild -lt $buildFloor) {
            $platformLabel = if ([bool]$Platform.IsServer) { "Windows Server" } else { "Windows client" }
            [void]$reasons.Add(
                "$($spec.DisplayName) requires $platformLabel build $buildFloor or later"
            )
        }
        if ([bool]$Platform.IsServerCore -and -not [bool]$spec.SupportsServerCore) {
            [void]$reasons.Add("$($spec.DisplayName) is not supported on Server Core")
        }
        if ([string]$Platform.Architecture -notin @($spec.Architectures)) {
            [void]$reasons.Add(
                "$($spec.DisplayName) is not approved for architecture '$($Platform.Architecture)'"
            )
        }
        if ([string]$Platform.ExecutionContext -notin @($spec.Contexts)) {
            [void]$reasons.Add(
                "$($spec.DisplayName) is not supported in the $($Platform.ExecutionContext) context"
            )
        }
        if (-not [string]::IsNullOrWhiteSpace([string]$spec.ManufacturerPattern) -and
            ([string]$SystemInfo.Manufacturer) -notmatch [string]$spec.ManufacturerPattern) {
            [void]$reasons.Add(
                "$($spec.DisplayName) does not match manufacturer '$($SystemInfo.Manufacturer)'"
            )
        }
        if ($reasons.Count -eq 0) {
            switch ([string]$VersionState.Status) {
                { $_ -in @("Ready", "BuiltInFallback") } {}
                "AcquisitionRequired" {
                    if ([string]$Platform.ExecutionContext -notin @($spec.BootstrapContexts)) {
                        [void]$reasons.Add(
                            "$($spec.DisplayName) is not installed at the approved version and cannot be bootstrapped in the $($Platform.ExecutionContext) context"
                        )
                    }
                }
                default {
                    [void]$reasons.Add(
                        "$($spec.DisplayName) version gate is $($VersionState.Status): $($VersionState.Reason)"
                    )
                }
            }
        }
    }

    $status = if ($reasons.Count) {
        "Unsupported"
    } elseif ([string]$VersionState.Status -eq "AcquisitionRequired") {
        "RequiresAcquisition"
    } else {
        "Ready"
    }
    return [ordered]@{
        Provider = $Provider
        DisplayName = [string]$spec.DisplayName
        Status = $status
        Supported = ($reasons.Count -eq 0)
        Reason = $(if ($reasons.Count) {
            @($reasons) -join "; "
        } elseif ($status -eq "RequiresAcquisition") {
            "$($spec.DisplayName) requires verified acquisition of version $($spec.AcquisitionVersion): $($VersionState.Reason)"
        } else {
            "$($spec.DisplayName) capability and version checks passed"
        })
        DependencyName = [string]$spec.DependencyName
        VersionPolicy = [string]$spec.VersionPolicy
        MinimumVersion = [string]$spec.MinimumVersion
        AcquisitionVersion = [string]$spec.AcquisitionVersion
        VersionStatus = [string]$VersionState.Status
        DetectedVersion = [string]$VersionState.Version
        VersionSource = [string]$VersionState.Source
        VersionReason = [string]$VersionState.Reason
        Architectures = @($spec.Architectures)
        Contexts = @($spec.Contexts)
        BootstrapContexts = @($spec.BootstrapContexts)
    }
}

function Get-CapabilityAssessment {
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.IDictionary]$SystemInfo,
        [AllowNull()][System.Collections.IDictionary]$VersionInventory = $null
    )

    $matrix = Get-CapabilityMatrix
    $platform = Get-PlatformCapability -SystemInfo $SystemInfo -Matrix $matrix
    if ($null -eq $VersionInventory) {
        $VersionInventory = Get-ProviderVersionInventory -SystemInfo $SystemInfo -Matrix $matrix
    }
    $providers = [ordered]@{}
    foreach ($providerName in @("WindowsUpdate", "WindowsServicing", "Winget", "Dell", "Lenovo", "HP")) {
        $providers[$providerName] = Get-ProviderCapability -Provider $providerName `
            -SystemInfo $SystemInfo -Platform $platform -Matrix $matrix `
            -VersionState $VersionInventory[$providerName]
    }
    return [ordered]@{
        SchemaVersion = $script:CapabilitySchemaVersion
        EvaluatedAt = (Get-Date).ToUniversalTime().ToString("o")
        Platform = $platform
        Providers = $providers
    }
}

function Get-AssessedProviderCapability {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("WindowsUpdate", "WindowsServicing", "Winget", "Dell", "Lenovo", "HP")]
        [string]$Provider
    )

    if ($null -eq $script:CapabilityAssessment) {
        throw "Platform capability assessment has not been initialized"
    }
    return $script:CapabilityAssessment.Providers[$Provider]
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

function Get-ServiceStartupType {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ServiceName
    )

    $service = Get-Service -Name $ServiceName -ErrorAction Stop
    $startupType = [string]$service.StartType
    if ($startupType -eq "Automatic") {
        $servicePath = "HKLM:\SYSTEM\CurrentControlSet\Services\$ServiceName"
        $delayed = (Get-ItemProperty -LiteralPath $servicePath -Name "DelayedAutoStart" `
            -ErrorAction SilentlyContinue).DelayedAutoStart
        if ([int]$delayed -eq 1) { return "AutomaticDelayedStart" }
    }
    return $startupType
}

function Get-ServiceDelayedAutoStartSnapshot {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ServiceName
    )

    $servicePath = "HKLM:\SYSTEM\CurrentControlSet\Services\$ServiceName"
    $key = Get-Item -LiteralPath $servicePath -ErrorAction Stop
    $exists = @($key.GetValueNames()) -contains "DelayedAutoStart"
    return [ordered]@{
        Exists = $exists
        Value  = $(if ($exists) { [int]$key.GetValue("DelayedAutoStart") } else { 0 })
    }
}

function Set-ServiceStartupType {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseShouldProcessForStateChangingFunctions", "", Justification = "Called only after the service before-image is durable.")]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ServiceName,
        [Parameter(Mandatory = $true)]
        [ValidateSet("Automatic", "AutomaticDelayedStart", "Manual", "Disabled")]
        [string]$StartupType
    )

    $setServiceType = if ($StartupType -eq "AutomaticDelayedStart") { "Automatic" } else { $StartupType }
    Set-Service -Name $ServiceName -StartupType $setServiceType -ErrorAction Stop

    $servicePath = "HKLM:\SYSTEM\CurrentControlSet\Services\$ServiceName"
    if ($StartupType -eq "AutomaticDelayedStart") {
        New-ItemProperty -LiteralPath $servicePath -Name "DelayedAutoStart" -Value 1 `
            -PropertyType DWord -Force -ErrorAction Stop | Out-Null
    } elseif ($StartupType -eq "Automatic") {
        New-ItemProperty -LiteralPath $servicePath -Name "DelayedAutoStart" -Value 0 `
            -PropertyType DWord -Force -ErrorAction Stop | Out-Null
    }
}

function Get-ServiceSnapshot {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ServiceName
    )

    try {
        $service = Get-Service -Name $ServiceName -ErrorAction Stop
        $stableStatus = switch ([string]$service.Status) {
            "StartPending"    { [System.ServiceProcess.ServiceControllerStatus]::Running }
            "ContinuePending" { [System.ServiceProcess.ServiceControllerStatus]::Running }
            "StopPending"     { [System.ServiceProcess.ServiceControllerStatus]::Stopped }
            "PausePending"    { [System.ServiceProcess.ServiceControllerStatus]::Paused }
            default           { $null }
        }
        if ($null -ne $stableStatus) {
            $service.WaitForStatus($stableStatus, [TimeSpan]::FromSeconds(30))
            $service = Get-Service -Name $ServiceName -ErrorAction Stop
        }
        $delayedSnapshot = Get-ServiceDelayedAutoStartSnapshot -ServiceName $ServiceName
        return [ordered]@{
            Exists           = $true
            Name             = $ServiceName
            Status           = [string]$service.Status
            StartupType      = Get-ServiceStartupType -ServiceName $ServiceName
            DelayedAutoStartExists = [bool]$delayedSnapshot.Exists
            DelayedAutoStartValue  = [int]$delayedSnapshot.Value
            RestartOnRestore = $false
        }
    } catch {
        return [ordered]@{
            Exists = $false; Name = $ServiceName; Status = ""; StartupType = ""
            RestartOnRestore = $false
        }
    }
}

function Restore-ServiceSnapshot {
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.IDictionary]$Snapshot
    )

    if (-not [bool]$Snapshot.Exists) {
        return ($null -eq (Get-Service -Name ([string]$Snapshot.Name) -ErrorAction SilentlyContinue))
    }

    try {
        $serviceName = [string]$Snapshot.Name
        Set-ServiceStartupType -ServiceName $serviceName -StartupType ([string]$Snapshot.StartupType)
        $servicePath = "HKLM:\SYSTEM\CurrentControlSet\Services\$serviceName"
        if ([bool]$Snapshot.DelayedAutoStartExists) {
            New-ItemProperty -LiteralPath $servicePath -Name "DelayedAutoStart" `
                -Value ([int]$Snapshot.DelayedAutoStartValue) -PropertyType DWord -Force `
                -ErrorAction Stop | Out-Null
        } else {
            Remove-ItemProperty -LiteralPath $servicePath -Name "DelayedAutoStart" `
                -Force -ErrorAction SilentlyContinue
        }
        $current = Get-Service -Name $serviceName -ErrorAction Stop
        switch ([string]$Snapshot.Status) {
            "Running" {
                if ([bool]$Snapshot.RestartOnRestore -and $current.Status -eq "Running") {
                    Restart-Service -Name $serviceName -Force -ErrorAction Stop
                } elseif ($current.Status -ne "Running") {
                    Start-Service -Name $serviceName -ErrorAction Stop
                }
            }
            "Stopped" {
                if ($current.Status -ne "Stopped") {
                    Stop-Service -Name $serviceName -Force -ErrorAction Stop
                }
            }
            "Paused" {
                if ($current.Status -eq "Stopped") {
                    Start-Service -Name $serviceName -ErrorAction Stop
                }
                Suspend-Service -Name $serviceName -ErrorAction Stop
            }
            default {
                throw "Unsupported saved service status '$($Snapshot.Status)'"
            }
        }

        $deadline = (Get-Date).AddSeconds(30)
        do {
            $current = Get-Service -Name $serviceName -ErrorAction Stop
            $delayedCurrent = Get-ServiceDelayedAutoStartSnapshot -ServiceName $serviceName
            if ([string]$current.Status -eq [string]$Snapshot.Status -and
                (Get-ServiceStartupType -ServiceName $serviceName) -eq [string]$Snapshot.StartupType -and
                [bool]$delayedCurrent.Exists -eq [bool]$Snapshot.DelayedAutoStartExists -and
                (-not [bool]$Snapshot.DelayedAutoStartExists -or
                    [int]$delayedCurrent.Value -eq [int]$Snapshot.DelayedAutoStartValue)) {
                return $true
            }
            Start-Sleep -Milliseconds 250
        } while ((Get-Date) -lt $deadline)
        return $false
    } catch {
        return $false
    }
}

function Set-JournaledServiceState {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseShouldProcessForStateChangingFunctions", "", Justification = "The run-scoped mutation journal provides explicit recovery semantics.")]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ServiceName,
        [ValidateSet("Running", "Stopped")]
        [string]$DesiredStatus,
        [ValidateSet("Automatic", "AutomaticDelayedStart", "Manual", "Disabled")]
        [string]$StartupType,
        [string]$Scope = "Run",
        [string]$RecoveryAction = "Restore exact service status and startup mode"
    )

    $snapshot = Get-ServiceSnapshot -ServiceName $ServiceName
    if (-not [bool]$snapshot.Exists) {
        throw "Service '$ServiceName' does not exist"
    }
    if ([string]$snapshot.Status -eq $DesiredStatus -and
        [string]$snapshot.StartupType -eq $StartupType) {
        return ""
    }

    $entryId = Add-MutationJournalEntry -Type "Service" -Target $ServiceName `
        -Before $snapshot -RecoveryAction $RecoveryAction -Scope $Scope
    try {
        Set-ServiceStartupType -ServiceName $ServiceName -StartupType $StartupType
        $service = Get-Service -Name $ServiceName -ErrorAction Stop
        if ($DesiredStatus -eq "Running" -and $service.Status -ne "Running") {
            Start-Service -Name $ServiceName -ErrorAction Stop
        } elseif ($DesiredStatus -eq "Stopped" -and $service.Status -ne "Stopped") {
            Stop-Service -Name $ServiceName -Force -ErrorAction Stop
        }

        $deadline = (Get-Date).AddSeconds(30)
        do {
            $service = Get-Service -Name $ServiceName -ErrorAction Stop
            if ([string]$service.Status -eq $DesiredStatus -and
                (Get-ServiceStartupType -ServiceName $ServiceName) -eq $StartupType) {
                [void](Set-MutationJournalEntryState -EntryId $entryId -State "Applied")
                return $entryId
            }
            Start-Sleep -Milliseconds 250
        } while ((Get-Date) -lt $deadline)
        throw "Service '$ServiceName' did not reach $DesiredStatus/$StartupType"
    } catch {
        [void](Restore-MutationJournalScope -Scope $Scope)
        throw
    }
}

function Test-WindowsUpdateRuntime {
    $missingServices = @()
    foreach ($serviceName in @("wuauserv", "bits", "cryptsvc")) {
        if ($null -eq (Get-Service -Name $serviceName -ErrorAction SilentlyContinue)) {
            $missingServices += $serviceName
        }
    }
    if ($missingServices.Count -gt 0) {
        return [PSCustomObject]@{
            Healthy = $false
            Message = "Required services are missing: $($missingServices -join ', ')"
        }
    }

    try {
        $session = New-Object -ComObject "Microsoft.Update.Session" -ErrorAction Stop
        $searcher = $session.CreateUpdateSearcher()
        [void]$searcher.GetTotalHistoryCount()
        return [PSCustomObject]@{
            Healthy = $true
            Message = "Windows Update Agent diagnostics succeeded"
        }
    } catch {
        return [PSCustomObject]@{
            Healthy = $false
            Message = "Windows Update Agent diagnostics failed: $($_.Exception.Message)"
        }
    }
}

function Repair-WindowsUpdateServices {
    Write-Log "Diagnosing Windows Update services..." "STEP"
    $diagnostics = Test-WindowsUpdateRuntime
    if ($diagnostics.Healthy) {
        Write-Log "$($diagnostics.Message); repair mutations were not needed" "SUCCESS"
        return $true
    }

    Write-Log $diagnostics.Message "WARNING"
    if ($DryRun) {
        Write-Log "Would journal service state, rename the update caches reversibly, validate WUA, and restore exact service state" "INFO"
        return $true
    }

    $scope = "WindowsUpdateRepair"
    $serviceNames = @("wuauserv", "bits", "cryptsvc")
    try {
        Write-Log "Stopping Windows Update services with durable recovery state..." "DEBUG"
        foreach ($serviceName in $serviceNames) {
            [void](Set-JournaledServiceState -ServiceName $serviceName -DesiredStatus "Stopped" `
                -StartupType "Manual" -Scope $scope `
                -RecoveryAction "Restore exact Windows Update service status and startup mode")
        }

        Write-Log "Replacing Windows Update caches with reversible run-scoped directories..." "DEBUG"
        [void](Invoke-JournaledDirectoryReset -Path (Join-Path $script:WindowsRoot "SoftwareDistribution") `
            -Services @("wuauserv", "bits") -Scope $scope)
        [void](Invoke-JournaledDirectoryReset -Path (Join-Path $script:WindowsRoot "System32\catroot2") `
            -Services @("cryptsvc") -Scope $scope)

        Write-Log "Starting Windows Update services for validation..." "DEBUG"
        foreach ($serviceName in $serviceNames) {
            [void](Set-JournaledServiceState -ServiceName $serviceName -DesiredStatus "Running" `
                -StartupType "Manual" -Scope $scope `
                -RecoveryAction "Restore exact Windows Update service status and startup mode")
        }

        $verification = Test-WindowsUpdateRuntime
        if (-not $verification.Healthy) {
            throw "Repair verification failed: $($verification.Message)"
        }
        Write-Log "Windows Update Agent repaired; original cache backups remain recoverable until run finalization" "SUCCESS"
        return $true
    } catch {
        $message = "Windows Update repair failed: $($_.Exception.Message)"
        Write-Log $message "WARNING"
        if (-not (Restore-MutationJournalScope -Scope $scope)) {
            Write-Log "Windows Update repair rollback could not be verified" "ERROR"
        }
        return $false
    }
}

function Set-WSUSBypass {
    param([switch]$Enable)

    $auPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"
    $scope = "WSUS"
    if (-not $Enable) {
        if ($DryRun) { return $true }
        Write-Log "Restoring exact WSUS policy and Windows Update service state..." "DEBUG"
        if (-not (Restore-MutationJournalScope -Scope $scope)) {
            throw "WSUS mutation journal recovery could not be verified"
        }
        return $true
    }

    Write-Log "Configuring temporary WSUS bypass (direct to Microsoft)..." "DEBUG"
    if (-not (Test-Path -LiteralPath $auPath)) {
        Write-Log "No WSUS AU policy key is present; Microsoft Update is already the direct source" "INFO"
        return $true
    }
    if ($DryRun) {
        Write-Log "Would set UseWUServer to 0, restart wuauserv, then restore the exact prior value and service state" "INFO"
        return $true
    }

    try {
        $serviceSnapshot = Get-ServiceSnapshot -ServiceName "wuauserv"
        if (-not [bool]$serviceSnapshot.Exists) { throw "Windows Update service does not exist" }
        $serviceSnapshot.RestartOnRestore = $true
        $serviceEntryId = Add-MutationJournalEntry -Type "Service" -Target "wuauserv" `
            -Before $serviceSnapshot -RecoveryAction "Restore exact wuauserv status/startup mode and restart after WSUS policy recovery" `
            -Scope $scope

        $registrySnapshot = Get-RegistryValueSnapshot -Path $auPath -Name "UseWUServer"
        $registryEntryId = Add-MutationJournalEntry -Type "RegistryValue" `
            -Target "$auPath\UseWUServer" -Before $registrySnapshot `
            -RecoveryAction "Restore the exact UseWUServer value, type, and existence" -Scope $scope

        New-ItemProperty -LiteralPath $auPath -Name "UseWUServer" -Value 0 `
            -PropertyType DWord -Force -ErrorAction Stop | Out-Null
        $current = Get-RegistryValueSnapshot -Path $auPath -Name "UseWUServer"
        if (-not $current.Exists -or [string]$current.Kind -ne "DWord" -or [int]$current.Value -ne 0) {
            throw "UseWUServer write verification failed"
        }
        [void](Set-MutationJournalEntryState -EntryId $registryEntryId -State "Applied")

        Restart-Service -Name "wuauserv" -Force -ErrorAction Stop
        $deadline = (Get-Date).AddSeconds(30)
        do {
            $service = Get-Service -Name "wuauserv" -ErrorAction Stop
            if ($service.Status -eq "Running") { break }
            Start-Sleep -Milliseconds 250
        } while ((Get-Date) -lt $deadline)
        if ($service.Status -ne "Running") {
            throw "wuauserv did not return to Running after the policy change"
        }
        [void](Set-MutationJournalEntryState -EntryId $serviceEntryId -State "Applied")
        Write-Log "WSUS bypass enabled with a durable exact-state recovery record" "SUCCESS"
        return $true
    } catch {
        [void](Restore-MutationJournalScope -Scope $scope)
        throw
    }
}

# ============================================================================
# POST-REBOOT CONTINUATION
# ============================================================================

function Get-ScheduledTaskSnapshot {
    param(
        [string]$TaskName = $script:TaskName,
        [string]$TaskPath = "\"
    )

    if ($TaskName -ne $script:TaskName -or $TaskPath -ne "\") {
        throw "Scheduled-task snapshot target is outside the continuation-task contract"
    }
    $task = Get-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -ErrorAction SilentlyContinue
    return [ordered]@{
        Exists   = ($null -ne $task)
        TaskName = $TaskName
        TaskPath = $TaskPath
        Xml      = $(if ($null -ne $task) {
            [string](Export-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -ErrorAction Stop)
        } else { "" })
    }
}

function Restore-ScheduledTaskSnapshot {
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.IDictionary]$Snapshot
    )

    if ([string]$Snapshot.TaskName -ne $script:TaskName -or [string]$Snapshot.TaskPath -ne "\") {
        return $false
    }
    try {
        $current = Get-ScheduledTask -TaskName ([string]$Snapshot.TaskName) `
            -TaskPath ([string]$Snapshot.TaskPath) -ErrorAction SilentlyContinue
        if ($null -ne $current) {
            Unregister-ScheduledTask -TaskName ([string]$Snapshot.TaskName) `
                -TaskPath ([string]$Snapshot.TaskPath) -Confirm:$false -ErrorAction Stop
        }
        if ([bool]$Snapshot.Exists) {
            if ([string]::IsNullOrWhiteSpace([string]$Snapshot.Xml)) {
                throw "Saved scheduled-task XML is missing"
            }
            Register-ScheduledTask -TaskName ([string]$Snapshot.TaskName) `
                -TaskPath ([string]$Snapshot.TaskPath) -Xml ([string]$Snapshot.Xml) `
                -Force -ErrorAction Stop | Out-Null
        }

        $restored = Get-ScheduledTask -TaskName ([string]$Snapshot.TaskName) `
            -TaskPath ([string]$Snapshot.TaskPath) -ErrorAction SilentlyContinue
        if (-not [bool]$Snapshot.Exists) { return ($null -eq $restored) }
        if ($null -eq $restored) { return $false }
        $restoredXml = [string](Export-ScheduledTask -TaskName ([string]$Snapshot.TaskName) `
            -TaskPath ([string]$Snapshot.TaskPath) -ErrorAction Stop)
        $expectedNormalized = [regex]::Replace(([string]$Snapshot.Xml).Trim(), ">\s+<", "><")
        $actualNormalized = [regex]::Replace($restoredXml.Trim(), ">\s+<", "><")
        return ($expectedNormalized -ceq $actualNormalized)
    } catch {
        return $false
    }
}

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
        $taskSnapshot = Get-ScheduledTaskSnapshot
        $taskEntryId = Add-MutationJournalEntry -Type "ScheduledTask" `
            -Target "\$($script:TaskName)" -Before $taskSnapshot `
            -RecoveryAction "Remove the run task and restore the exact preexisting task XML or absence" `
            -Scope "Continuation"

        # Keep the task command line free of saved parameters and webhook URLs;
        # the validated state file is the sole resume contract.
        if ([bool]$taskSnapshot.Exists) {
            Unregister-ScheduledTask -TaskName $script:TaskName -TaskPath "\" `
                -Confirm:$false -ErrorAction Stop
        }

        $powershellPath = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"
        $arguments = "-NoProfile -NonInteractive -ExecutionPolicy Bypass -File `"$scriptPath`""
        $action = New-ScheduledTaskAction -Execute $powershellPath -Argument $arguments
        $trigger = New-ScheduledTaskTrigger -AtStartup
        $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
        $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
            -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Hours 4)

        Register-ScheduledTask -TaskName $script:TaskName -TaskPath "\" -Action $action `
            -Trigger $trigger -Principal $principal -Settings $settings -Force `
            -ErrorAction Stop | Out-Null
        if (-not (Test-ContinuationTask)) {
            throw "Scheduled task registration verification failed"
        }
        [void](Set-MutationJournalEntryState -EntryId $taskEntryId -State "Applied")

        $state.Phase = "AwaitingReboot"
        if (-not (Save-State -State $state)) {
            throw "Scheduled task was created but final state commit failed"
        }

        $script:ContinuationState = $state
        $script:ContinuationRegistered = $true
        Write-Log "Continuation task registered: $($script:TaskName)" "SUCCESS"
        return $true
    } catch {
        $taskRecovered = Restore-MutationJournalScope -Scope "Continuation"
        if ($taskRecovered) { [void](Complete-MutationJournal) }
        [void](Clear-State)
        $script:ContinuationRegistered = $false
        Write-Log "Failed to create continuation task: $($_.Exception.Message)" "WARNING"
        return $false
    }
}

function Test-ContinuationTask {
    try {
        return $null -ne (Get-ScheduledTask -TaskName $script:TaskName -TaskPath "\" -ErrorAction SilentlyContinue)
    } catch {
        return $false
    }
}

function Unregister-ContinuationTask {
    param([switch]$PreserveState)

    $success = $true
    try {
        $continuationEntries = @()
        if ($null -eq $script:MutationJournal) {
            $script:MutationJournal = Get-MutationJournal -RunId $script:RunId
        }
        if ($null -ne $script:MutationJournal) {
            $continuationEntries = @($script:MutationJournal.Entries | Where-Object {
                [string]$_.Scope -eq "Continuation" -and
                [string]$_.State -notin @("Restored", "Committed")
            })
        }

        if ($continuationEntries.Count -gt 0) {
            $success = Restore-MutationJournalScope -Scope "Continuation"
        } elseif (Test-ContinuationTask) {
            Unregister-ScheduledTask -TaskName $script:TaskName -TaskPath "\" `
                -Confirm:$false -ErrorAction Stop
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

    $key = Get-Item -LiteralPath $Path -ErrorAction SilentlyContinue
    if ($null -eq $key) {
        return [ordered]@{
            Path = $Path; Name = $Name; KeyExists = $false
            Exists = $false; Value = $null; Kind = ""
        }
    }
    $valueNames = @($key.GetValueNames())
    $exists = $valueNames -contains $Name
    return [ordered]@{
        Path      = $Path
        Name      = $Name
        KeyExists = $true
        Exists    = $exists
        Value     = $(if ($exists) { $key.GetValue($Name, $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames) } else { $null })
        Kind      = $(if ($exists) { $key.GetValueKind($Name).ToString() } else { "" })
    }
}

function Test-MutationRegistryTarget {
    param(
        [string]$Path,
        [string]$Name
    )

    $normalizedPath = $Path.TrimEnd("\")
    if ($normalizedPath -eq "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" -and
        $Name -eq "UseWUServer") {
        return $true
    }
    return (
        $normalizedPath.StartsWith(
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\",
            [StringComparison]::OrdinalIgnoreCase
        ) -and $Name -eq "StateFlags0100"
    )
}

function ConvertTo-RegistryValueKind {
    param(
        [AllowNull()][object]$Value,
        [string]$Kind
    )

    switch ($Kind) {
        "Binary" {
            [byte[]]$bytes = @($Value | ForEach-Object { [byte]$_ })
            return ,$bytes
        }
        "MultiString" {
            [string[]]$strings = @($Value | ForEach-Object { [string]$_ })
            return ,$strings
        }
        "DWord" { return [int]$Value }
        "QWord" { return [long]$Value }
        "String" { return [string]$Value }
        "ExpandString" { return [string]$Value }
        default { throw "Unsupported registry value kind '$Kind'" }
    }
}

function Test-RegistrySnapshotEqual {
    param(
        [System.Collections.IDictionary]$Expected,
        [System.Collections.IDictionary]$Actual
    )

    if ([bool]$Expected.KeyExists -ne [bool]$Actual.KeyExists -or
        [bool]$Expected.Exists -ne [bool]$Actual.Exists) {
        return $false
    }
    if (-not [bool]$Expected.Exists) { return $true }
    if ([string]$Expected.Kind -ne [string]$Actual.Kind) { return $false }

    if ([string]$Expected.Kind -in @("Binary", "MultiString")) {
        $expectedValues = @($Expected.Value)
        $actualValues = @($Actual.Value)
        if ($expectedValues.Count -ne $actualValues.Count) { return $false }
        for ($index = 0; $index -lt $expectedValues.Count; $index++) {
            if ([string]$expectedValues[$index] -cne [string]$actualValues[$index]) {
                return $false
            }
        }
        return $true
    }
    return ([string]$Expected.Value -ceq [string]$Actual.Value)
}

function Restore-RegistryValueSnapshot {
    param([System.Collections.IDictionary]$Snapshot)

    if (-not (Test-MutationRegistryTarget -Path ([string]$Snapshot.Path) -Name ([string]$Snapshot.Name))) {
        return $false
    }
    try {
        if ([bool]$Snapshot.Exists) {
            if (-not (Test-Path -LiteralPath $Snapshot.Path)) {
                New-Item -Path $Snapshot.Path -Force -ErrorAction Stop | Out-Null
            }
            $value = ConvertTo-RegistryValueKind -Value $Snapshot.Value -Kind ([string]$Snapshot.Kind)
            New-ItemProperty -LiteralPath $Snapshot.Path -Name $Snapshot.Name -Value $value `
                -PropertyType $Snapshot.Kind -Force -ErrorAction Stop | Out-Null
        } else {
            Remove-ItemProperty -LiteralPath $Snapshot.Path -Name $Snapshot.Name -Force -ErrorAction SilentlyContinue
            if (-not [bool]$Snapshot.KeyExists -and (Test-Path -LiteralPath $Snapshot.Path)) {
                $key = Get-Item -LiteralPath $Snapshot.Path -ErrorAction Stop
                if (@($key.GetValueNames()).Count -eq 0 -and @($key.GetSubKeyNames()).Count -eq 0) {
                    Remove-Item -LiteralPath $Snapshot.Path -Force -ErrorAction Stop
                }
            }
        }

        $restored = Get-RegistryValueSnapshot -Path $Snapshot.Path -Name $Snapshot.Name
        return Test-RegistrySnapshotEqual -Expected $Snapshot -Actual $restored
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
    $flagFailures = [System.Collections.ArrayList]::new()
    $cleanMgrSucceeded = $false
    $cleanupScope = "Cleanup"

    try {
        foreach ($item in $cleanupItems) {
            $itemPath = Join-Path $volCachePath $item
            if (-not (Test-Path -LiteralPath $itemPath)) { continue }

            try {
                $snapshot = Get-RegistryValueSnapshot -Path $itemPath -Name "StateFlags0100"
                $entryId = Add-MutationJournalEntry -Type "RegistryValue" `
                    -Target "$itemPath\StateFlags0100" -Before $snapshot `
                    -RecoveryAction "Restore the exact cleanmgr StateFlags0100 value, type, and existence" `
                    -Scope $cleanupScope
                New-ItemProperty -LiteralPath $itemPath -Name "StateFlags0100" -Value 2 `
                    -PropertyType DWord -Force -ErrorAction Stop | Out-Null
                $written = Get-RegistryValueSnapshot -Path $itemPath -Name "StateFlags0100"
                if (-not $written.Exists -or [string]$written.Kind -ne "DWord" -or
                    [int]$written.Value -ne 2) {
                    throw "StateFlags0100 write verification failed"
                }
                [void](Set-MutationJournalEntryState -EntryId $entryId -State "Applied")
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
        if (-not (Restore-MutationJournalScope -Scope $cleanupScope)) {
            [void]$flagFailures.Add("Could not restore one or more cleanmgr registry values")
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
# VERIFIED DEPENDENCY ACQUISITION
# ============================================================================

function Get-AcquisitionManifest {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseShouldProcessForStateChangingFunctions", "", Justification = "Returns immutable acquisition metadata only.")]
    param()

    return @{
        WinGet = [ordered]@{
            Name             = "WinGet"
            Kind             = "AppxBundle"
            Uri              = "https://github.com/microsoft/winget-cli/releases/download/v1.29.280/Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle"
            AllowedHosts     = @("github.com", "release-assets.githubusercontent.com")
            ExactVersion     = "1.29.280"
            MinimumVersion   = "1.29.280"
            Architectures    = @("x64", "x86", "arm64")
            Sha256           = "0809FA9F52E395D6E7DE692331DCE847AC991952675116BB4D8AAE2DDCC20946"
            PublisherPattern = "^CN=Microsoft Corporation,"
            DependencyBundle = [ordered]@{
                Uri          = "https://github.com/microsoft/winget-cli/releases/download/v1.29.280/DesktopAppInstaller_Dependencies.zip"
                AllowedHosts = @("github.com", "release-assets.githubusercontent.com")
                Sha256       = "3BBFCAA5CB011C48FAC48D896D64A5C7C6898859A9F3D01555C8CD000F4E2962"
                Packages     = @(
                    "Microsoft.VCLibs.140.00_14.0.33519.0_{0}.appx",
                    "Microsoft.VCLibs.140.00.UWPDesktop_14.0.33728.0_{0}.appx",
                    "Microsoft.WindowsAppRuntime.1.8_8000.616.304.0_{0}.appx"
                )
                Versions     = @("14.0.33519.0", "14.0.33728.0", "8000.616.304.0")
            }
            License = [ordered]@{
                Uri          = "https://github.com/microsoft/winget-cli/releases/download/v1.29.280/e53e159d00e04f729cc2180cffd1c02e_License1.xml"
                AllowedHosts = @("github.com", "release-assets.githubusercontent.com")
                Sha256       = "BCB15118EC47DF24E3E6013A7006147C3A15B3A8104ED660FE87C8B4ED01F485"
            }
        }
        PSWindowsUpdate = [ordered]@{
            Name              = "PSWindowsUpdate"
            Kind              = "PowerShellModule"
            Uri               = "https://www.powershellgallery.com/api/v2/package/PSWindowsUpdate/2.2.1.5"
            AllowedHosts      = @("www.powershellgallery.com", "cdn.powershellgallery.com")
            ExactVersion      = "2.2.1.5"
            MinimumVersion    = "2.2.1.5"
            Architectures     = @("neutral")
            Sha256            = "174E05B1C194377F8D9AE0B004C93304C662EA9B5FC4DAFA19426049D2D5CF50"
            PublisherPattern  = ""
            PayloadTreeSha256 = "B0C4D34C5CDD931459EBCE0A509050A1C9AC305DDCF64978EDB86478DBCF2E62"
        }
        LSUClient = [ordered]@{
            Name              = "LSUClient"
            Kind              = "PowerShellModule"
            Uri               = "https://www.powershellgallery.com/api/v2/package/LSUClient/1.8.1"
            AllowedHosts      = @("www.powershellgallery.com", "cdn.powershellgallery.com")
            ExactVersion      = "1.8.1"
            MinimumVersion    = "1.8.1"
            Architectures     = @("neutral")
            Sha256            = "05F8DC57356FED994EB69D232B4C425DD450063352A5EED6A20C03C618848C94"
            PublisherPattern  = ""
            PayloadTreeSha256 = "CE2E9D22A7345C9CE01A7486E1E7E7FE868F88698957413F5D235182548FC96F"
        }
        DellCommandUpdate = [ordered]@{
            Name                       = "Dell Command Update"
            Kind                       = "WinGetPackage"
            Uri                        = "https://dl.dell.com/FOLDER14424243M/1/Dell-Command-Update-Application_RXT5N_WIN64_5.7.0_A00.EXE"
            AllowedHosts                = @("dl.dell.com")
            ExactVersion                = "5.7.0"
            MinimumVersion              = "5.7.0"
            Architectures               = @("x64")
            Sha256                      = "B6D0D06EDF25A7D9092208B2686502E031DEDA5CCB2D283BA8A434B27891F3CF"
            PublisherPattern            = "^CN=Dell (?:Technologies )?Inc\.,"
            PublisherDisplayName        = "Dell Inc."
            PackageId                   = "Dell.CommandUpdate"
            PackageSource               = "winget"
            InventoryCollectorPath      = "C:\Program Files (x86)\Dell\UpdateService\Service\InvColPC\invcol.exe"
            InventoryCollectorMinimum   = "13.8.0"
            InventoryPublisherPattern   = "^CN=Dell (?:Technologies )?Inc\.,"
        }
        HPIA = [ordered]@{
            Name                    = "HP Image Assistant"
            Kind                    = "SelfExtractingArchive"
            Uri                     = "https://hpia.hpcloud.hp.com/downloads/hpia/hp-hpia-5.3.6.exe"
            AllowedHosts            = @("hpia.hpcloud.hp.com")
            ExactVersion            = "5.3.6"
            MinimumVersion          = "5.3.3"
            Architectures           = @("x64", "arm64")
            Sha256                  = "5E205A0300C1DC4F59D8A08CAA5FB1DAF434F1DC6CB05DF4E4D3F9EC83CD9CB7"
            PublisherPattern        = "^CN=HP Inc\.,"
            InstalledSha256         = "3A2658848E270F99ABA5FC3DAA4851891238A73A0C684A8404E547F23EF3DD85"
            InstalledMinimumVersion   = "5.3.3"
            InstalledPublisherPattern = "^CN=HP Inc\.,"
        }
    }
}

function Get-AcquisitionManifestEntry {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    if ($null -eq $script:AcquisitionManifest) {
        $script:AcquisitionManifest = Get-AcquisitionManifest
    }
    if (-not $script:AcquisitionManifest.ContainsKey($Name)) {
        throw "Dependency '$Name' is not present in the acquisition manifest"
    }
    return $script:AcquisitionManifest[$Name]
}

function Get-DependencyCacheArtifactPath {
    param(
        [Parameter(Mandatory = $true)]
        [ValidatePattern("^[A-Fa-f0-9]{64}$")]
        [string]$ExpectedSha256,
        [string]$CachePath = ""
    )

    if ([string]::IsNullOrWhiteSpace($CachePath)) {
        $CachePath = [string]$script:DependencyCachePath
    }
    if ([string]::IsNullOrWhiteSpace($CachePath)) { return $null }
    try {
        $root = [IO.Path]::GetFullPath($CachePath).TrimEnd("\", "/")
        return Join-Path (Join-Path $root "sha256") $ExpectedSha256.ToUpperInvariant()
    } catch {
        return $null
    }
}

function Test-DependencyCacheArtifact {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,
        [Parameter(Mandatory = $true)]
        [ValidatePattern("^[A-Fa-f0-9]{64}$")]
        [string]$ExpectedSha256,
        [string]$PublisherPattern = "",
        [string]$CachePath = ""
    )

    $artifactPath = Get-DependencyCacheArtifactPath -ExpectedSha256 $ExpectedSha256 -CachePath $CachePath
    if ([string]::IsNullOrWhiteSpace($artifactPath)) {
        return [PSCustomObject]@{
            Valid = $false; Name = $Name; Path = ""; Sha256 = ""; Publisher = ""; Thumbprint = ""
            Status = "Unavailable"; Reason = "Dependency cache path is not configured"
        }
    }
    if (-not (Test-Path -LiteralPath $artifactPath -PathType Leaf)) {
        return [PSCustomObject]@{
            Valid = $false; Name = $Name; Path = $artifactPath; Sha256 = ""; Publisher = ""; Thumbprint = ""
            Status = "CacheMiss"; Reason = "No content-addressed artifact matches the expected SHA-256"
        }
    }

    try {
        $verification = Test-AcquiredFile -Path $artifactPath -ExpectedSha256 $ExpectedSha256 `
            -PublisherPattern $PublisherPattern
        return [PSCustomObject]@{
            Valid = [bool]$verification.Valid
            Name = $Name
            Path = $artifactPath
            Sha256 = [string]$verification.Sha256
            Publisher = [string]$verification.Publisher
            Thumbprint = [string]$verification.Thumbprint
            Status = $(if ($verification.Valid) { "Ready" } else { "Rejected" })
            Reason = [string]$verification.Reason
        }
    } catch {
        return [PSCustomObject]@{
            Valid = $false; Name = $Name; Path = $artifactPath; Sha256 = ""; Publisher = ""; Thumbprint = ""
            Status = "Rejected"; Reason = Protect-EvidenceText -Text $_.Exception.Message
        }
    }
}

function Get-SourceReadiness {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,
        [Parameter(Mandatory = $true)]
        [string]$Uri,
        [Parameter(Mandatory = $true)]
        [string[]]$AllowedHosts,
        [string]$ExpectedSha256 = "",
        [string]$PublisherPattern = "",
        [bool]$OfflineMode = $false,
        [string]$CachePath = "",
        [ValidateRange(1, 600)]
        [int]$TimeoutSeconds = 30
    )

    if (-not $PSBoundParameters.ContainsKey("OfflineMode")) {
        $OfflineMode = [bool]$script:Offline
    }
    if (-not $PSBoundParameters.ContainsKey("CachePath")) {
        $CachePath = [string]$script:DependencyCachePath
    }
    if (-not $PSBoundParameters.ContainsKey("TimeoutSeconds") -and [int]$script:SourceTimeoutSeconds -gt 0) {
        $TimeoutSeconds = [int]$script:SourceTimeoutSeconds
    }

    $cacheEvidence = $null
    if (-not [string]::IsNullOrWhiteSpace($ExpectedSha256)) {
        $cacheEvidence = Test-DependencyCacheArtifact -Name $Name -ExpectedSha256 $ExpectedSha256 `
            -PublisherPattern $PublisherPattern -CachePath $CachePath
    }
    $base = [ordered]@{
        Name = $Name
        Uri = $Uri
        Host = ""
        Status = "Unknown"
        Ready = $false
        Offline = $OfflineMode
        Cache = $cacheEvidence
        TimeoutSeconds = $TimeoutSeconds
        Proxy = "System default"
        HttpStatus = 0
        Reason = ""
        EvaluatedAt = (Get-Date).ToUniversalTime().ToString("o")
    }

    try {
        $parsed = [Uri]$Uri
        $base.Host = $parsed.DnsSafeHost
    } catch {}

    if (-not (Test-AcquisitionUri -Uri $Uri -AllowedHosts $AllowedHosts)) {
        $base.Status = "Blocked"
        $base.Reason = "Source URI is not HTTPS on an approved origin"
        return [PSCustomObject]$base
    }
    if ($cacheEvidence -and $cacheEvidence.Valid -and $OfflineMode) {
        $base.Status = "OfflineCache"
        $base.Ready = $true
        $base.Reason = "Verified content-addressed cache artifact is available; network access is disabled"
        return [PSCustomObject]$base
    }
    if ($OfflineMode) {
        $base.Status = "Unavailable"
        $base.Reason = if ($cacheEvidence) {
            "Offline mode requires a matching verified cache artifact: $($cacheEvidence.Reason)"
        } else {
            "Offline mode requires a verified cache artifact"
        }
        return [PSCustomObject]$base
    }

    try {
        $response = Invoke-WebRequest -Uri $Uri -Method Head -TimeoutSec $TimeoutSeconds `
            -MaximumRedirection 0 -UseBasicParsing -ErrorAction Stop
        $base.HttpStatus = [int]$response.StatusCode
        if ($base.HttpStatus -ge 200 -and $base.HttpStatus -lt 400) {
            $base.Status = "Ready"
            $base.Ready = $true
            $base.Reason = "Approved origin responded to a bounded readiness probe"
        } else {
            $base.Status = "Unavailable"
            $base.Reason = "Approved origin returned HTTP $($base.HttpStatus)"
        }
    } catch {
        $statusCode = 0
        try {
            if ($null -ne $_.Exception.Response) { $statusCode = [int]$_.Exception.Response.StatusCode }
        } catch {}
        $base.HttpStatus = $statusCode
        $base.Status = "Unavailable"
        $base.Reason = "Source readiness probe failed after $TimeoutSeconds seconds: $(Protect-EvidenceText -Text $_.Exception.Message)"
    }
    return [PSCustomObject]$base
}

function Get-DependencyReadiness {
    param(
        [string[]]$Names = @(),
        [string]$CachePath = "",
        [bool]$OfflineMode = $false,
        [ValidateRange(1, 600)]
        [int]$TimeoutSeconds = 30
    )

    if ($Names.Count -eq 0) {
        $Names = @("PSWindowsUpdate", "WinGet", "LSUClient", "DellCommandUpdate", "HPIA")
    }
    $sources = [System.Collections.ArrayList]::new()
    $providers = [ordered]@{}
    foreach ($name in $Names) {
        $spec = Get-AcquisitionManifestEntry -Name $name
        $providerSources = [System.Collections.ArrayList]::new()
        $sourceSpecs = @([ordered]@{
            Name = [string]$spec.Name
            Uri = [string]$spec.Uri
            AllowedHosts = @($spec.AllowedHosts)
            Sha256 = [string]$spec.Sha256
            PublisherPattern = [string]$spec.PublisherPattern
        })
        if ($name -eq "WinGet") {
            $sourceSpecs += [ordered]@{
                Name = "WinGet dependency bundle"
                Uri = [string]$spec.DependencyBundle.Uri
                AllowedHosts = @($spec.DependencyBundle.AllowedHosts)
                Sha256 = [string]$spec.DependencyBundle.Sha256
                PublisherPattern = [string]$spec.PublisherPattern
            }
            $sourceSpecs += [ordered]@{
                Name = "WinGet license"
                Uri = [string]$spec.License.Uri
                AllowedHosts = @($spec.License.AllowedHosts)
                Sha256 = [string]$spec.License.Sha256
                PublisherPattern = ""
            }
        }
        foreach ($sourceSpec in $sourceSpecs) {
            $source = Get-SourceReadiness -Name ([string]$sourceSpec.Name) `
                -Uri ([string]$sourceSpec.Uri) -AllowedHosts @($sourceSpec.AllowedHosts) `
                -ExpectedSha256 ([string]$sourceSpec.Sha256) `
                -PublisherPattern ([string]$sourceSpec.PublisherPattern) `
                -OfflineMode $OfflineMode -CachePath $CachePath -TimeoutSeconds $TimeoutSeconds
            [void]$providerSources.Add($source)
            [void]$sources.Add($source)
        }
        $providers[$name] = [ordered]@{
            Name = $name
            Status = if (@($providerSources | Where-Object { -not $_.Ready }).Count -eq 0) { "Ready" } else { "Unavailable" }
            Sources = @($providerSources)
            Reason = @($providerSources | Where-Object { -not $_.Ready } | ForEach-Object { $_.Reason }) -join "; "
        }
    }
    return [PSCustomObject][ordered]@{
        SchemaVersion = 1
        EvaluatedAt = (Get-Date).ToUniversalTime().ToString("o")
        Mode = if ($OfflineMode) { "Offline" } else { "Online" }
        CachePath = [string]$CachePath
        TimeoutSeconds = $TimeoutSeconds
        Ready = (@($sources | Where-Object { -not $_.Ready }).Count -eq 0)
        Sources = @($sources)
        Providers = $providers
    }
}

function Get-SystemArchitecture {
    $architecture = [string]$env:PROCESSOR_ARCHITEW6432
    if ([string]::IsNullOrWhiteSpace($architecture)) {
        $architecture = [string]$env:PROCESSOR_ARCHITECTURE
    }

    switch -Regex ($architecture.ToUpperInvariant()) {
        "^(AMD64|X64)$" { return "x64" }
        "^ARM64$"       { return "arm64" }
        "^X86$"         { return "x86" }
        default         { return "unknown" }
    }
}

function Test-AcquisitionUri {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri,
        [Parameter(Mandatory = $true)]
        [string[]]$AllowedHosts
    )

    try {
        $parsed = [Uri]$Uri
        if (-not $parsed.IsAbsoluteUri -or $parsed.Scheme -ne "https") { return $false }
        if (-not $parsed.IsDefaultPort -and $parsed.Port -ne 443) { return $false }
        if (-not [string]::IsNullOrWhiteSpace($parsed.UserInfo)) { return $false }
        $hostName = $parsed.DnsSafeHost.ToLowerInvariant()
        return (@($AllowedHosts | ForEach-Object { $_.ToLowerInvariant() }) -contains $hostName)
    } catch {
        return $false
    }
}

function ConvertTo-SafeVersion {
    param([AllowNull()][object]$Value)

    $text = [string]$Value
    if ($text -match "(\d+(?:\.\d+){1,3})") {
        try { return [Version]$Matches[1] } catch { return $null }
    }
    return $null
}

function Test-VersionAtLeast {
    param(
        [AllowNull()][object]$Version,
        [Parameter(Mandatory = $true)]
        [string]$MinimumVersion
    )

    $actual = ConvertTo-SafeVersion -Value $Version
    $minimum = ConvertTo-SafeVersion -Value $MinimumVersion
    return ($null -ne $actual -and $null -ne $minimum -and $actual -ge $minimum)
}

function Get-AuthenticodeEvidence {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter(Mandatory = $true)]
        [string]$PublisherPattern
    )

    try {
        $signature = Get-AuthenticodeSignature -LiteralPath $Path -ErrorAction Stop
        $subject = if ($signature.SignerCertificate) { [string]$signature.SignerCertificate.Subject } else { "" }
        $thumbprint = if ($signature.SignerCertificate) { [string]$signature.SignerCertificate.Thumbprint } else { "" }
        if ($signature.Status -ne [System.Management.Automation.SignatureStatus]::Valid) {
            return [PSCustomObject]@{
                Valid = $false; Status = [string]$signature.Status; Subject = $subject
                Thumbprint = $thumbprint; Reason = "Authenticode status is $($signature.Status)"
            }
        }
        if ($subject -notmatch $PublisherPattern) {
            return [PSCustomObject]@{
                Valid = $false; Status = [string]$signature.Status; Subject = $subject
                Thumbprint = $thumbprint; Reason = "Signer '$subject' does not match the approved publisher"
            }
        }
        return [PSCustomObject]@{
            Valid = $true; Status = [string]$signature.Status; Subject = $subject
            Thumbprint = $thumbprint; Reason = ""
        }
    } catch {
        return [PSCustomObject]@{
            Valid = $false; Status = "Error"; Subject = ""; Thumbprint = ""
            Reason = "Authenticode verification failed: $($_.Exception.Message)"
        }
    }
}

function Test-AcquiredFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter(Mandatory = $true)]
        [ValidatePattern("^[A-Fa-f0-9]{64}$")]
        [string]$ExpectedSha256,
        [string]$PublisherPattern = ""
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return [PSCustomObject]@{
            Valid = $false; Path = $Path; Sha256 = ""; Publisher = ""; Thumbprint = ""
            Reason = "Downloaded file is missing"
        }
    }

    try {
        $actualHash = (Get-FileHash -LiteralPath $Path -Algorithm SHA256 -ErrorAction Stop).Hash.ToUpperInvariant()
    } catch {
        return [PSCustomObject]@{
            Valid = $false; Path = $Path; Sha256 = ""; Publisher = ""; Thumbprint = ""
            Reason = "SHA-256 verification failed: $($_.Exception.Message)"
        }
    }

    if ($actualHash -ne $ExpectedSha256.ToUpperInvariant()) {
        return [PSCustomObject]@{
            Valid = $false; Path = $Path; Sha256 = $actualHash; Publisher = ""; Thumbprint = ""
            Reason = "SHA-256 mismatch"
        }
    }

    $publisher = ""
    $thumbprint = ""
    if (-not [string]::IsNullOrWhiteSpace($PublisherPattern)) {
        $signature = Get-AuthenticodeEvidence -Path $Path -PublisherPattern $PublisherPattern
        if (-not $signature.Valid) {
            return [PSCustomObject]@{
                Valid = $false; Path = $Path; Sha256 = $actualHash; Publisher = $signature.Subject
                Thumbprint = $signature.Thumbprint; Reason = $signature.Reason
            }
        }
        $publisher = $signature.Subject
        $thumbprint = $signature.Thumbprint
    }

    return [PSCustomObject]@{
        Valid = $true; Path = $Path; Sha256 = $actualHash; Publisher = $publisher
        Thumbprint = $thumbprint; Reason = ""
    }
}

function Invoke-VerifiedDownload {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,
        [Parameter(Mandatory = $true)]
        [string]$Uri,
        [Parameter(Mandatory = $true)]
        [string[]]$AllowedHosts,
        [Parameter(Mandatory = $true)]
        [string]$ExpectedSha256,
        [Parameter(Mandatory = $true)]
        [string]$DestinationPath,
        [string]$PublisherPattern = "",
        [int]$MaximumRedirects = 8,
        [bool]$OfflineMode = $false,
        [string]$CachePath = "",
        [ValidateRange(1, 600)]
        [int]$TimeoutSeconds = 30
    )

    if (-not (Test-AcquisitionUri -Uri $Uri -AllowedHosts $AllowedHosts)) {
        throw "$Name has an unapproved acquisition URI: $Uri"
    }
    if (Test-Path -LiteralPath $DestinationPath) {
        throw "$Name download destination already exists: $DestinationPath"
    }

    $destinationDirectory = Split-Path -Parent $DestinationPath
    if (-not (Test-Path -LiteralPath $destinationDirectory)) {
        New-Item -ItemType Directory -Path $destinationDirectory -Force -ErrorAction Stop | Out-Null
    }

    if ([string]::IsNullOrWhiteSpace($CachePath)) {
        $CachePath = [string]$script:DependencyCachePath
    }
    $cacheArtifact = Test-DependencyCacheArtifact -Name $Name -ExpectedSha256 $ExpectedSha256 `
        -PublisherPattern $PublisherPattern -CachePath $CachePath
    if ($cacheArtifact.Valid) {
        try {
            [IO.File]::Copy($cacheArtifact.Path, $DestinationPath, $false)
            $cachedVerification = Test-AcquiredFile -Path $DestinationPath -ExpectedSha256 $ExpectedSha256 `
                -PublisherPattern $PublisherPattern
            if (-not $cachedVerification.Valid) {
                throw "$Name cache copy failed verification: $($cachedVerification.Reason)"
            }
            $cachedVerification.Path = $DestinationPath
            return $cachedVerification
        } catch {
            Remove-Item -LiteralPath $DestinationPath -Force -ErrorAction SilentlyContinue
            if ($OfflineMode) { throw }
        }
    }
    if ($OfflineMode) {
        throw "$Name is unavailable in offline mode: $($cacheArtifact.Reason)"
    }
    if (-not (Test-DownloadAllowed)) {
        throw "$Name was deferred by the network download policy: $($script:DownloadPolicy.Reason)"
    }

    $partialPath = "$DestinationPath.partial.$([guid]::NewGuid().ToString('N'))"
    $handler = $null
    $client = $null
    $response = $null

    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Add-Type -AssemblyName System.Net.Http -ErrorAction Stop
        $handler = New-Object System.Net.Http.HttpClientHandler
        $handler.AllowAutoRedirect = $false
        $client = New-Object System.Net.Http.HttpClient($handler)
        $client.Timeout = [TimeSpan]::FromSeconds($TimeoutSeconds)
        $client.DefaultRequestHeaders.UserAgent.ParseAdd("SystemUpdatePro/$($script:Version)")

        $currentUri = [Uri]$Uri
        $redirectCount = 0
        while ($true) {
            if (-not (Test-AcquisitionUri -Uri $currentUri.AbsoluteUri -AllowedHosts $AllowedHosts)) {
                throw "$Name redirected to an unapproved origin: $($currentUri.AbsoluteUri)"
            }

            $response = $client.GetAsync(
                $currentUri,
                [System.Net.Http.HttpCompletionOption]::ResponseHeadersRead
            ).GetAwaiter().GetResult()
            $statusCode = [int]$response.StatusCode
            if ($statusCode -in @(301, 302, 303, 307, 308)) {
                if ($redirectCount -ge $MaximumRedirects) {
                    throw "$Name exceeded $MaximumRedirects HTTPS redirects"
                }
                $location = $response.Headers.Location
                if ($null -eq $location) { throw "$Name returned a redirect without a Location header" }
                $nextUri = if ($location.IsAbsoluteUri) { $location } else { New-Object Uri($currentUri, $location) }
                $response.Dispose()
                $response = $null
                $currentUri = $nextUri
                $redirectCount++
                continue
            }

            if (-not $response.IsSuccessStatusCode) {
                throw "$Name download returned HTTP $statusCode"
            }

            $inputStream = $response.Content.ReadAsStreamAsync().GetAwaiter().GetResult()
            try {
                $outputStream = New-Object System.IO.FileStream(
                    $partialPath,
                    [System.IO.FileMode]::CreateNew,
                    [System.IO.FileAccess]::Write,
                    [System.IO.FileShare]::None
                )
                try {
                    $inputStream.CopyTo($outputStream)
                    $outputStream.Flush()
                } finally {
                    $outputStream.Dispose()
                }
            } finally {
                $inputStream.Dispose()
            }
            break
        }

        $verification = Test-AcquiredFile -Path $partialPath -ExpectedSha256 $ExpectedSha256 `
            -PublisherPattern $PublisherPattern
        if (-not $verification.Valid) {
            throw "$Name was rejected before execution: $($verification.Reason)"
        }

        Move-Item -LiteralPath $partialPath -Destination $DestinationPath -ErrorAction Stop
        $verification.Path = $DestinationPath
        return $verification
    } finally {
        if ($response) { $response.Dispose() }
        if ($client) { $client.Dispose() }
        if ($handler) { $handler.Dispose() }
        Remove-Item -LiteralPath $partialPath -Force -ErrorAction SilentlyContinue
    }
}

function Add-AcquisitionProvenance {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseShouldProcessForStateChangingFunctions", "", Justification = "Updates in-memory run evidence only.")]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,
        [Parameter(Mandatory = $true)]
        [string]$Version,
        [Parameter(Mandatory = $true)]
        [string]$SourceUri,
        [string]$Sha256 = "",
        [string]$Publisher = "",
        [string]$Thumbprint = "",
        [string]$Architecture = "",
        [string]$InstallPath = "",
        [ValidateSet("Installed", "VerifiedExisting", "Provisioned")]
        [string]$Status = "Installed",
        [AllowEmptyCollection()][string[]]$Evidence = @()
    )

    if ($null -eq $script:AcquisitionProvenance) {
        $script:AcquisitionProvenance = [System.Collections.ArrayList]::new()
    }
    for ($index = $script:AcquisitionProvenance.Count - 1; $index -ge 0; $index--) {
        if ([string]$script:AcquisitionProvenance[$index].Name -eq $Name) {
            $script:AcquisitionProvenance.RemoveAt($index)
        }
    }

    $record = [PSCustomObject][ordered]@{
        ManifestVersion = $script:AcquisitionManifestVersion
        Name            = $Name
        Version         = $Version
        Status          = $Status
        SourceUri       = $SourceUri
        Sha256          = $Sha256
        Publisher       = $Publisher
        Thumbprint      = $Thumbprint
        Architecture    = $Architecture
        InstallPath     = $InstallPath
        VerifiedAt      = (Get-Date).ToUniversalTime().ToString("o")
        Evidence        = @($Evidence)
    }
    [void]$script:AcquisitionProvenance.Add($record)
    return $record
}

function New-SystemUpdateProTemporaryDirectory {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseShouldProcessForStateChangingFunctions", "", Justification = "Creates a uniquely named run-owned temporary staging directory.")]
    param([string]$Purpose = "Work")

    $safePurpose = $Purpose -replace "[^A-Za-z0-9_-]", ""
    $path = Join-Path ([IO.Path]::GetTempPath()) "SystemUpdatePro_$safePurpose`_$([guid]::NewGuid().ToString('N'))"
    New-Item -ItemType Directory -Path $path -Force -ErrorAction Stop | Out-Null
    return $path
}

function Remove-SystemUpdateProTemporaryDirectory {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseShouldProcessForStateChangingFunctions", "", Justification = "Removes only a validated run-owned temporary directory.")]
    param([AllowNull()][string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) { return }
    $temporaryRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd("\", "/")
    $resolvedPath = [IO.Path]::GetFullPath($Path).TrimEnd("\", "/")
    $expectedPrefix = "$temporaryRoot$([IO.Path]::DirectorySeparatorChar)SystemUpdatePro_"
    if (-not $resolvedPath.StartsWith($expectedPrefix, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to remove an unowned temporary directory: $resolvedPath"
    }
    Remove-Item -LiteralPath $resolvedPath -Recurse -Force -ErrorAction SilentlyContinue
}

# ============================================================================
# WINGET MANAGEMENT
# ============================================================================

function Get-PolicyDocument {
    param([string]$Path = "")

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return [ordered]@{ SchemaVersion = 1 }
    }
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Policy file does not exist: $Path"
    }
    $access = Test-EvidencePathAccess -Path $Path
    if (-not $access.Valid) {
        throw "Policy file ACL is not trusted: $($access.Reason)"
    }
    $read = Read-ProtectedJsonFile -Path $Path
    if (-not $read.Success) { throw "Policy file could not be read: $($read.Error)" }
    $document = ConvertTo-Hashtable -InputObject $read.Data
    if ($document -isnot [System.Collections.IDictionary]) {
        throw "Policy file root must be a JSON object"
    }
    if ($document.Contains("schema_version") -and [int]$document.schema_version -ne 1) {
        throw "Unsupported policy schema version"
    }
    return $document
}

function Get-WingetScopePlan {
    param(
        [AllowNull()][System.Collections.IDictionary]$SystemInfo = $null,
        [bool]$WingetAvailable = $true
    )

    if ($null -eq $SystemInfo) {
        $SystemInfo = $script:CurrentSystemInfo
    }
    if ($null -eq $SystemInfo) {
        $SystemInfo = @{ ExecutionContext = "Unknown" }
    }
    $context = [string]$SystemInfo.ExecutionContext
    $scopes = [System.Collections.ArrayList]::new()
    $machineStatus = if ($WingetAvailable -and $context -in @("AdministratorUser", "System")) {
        "Available"
    } elseif (-not $WingetAvailable) { "Unavailable" } else { "Skipped" }
    [void]$scopes.Add([ordered]@{
        Scope = "machine"
        Status = $machineStatus
        CanUpgrade = ($machineStatus -eq "Available")
        User = ""
        Reason = switch ($machineStatus) {
            "Available" { "Machine scope can be serviced by the elevated engine" }
            "Unavailable" { "WinGet is not available for machine scope" }
            default { "Execution context is not approved for machine scope" }
        }
    })
    $userStatus = if (-not $WingetAvailable) { "Unavailable" } elseif ($context -eq "AdministratorUser") {
        "Available"
    } elseif ($context -eq "System") { "Skipped" } else { "Unavailable" }
    [void]$scopes.Add([ordered]@{
        Scope = "current-user"
        Status = $userStatus
        CanUpgrade = ($userStatus -eq "Available")
        User = if ($userStatus -eq "Available") { [string]$env:USERNAME } else { "" }
        Reason = switch ($userStatus) {
            "Available" { "Current administrator user's scope is visible to this process" }
            "Skipped" { "SYSTEM cannot safely claim visibility of a per-user package scope" }
            default { "No authenticated user scope is available" }
        }
    })
    [void]$scopes.Add([ordered]@{
        Scope = "other-user"
        Status = "Skipped"
        CanUpgrade = $false
        User = ""
        Reason = "Other user profiles require a separately authenticated non-elevating session helper"
    })
    return [PSCustomObject][ordered]@{
        SchemaVersion = 1
        ExecutionContext = $context
        Scopes = @($scopes)
        SystemClaimsPerUserSuccess = ($context -eq "AdministratorUser")
        Reason = if ($context -eq "System") {
            "SYSTEM results are limited to machine scope; unseen per-user packages remain explicitly skipped"
        } else {
            "Machine and current-user scopes are modeled independently"
        }
    }
}

function ConvertFrom-WingetPackageOutput {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Output
    )

    $text = (@($Output) | ForEach-Object { [string]$_ }) -join "`n"
    $jsonStart = $text.IndexOf("{")
    if ($jsonStart -lt 0) { $jsonStart = $text.IndexOf("[") }
    if ($jsonStart -ge 0) {
        try {
            $parsed = ConvertTo-Hashtable -InputObject ($text.Substring($jsonStart) | ConvertFrom-Json -ErrorAction Stop)
            $candidateItems = if ($parsed -is [System.Collections.IDictionary] -and $parsed.Contains("packages")) {
                @($parsed.packages)
            } else { @($parsed) }
            if ($candidateItems.Count -gt 0) { return @($candidateItems) }
        } catch {}
    }
    $items = [System.Collections.ArrayList]::new()
    foreach ($line in @($Output)) {
        $value = [string]$line
        if ($value -notmatch "\S" -or $value -match "^(Name|Name\s+Id|[-\\|])" -or
            $value -match "(?i)(upgrades available|no installed package)" ) { continue }
        $columns = @($value -split "\s{2,}" | Where-Object { $_ -match "\S" })
        if ($columns.Count -ge 2 -and $columns[1] -match "^[A-Za-z0-9][A-Za-z0-9_.-]*$") {
            [void]$items.Add([ordered]@{
                Name = [string]$columns[0]
                Id = [string]$columns[1]
                Version = if ($columns.Count -ge 3) { [string]$columns[2] } else { "" }
                AvailableVersion = if ($columns.Count -ge 4) { [string]$columns[3] } else { "" }
                Source = if ($columns.Count -ge 5) { [string]$columns[4] } else { "" }
            })
        }
    }
    return @($items)
}

function Get-WingetScopeInventory {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("machine", "current-user", "other-user")]
        [string]$Scope,
        [AllowNull()][System.Collections.IDictionary]$SystemInfo = $null
    )

    $plan = Get-WingetScopePlan -SystemInfo $SystemInfo
    $scopePlan = @($plan.Scopes | Where-Object { $_.Scope -eq $Scope })[0]
    if ($null -eq $scopePlan) {
        return [PSCustomObject]@{ SchemaVersion = 1; Scope = $Scope; Status = "Unavailable"; Items = @(); ExitCode = $null; Reason = "Scope is not modeled" }
    }
    if (-not $scopePlan.CanUpgrade) {
        return [PSCustomObject]@{
            SchemaVersion = 1; Scope = $Scope; Status = [string]$scopePlan.Status
            Items = @(); ExitCode = $null; Reason = [string]$scopePlan.Reason
        }
    }
    try {
        $arguments = @("list", "--scope", $(if ($Scope -eq "current-user") { "user" } else { $Scope }),
            "--accept-source-agreements", "--disable-interactivity")
        $output = @(& winget @arguments 2>&1)
        $exitCode = $LASTEXITCODE
        $items = if ($exitCode -eq 0) { ConvertFrom-WingetPackageOutput -Output $output } else { @() }
        return [PSCustomObject]@{
            SchemaVersion = 1; Scope = $Scope
            Status = if ($exitCode -eq 0) { "Ready" } else { "Unavailable" }
            Items = @($items); ExitCode = $exitCode
            Reason = if ($exitCode -eq 0) { "Scope inventory completed" } else { "WinGet scope inventory exited with $exitCode" }
        }
    } catch {
        return [PSCustomObject]@{
            SchemaVersion = 1; Scope = $Scope; Status = "Unavailable"; Items = @(); ExitCode = $null
            Reason = "WinGet scope inventory failed: $(Protect-EvidenceText -Text $_.Exception.Message)"
        }
    }
}

function Get-WingetPackagePolicy {
    param(
        [string]$Path = "",
        [string]$ExcludePath = ""
    )

    $policy = [ordered]@{
        SchemaVersion = 1
        Exclude = [System.Collections.ArrayList]::new()
        Pins = [ordered]@{}
        Conflicts = [ordered]@{}
        Sources = @()
        Errors = [System.Collections.ArrayList]::new()
    }
    try {
        $document = Get-PolicyDocument -Path $Path
        if ($document.Contains("winget") -and $document.winget -is [System.Collections.IDictionary]) {
            $document = $document.winget
        }
        foreach ($pattern in @($document.exclude)) { if (-not [string]::IsNullOrWhiteSpace([string]$pattern)) { [void]$policy.Exclude.Add([string]$pattern) } }
        if ($document.Contains("pins") -and $document.pins -is [System.Collections.IDictionary]) {
            foreach ($key in $document.pins.Keys) { $policy.Pins[[string]$key] = [string]$document.pins[$key] }
        }
        if ($document.Contains("conflicts") -and $document.conflicts -is [System.Collections.IDictionary]) {
            foreach ($key in $document.conflicts.Keys) { $policy.Conflicts[[string]$key] = ConvertTo-Hashtable -InputObject $document.conflicts[$key] }
        }
        if (-not [string]::IsNullOrWhiteSpace($Path)) { $policy.Sources += "policy:$Path" }
    } catch {
        [void]$policy.Errors.Add((Protect-EvidenceText -Text $_.Exception.Message))
    }
    if ([string]::IsNullOrWhiteSpace($ExcludePath)) {
        $ExcludePath = Join-Path $script:DataPath "winget-exclude.txt"
    }
    if (Test-Path -LiteralPath $ExcludePath -PathType Leaf) {
        try {
            foreach ($line in @(Get-Content -LiteralPath $ExcludePath -ErrorAction Stop)) {
                $pattern = ([string]$line).Trim()
                if ($pattern -and $pattern -notmatch "^#") { [void]$policy.Exclude.Add($pattern) }
            }
            $policy.Sources += "exclude-file:$ExcludePath"
        } catch {
            [void]$policy.Errors.Add((Protect-EvidenceText -Text $_.Exception.Message))
        }
    }
    $policy.Exclude = @($policy.Exclude | Select-Object -Unique)
    return [PSCustomObject]$policy
}

function Get-WingetPackagePlan {
    param(
        [AllowEmptyCollection()][object[]]$Packages = @(),
        [AllowNull()][object]$Policy = $null,
        [string[]]$RunningProcessNames = @(),
        [string]$Scope = "machine"
    )

    if ($null -eq $Policy) { $Policy = Get-WingetPackagePolicy }
    $planned = [System.Collections.ArrayList]::new()
    foreach ($package in @($Packages)) {
        $id = [string](Get-ResultValue -Result $package -Names @("Id", "id", "PackageIdentifier") -Default "")
        $name = [string](Get-ResultValue -Result $package -Names @("Name", "name") -Default $id)
        $current = [string](Get-ResultValue -Result $package -Names @("Version", "version", "InstalledVersion") -Default "")
        $available = [string](Get-ResultValue -Result $package -Names @("AvailableVersion", "available_version", "LatestVersion") -Default "")
        $source = [string](Get-ResultValue -Result $package -Names @("Source", "source") -Default "")
        $item = [ordered]@{
            Scope = $Scope; Id = $id; Name = $name; CurrentVersion = $current
            AvailableVersion = $available; Source = $source; Status = "Eligible"
            Reason = "Package is eligible for unattended upgrade"; Conflict = $false
            Deadline = ""; PinnedVersion = ""
        }
        $excluded = @($Policy.Exclude | Where-Object { $id -like [string]$_ -or $name -like [string]$_ }).Count -gt 0
        if ($excluded) {
            $item.Status = "Excluded"; $item.Reason = "Matched a wildcard package exclusion"
        } elseif ($Policy.Pins -is [System.Collections.IDictionary] -and $Policy.Pins.Contains($id)) {
            $pin = [string]$Policy.Pins[$id]
            $item.PinnedVersion = $pin
            $currentVersion = ConvertTo-SafeVersion -Value $current
            $pinVersion = ConvertTo-SafeVersion -Value $pin
            if ($null -eq $currentVersion -or $null -eq $pinVersion -or $currentVersion -lt $pinVersion) {
                $item.Status = "Pinned"; $item.Reason = "Upgrade would cross the administrator's maximum pinned version"
            } elseif ($null -ne (ConvertTo-SafeVersion -Value $available) -and
                (ConvertTo-SafeVersion -Value $available) -gt $pinVersion) {
                $item.Status = "Pinned"; $item.Reason = "Available version exceeds the administrator's maximum pinned version"
            }
        }
        if ($item.Status -eq "Eligible" -and $Policy.Conflicts -is [System.Collections.IDictionary]) {
            $conflictRule = $null
            if ($Policy.Conflicts.Contains($id)) { $conflictRule = $Policy.Conflicts[$id] }
            elseif ($Policy.Conflicts.Contains($name)) { $conflictRule = $Policy.Conflicts[$name] }
            if ($null -ne $conflictRule) {
                $processes = @()
                if ($conflictRule -is [System.Collections.IDictionary]) {
                    if ($conflictRule.Contains("processes")) { $processes += @($conflictRule.processes) }
                    if ($conflictRule.Contains("Processes")) { $processes += @($conflictRule.Processes) }
                }
                $matched = @($processes | Where-Object { $RunningProcessNames -contains [string]$_ })
                if ($matched.Count -gt 0) {
                    $item.Conflict = $true
                    $action = [string]$conflictRule.action
                    if ([string]::IsNullOrWhiteSpace($action)) { $action = [string]$conflictRule.Action }
                    if ($action -eq "close-with-deadline") {
                        $minutes = 15
                        [void][int]::TryParse([string]$conflictRule.deadline_minutes, [ref]$minutes)
                        $item.Status = "Deferred"
                        $item.Deadline = (Get-Date).AddMinutes([math]::Max(1, $minutes)).ToUniversalTime().ToString("o")
                        $item.Reason = "Conflict requires a user-session deadline; unattended execution will not force-close the process"
                    } elseif ($action -eq "skip") {
                        $item.Status = "Skipped"; $item.Reason = "Configured conflicting process is running"
                    } else {
                        $item.Status = "Deferred"; $item.Reason = "Configured conflicting process is running; package was deferred"
                    }
                }
            }
        }
        [void]$planned.Add([PSCustomObject]$item)
    }
    return [PSCustomObject][ordered]@{
        SchemaVersion = 1; Scope = $Scope; EvaluatedAt = (Get-Date).ToUniversalTime().ToString("o")
        Items = @($planned)
        Eligible = @($planned | Where-Object Status -eq "Eligible")
        Skipped = @($planned | Where-Object Status -in @("Excluded", "Skipped", "Pinned"))
        Deferred = @($planned | Where-Object Status -eq "Deferred")
        Conflicts = @($planned | Where-Object Conflict)
    }
}

function Get-EndpointCohort {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DeviceIdentity,
        [string]$Cohort = "default"
    )

    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [Text.Encoding]::UTF8.GetBytes("$Cohort|$DeviceIdentity")
        $digest = $sha.ComputeHash($bytes)
        $value = [BitConverter]::ToUInt32($digest, 0)
        return [int]($value % 100)
    } finally { $sha.Dispose() }
}

function Get-RolloutPolicy {
    param([string]$Path = "")

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return [PSCustomObject][ordered]@{
            SchemaVersion = 1; Enabled = $false; Cohort = "default"; PercentageStart = 0; PercentageEnd = 100
            MinimumSuccessCount = 0; MinimumSuccessRate = 1.0; BakeTimeHours = 0; Deadline = ""
            EmergencyOverride = $false; Approved = @{}; Reason = "No rollout policy configured"
        }
    }
    try {
        $document = Get-PolicyDocument -Path $Path
        $rollout = if ($document.Contains("rollout")) { $document.rollout } else { $document }
        $start = 0; $end = 100; $minimumCount = 1; $minimumRate = 1.0; $bakeHours = 0
        [void][int]::TryParse([string]$rollout.percentage_start, [ref]$start)
        [void][int]::TryParse([string]$rollout.percentage_end, [ref]$end)
        [void][int]::TryParse([string]$rollout.minimum_success_count, [ref]$minimumCount)
        [void][double]::TryParse([string]$rollout.minimum_success_rate, [ref]$minimumRate)
        [void][int]::TryParse([string]$rollout.bake_time_hours, [ref]$bakeHours)
        return [PSCustomObject][ordered]@{
            SchemaVersion = 1; Enabled = $true
            Cohort = if ($rollout.cohort) { [string]$rollout.cohort } else { "default" }
            PercentageStart = [math]::Max(0, [math]::Min(99, $start))
            PercentageEnd = [math]::Max(1, [math]::Min(100, $end))
            MinimumSuccessCount = [math]::Max(0, $minimumCount)
            MinimumSuccessRate = [math]::Max(0.0, [math]::Min(1.0, $minimumRate))
            BakeTimeHours = [math]::Max(0, $bakeHours)
            Deadline = [string]$rollout.deadline
            EmergencyOverride = [bool]$rollout.emergency_override
            Approved = if ($rollout.approved) { ConvertTo-Hashtable -InputObject $rollout.approved } else { @{} }
            Reason = "Rollout policy loaded from protected local evidence"
        }
    } catch {
        return [PSCustomObject][ordered]@{
            SchemaVersion = 1; Enabled = $false; Cohort = "default"; PercentageStart = 0; PercentageEnd = 0
            MinimumSuccessCount = 0; MinimumSuccessRate = 1.0; BakeTimeHours = 0; Deadline = ""
            EmergencyOverride = $false; Approved = @{}; Reason = Protect-EvidenceText -Text $_.Exception.Message
            Error = $true
        }
    }
}

function Get-RolloutEvidence {
    param([AllowEmptyCollection()][object[]]$HistoryEntries = @())

    $successEntries = @($HistoryEntries | Where-Object {
        [string]$_.status -in @("Succeeded", "SucceededRebootRequired")
    })
    $failureEntries = @($HistoryEntries | Where-Object {
        [string]$_.status -in @("Partial", "Failed")
    })
    $lastSuccess = @($successEntries | Sort-Object timestamp -Descending | Select-Object -First 1)
    return [PSCustomObject][ordered]@{
        SchemaVersion = 1
        SuccessCount = $successEntries.Count
        FailureCount = $failureEntries.Count
        LastSuccessAt = if ($lastSuccess.Count -gt 0) { [string]$lastSuccess[0].timestamp } else { "" }
        SampleCount = $successEntries.Count + $failureEntries.Count
        Source = "local update history"
    }
}

function Evaluate-RolloutPromotion {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Policy,
        [Parameter(Mandatory = $true)]
        [object]$Evidence,
        [int]$CohortValue = 0,
        [bool]$EmergencyOverride = $false
    )

    $reasons = [System.Collections.ArrayList]::new()
    if (-not [bool]$Policy.Enabled) {
        return [PSCustomObject][ordered]@{
            SchemaVersion = 1; Decision = "NotConfigured"; CohortValue = $CohortValue
            SuccessCount = 0; FailureCount = 0; SuccessRate = 0.0; Reasons = @([string]$Policy.Reason)
            AuditedOverride = $false; EvaluatedAt = (Get-Date).ToUniversalTime().ToString("o")
        }
    }
    if ($CohortValue -lt [int]$Policy.PercentageStart -or $CohortValue -ge [int]$Policy.PercentageEnd) {
        [void]$reasons.Add("Endpoint cohort value $CohortValue is outside the assigned rollout range")
        return [PSCustomObject][ordered]@{
            SchemaVersion = 1; Decision = "Hold"; CohortValue = $CohortValue
            SuccessCount = [int]$Evidence.SuccessCount; FailureCount = [int]$Evidence.FailureCount
            SuccessRate = 0.0; Reasons = @($reasons); AuditedOverride = $false
            EvaluatedAt = (Get-Date).ToUniversalTime().ToString("o")
        }
    }
    if ($EmergencyOverride -or [bool]$Policy.EmergencyOverride) {
        [void]$reasons.Add("Emergency override was explicitly recorded by the administrator")
        return [PSCustomObject][ordered]@{
            SchemaVersion = 1; Decision = "Promote"; CohortValue = $CohortValue
            SuccessCount = [int]$Evidence.SuccessCount; FailureCount = [int]$Evidence.FailureCount
            SuccessRate = 1.0; Reasons = @($reasons); AuditedOverride = $true
            EvaluatedAt = (Get-Date).ToUniversalTime().ToString("o")
        }
    }
    $successCount = [int]$Evidence.SuccessCount
    $failureCount = [int]$Evidence.FailureCount
    $total = $successCount + $failureCount
    $rate = if ($total -gt 0) { [double]$successCount / $total } else { 0.0 }
    if ($failureCount -gt 0 -and $rate -lt [double]$Policy.MinimumSuccessRate) {
        [void]$reasons.Add("Observed success rate $([math]::Round($rate * 100, 2))% is below the required $([math]::Round([double]$Policy.MinimumSuccessRate * 100, 2))%")
        $decision = "Halt"
    } elseif ($successCount -lt [int]$Policy.MinimumSuccessCount) {
        [void]$reasons.Add("Waiting for at least $($Policy.MinimumSuccessCount) successful endpoint result(s)")
        $decision = "Hold"
    } else {
        $lastSuccess = [datetime]::MinValue
        [void][datetime]::TryParse([string]$Evidence.LastSuccessAt, [ref]$lastSuccess)
        if ($lastSuccess -eq [datetime]::MinValue -or
            ((Get-Date).ToUniversalTime() - $lastSuccess.ToUniversalTime()).TotalHours -lt [double]$Policy.BakeTimeHours) {
            [void]$reasons.Add("Bake time has not elapsed")
            $decision = "Hold"
        } else {
            [void]$reasons.Add("Success count, rate, and bake-time requirements passed")
            $decision = "Promote"
        }
    }
    return [PSCustomObject][ordered]@{
        SchemaVersion = 1; Decision = $decision; CohortValue = $CohortValue
        SuccessCount = $successCount; FailureCount = $failureCount; SuccessRate = $rate
        Reasons = @($reasons); AuditedOverride = $false
        EvaluatedAt = (Get-Date).ToUniversalTime().ToString("o")
    }
}

function Get-WingetTrustEvidence {
    $spec = Get-AcquisitionManifestEntry -Name "WinGet"

    try {
        $null = Get-Command winget -ErrorAction Stop
        $versionOutput = & winget --version 2>&1
        if ($LASTEXITCODE -ne 0) {
            throw "winget --version exited with $LASTEXITCODE"
        }
        $version = ConvertTo-SafeVersion -Value (($versionOutput | Out-String).Trim())
        if (-not (Test-VersionAtLeast -Version $version -MinimumVersion $spec.MinimumVersion)) {
            throw "WinGet $version is below the approved minimum $($spec.MinimumVersion)"
        }

        $package = $null
        try {
            $package = Get-AppxPackage -Name "Microsoft.DesktopAppInstaller" -AllUsers -ErrorAction Stop |
                Sort-Object Version -Descending | Select-Object -First 1
        } catch {
            $package = $null
        }
        if (-not $package) {
            $package = Get-AppxPackage -Name "Microsoft.DesktopAppInstaller" -ErrorAction Stop |
                Sort-Object Version -Descending | Select-Object -First 1
        }
        if (-not $package) { throw "The Microsoft Desktop App Installer package was not found" }
        if ([string]$package.Publisher -notmatch $spec.PublisherPattern) {
            throw "App Installer publisher '$($package.Publisher)' is not approved"
        }

        return [PSCustomObject]@{
            Valid        = $true
            Version      = [string]$version
            PackageVersion = [string]$package.Version
            Publisher    = [string]$package.Publisher
            Architecture = ([string]$package.Architecture).ToLowerInvariant()
            InstallPath  = [string]$package.InstallLocation
            Reason       = ""
        }
    } catch {
        return [PSCustomObject]@{
            Valid = $false; Version = ""; PackageVersion = ""; Publisher = ""
            Architecture = ""; InstallPath = ""; Reason = $_.Exception.Message
        }
    }
}

function Test-WingetInstalled {
    $evidence = Get-WingetTrustEvidence
    if (-not $evidence.Valid) { return $false }

    $spec = Get-AcquisitionManifestEntry -Name "WinGet"
    [void](Add-AcquisitionProvenance -Name $spec.Name -Version $evidence.Version `
        -SourceUri $spec.Uri -Publisher $evidence.Publisher -Architecture $evidence.Architecture `
        -InstallPath $evidence.InstallPath -Status "VerifiedExisting" `
        -Evidence @("App Installer package $($evidence.PackageVersion)"))
    return $true
}

function Install-Winget {
    $spec = Get-AcquisitionManifestEntry -Name "WinGet"
    if (Test-WingetInstalled) {
        Write-Log "Verified WinGet v$($spec.MinimumVersion) or later from Microsoft App Installer" "DEBUG"
        return $true
    }

    Write-Log "Installing verified WinGet $($spec.ExactVersion)..." "STEP"

    if ($DryRun) {
        Write-Log "Would install the pinned WinGet bundle and signed architecture-specific dependencies" "INFO"
        return $true
    }

    $architecture = Get-SystemArchitecture
    if ($architecture -notin @($spec.Architectures)) {
        Write-Log "WinGet acquisition does not support architecture '$architecture'" "WARNING"
        return $false
    }

    # Re-register a trusted package that is present but whose execution alias
    # has not been registered for the current account.
    try {
        Add-AppxPackage -RegisterByFamilyName -MainPackage Microsoft.DesktopAppInstaller_8wekyb3d8bbwe -ErrorAction Stop
        Start-Sleep -Seconds 2
        if (Test-WingetInstalled) {
            Write-Log "WinGet restored from the verified App Installer package" "SUCCESS"
            return $true
        }
    } catch {}

    $tempDir = New-SystemUpdateProTemporaryDirectory -Purpose "WinGet"

    try {
        $bundlePath = Join-Path $tempDir "winget.msixbundle"
        $dependenciesZip = Join-Path $tempDir "dependencies.zip"
        $licensePath = Join-Path $tempDir "License.xml"
        $dependenciesRoot = Join-Path $tempDir "dependencies"

        $bundleEvidence = Invoke-VerifiedDownload -Name "WinGet bundle" -Uri $spec.Uri `
            -AllowedHosts $spec.AllowedHosts -ExpectedSha256 $spec.Sha256 `
            -PublisherPattern $spec.PublisherPattern -DestinationPath $bundlePath
        $dependencyEvidence = Invoke-VerifiedDownload -Name "WinGet dependency bundle" `
            -Uri $spec.DependencyBundle.Uri -AllowedHosts $spec.DependencyBundle.AllowedHosts `
            -ExpectedSha256 $spec.DependencyBundle.Sha256 -DestinationPath $dependenciesZip
        $null = Invoke-VerifiedDownload -Name "WinGet license" -Uri $spec.License.Uri `
            -AllowedHosts $spec.License.AllowedHosts -ExpectedSha256 $spec.License.Sha256 `
            -DestinationPath $licensePath

        Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
        [IO.Compression.ZipFile]::ExtractToDirectory($dependenciesZip, $dependenciesRoot)
        $dependencyPaths = @()
        foreach ($packageTemplate in @($spec.DependencyBundle.Packages)) {
            $packageName = $packageTemplate -f $architecture
            $packagePath = Join-Path (Join-Path $dependenciesRoot $architecture) $packageName
            if (-not (Test-Path -LiteralPath $packagePath -PathType Leaf)) {
                throw "WinGet dependency '$packageName' is missing from the verified bundle"
            }
            $signature = Get-AuthenticodeEvidence -Path $packagePath -PublisherPattern $spec.PublisherPattern
            if (-not $signature.Valid) {
                throw "WinGet dependency '$packageName' was rejected before installation: $($signature.Reason)"
            }
            $dependencyPaths += $packagePath
        }

        try {
            Add-AppxProvisionedPackage -Online -PackagePath $bundlePath -LicensePath $licensePath `
                -DependencyPackagePath $dependencyPaths -ErrorAction Stop | Out-Null
        } catch {
            Write-Log "Machine provisioning failed; installing the same verified bundle for the current account" "DEBUG"
            Add-AppxPackage -Path $bundlePath -DependencyPath $dependencyPaths -ErrorAction Stop
        }

        Start-Sleep -Seconds 2

        if (Test-WingetInstalled) {
            $installedEvidence = Get-WingetTrustEvidence
            [void](Add-AcquisitionProvenance -Name $spec.Name -Version $installedEvidence.Version `
                -SourceUri $spec.Uri -Sha256 $bundleEvidence.Sha256 -Publisher $bundleEvidence.Publisher `
                -Thumbprint $bundleEvidence.Thumbprint -Architecture $architecture `
                -InstallPath $installedEvidence.InstallPath -Status "Provisioned" `
                -Evidence @(
                    "Dependency bundle SHA-256 $($dependencyEvidence.Sha256)",
                    "Dependencies $($spec.DependencyBundle.Versions -join ', ')"
                ))
            Write-Log "WinGet $($installedEvidence.Version) installed from verified Microsoft artifacts" "SUCCESS"
            return $true
        }
        throw "WinGet did not pass version and publisher verification after installation"
    } catch {
        Write-Log "Winget installation error: $($_.Exception.Message)" "WARNING"
    } finally {
        Remove-SystemUpdateProTemporaryDirectory -Path $tempDir
    }

    Write-Log "Failed to install Winget" "WARNING"
    return $false
}

function Invoke-WingetUpgradeAll {
    $result = @{
        Success = $false; RebootRequired = $false
        UpdateCount = 0; Available = 0; Attempted = 0; Installed = 0; Failed = 0; Skipped = 0
        ExitCode = $null; HResult = $null; Items = @(); Evidence = @(); Message = ""
        ScopeResults = @(); Plans = [System.Collections.ArrayList]::new(); Policy = $null
    }

    Write-Log "========== WINGET UPGRADE ALL ==========" "HEADER"

    $policy = Get-WingetPackagePolicy -Path ([string]$script:PolicyPath)
    $result.Policy = $policy

    if (-not (Test-WingetInstalled)) {
        if (-not (Install-Winget)) {
            $result.Message = "Winget not available"
            $result.Failed = 1
            Write-Log $result.Message "WARNING"
            return $result
        }
    }

    $scopePlan = Get-WingetScopePlan -SystemInfo $script:CurrentSystemInfo -WingetAvailable $true
    $result.ScopeResults = @($scopePlan.Scopes)
    $script:WingetScopeResults = @($scopePlan.Scopes)
    if (-not (Test-DownloadAllowed)) {
        $result.Success = $true
        $result.Skipped = 1
        $result.Message = "WinGet deferred by network download policy: $($script:DownloadPolicy.Reason)"
        $result.Evidence = @("download-policy:$($script:DownloadPolicy.Status)")
        Write-Log $result.Message "WARNING"
        return $result
    }

    try {
        # Source refresh mutates WinGet state, so never perform it during a dry run.
        if (-not $DryRun) {
            & winget source update --disable-interactivity 2>&1 | Out-Null
        }

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

        $hasPackagePolicy = @($policy.Exclude).Count -gt 0 -or
            ($policy.Pins -is [System.Collections.IDictionary] -and $policy.Pins.Count -gt 0) -or
            ($policy.Conflicts -is [System.Collections.IDictionary] -and $policy.Conflicts.Count -gt 0)
        if ($hasPackagePolicy) {
            $runningProcesses = @(Get-Process -ErrorAction SilentlyContinue | ForEach-Object { [string]$_.ProcessName })
            $policyItems = [System.Collections.ArrayList]::new()
            foreach ($scope in @($scopePlan.Scopes | Where-Object CanUpgrade)) {
                $inventory = Get-WingetScopeInventory -Scope ([string]$scope.Scope) -SystemInfo $script:CurrentSystemInfo
                $plan = Get-WingetPackagePlan -Packages @($inventory.Items) -Policy $policy `
                    -RunningProcessNames $runningProcesses -Scope ([string]$scope.Scope)
                [void]$result.Plans.Add($plan)
                foreach ($planItem in @($plan.Items)) {
                    [void]$policyItems.Add((New-UpdateItemResult -Name ([string]$planItem.Name) `
                        -Id ([string]$planItem.Id) `
                        -Status $(if ($planItem.Status -eq "Eligible") { "Attempted" } elseif ($planItem.Status -eq "Deferred") { "Skipped" } else { "Skipped" }) `
                        -Message ([string]$planItem.Reason) `
                        -Evidence @("scope:$($planItem.Scope)", "policy-status:$($planItem.Status)")))
                }
                foreach ($eligible in @($plan.Eligible)) {
                    $arguments = @("upgrade", "--id", [string]$eligible.Id, "--silent",
                        "--accept-package-agreements", "--accept-source-agreements", "--disable-interactivity",
                        "--scope", $(if ($scope.Scope -eq "current-user") { "user" } else { "machine" }))
                    $process = Start-Process -FilePath "winget" -ArgumentList $arguments -Wait -NoNewWindow -PassThru
                    $result.Attempted++
                    if ($process.ExitCode -in @(0, -1978335189)) {
                        $result.Installed++
                        $result.UpdateCount++
                    } else {
                        $result.Failed++
                    }
                }
            }
            $result.Items = @($policyItems)
            $result.Skipped = @($policyItems | Where-Object { $_.Skipped }).Count
            $result.Success = ($result.Failed -eq 0)
            $result.Message = "WinGet policy plan completed: $($result.Installed) installed, $($result.Skipped) skipped/deferred"
            $script:WingetUpdateCount = $result.Installed
            return $result
        }

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

function Get-PackageManagerAvailability {
    $managers = [System.Collections.ArrayList]::new()
    foreach ($name in @("choco", "scoop")) {
        $command = Get-Command $name -ErrorAction SilentlyContinue | Select-Object -First 1
        [void]$managers.Add([PSCustomObject][ordered]@{
            Name = switch ($name) { "choco" { "Chocolatey" }; "scoop" { "Scoop" } }
            CommandName = $name
            Path = if ($command) { [string]$command.Source } else { "" }
            Version = if ($command) { [string]$command.Version } else { "" }
            Status = if ($command) { "Ready" } else { "Unavailable" }
            Reason = if ($command) { "Package manager command is installed" } else { "$name is not installed" }
        })
    }
    return @($managers)
}

function Get-WingetStoreSourceAvailability {
    if (-not (Test-WingetInstalled)) {
        return [PSCustomObject]@{ Available = $false; Source = ""; Status = "Unavailable"; Reason = "WinGet is not installed" }
    }
    try {
        $output = @(& winget source list --disable-interactivity 2>&1)
        $exitCode = $LASTEXITCODE
        $storeLine = @($output | Where-Object { [string]$_ -match "(?i)StoreEdgeFD|msstore|Microsoft Store" }) | Select-Object -First 1
        if ($exitCode -eq 0 -and $storeLine) {
            return [PSCustomObject]@{ Available = $true; Source = "msstore"; Status = "Ready"; Reason = "StoreEdgeFD-backed Microsoft Store source is registered" }
        }
        return [PSCustomObject]@{ Available = $false; Source = "msstore"; Status = "Unavailable"; Reason = "StoreEdgeFD/msstore source is not registered" }
    } catch {
        return [PSCustomObject]@{ Available = $false; Source = "msstore"; Status = "Unavailable"; Reason = "WinGet source inventory failed: $($_.Exception.Message)" }
    }
}

function Invoke-WingetStoreUpgrade {
    $result = @{
        Success = $true; Status = "Skipped"; RebootRequired = $false
        Available = 0; Attempted = 0; Installed = 0; Failed = 0; Skipped = 0
        ExitCode = $null; HResult = $null; Items = @(); Evidence = @(); Message = ""
    }
    $source = Get-WingetStoreSourceAvailability
    if (-not $source.Available) {
        $result.Skipped = 1
        $result.Items += New-UpdateItemResult -Name "Microsoft Store (StoreEdgeFD)" -Status "Skipped" -Message $source.Reason
        $result.Message = $source.Reason
        return $result
    }
    if (-not (Test-DownloadAllowed)) {
        $result.Skipped = 1
        $result.Items += New-UpdateItemResult -Name "Microsoft Store (StoreEdgeFD)" -Status "Skipped" `
            -Message "Store upgrades deferred by the provider download policy"
        $result.Message = "StoreEdgeFD upgrade deferred by network policy"
        return $result
    }
    try {
        $arguments = @("upgrade", "--all", "--source", [string]$source.Source, "--silent",
            "--accept-package-agreements", "--accept-source-agreements", "--disable-interactivity")
        if ($DryRun) {
            $output = @(& winget @arguments 2>&1)
            $result.ExitCode = $LASTEXITCODE
            $result.Available = @($output | Where-Object { [string]$_ -match "\S" -and [string]$_ -notmatch "(?i)no applicable|source|Name|^-+" }).Count
            $result.Success = ($result.ExitCode -in @(0, -1978335189))
            if (-not $result.Success) { $result.Failed = 1 }
            $result.Items += New-UpdateItemResult -Name "Microsoft Store (StoreEdgeFD)" `
                -Status $(if ($result.Success) { "Available" } else { "Failed" }) `
                -ProviderCode $result.ExitCode -Message "StoreEdgeFD source queried without installation"
            $result.Message = "$($result.Available) StoreEdgeFD upgrade result line(s) available (dry run)"
            return $result
        }
        $process = Start-Process -FilePath "winget" -ArgumentList $arguments `
            -Wait -NoNewWindow -PassThru -ErrorAction Stop
        $result.Attempted = 1
        $result.ExitCode = $process.ExitCode
        $result.Success = ($process.ExitCode -in @(0, -1978335189))
        if ($result.Success) { $result.Installed = 1; $result.UpdateCount = 1 } else { $result.Failed = 1 }
        $result.Status = if ($result.Success) { "Succeeded" } else { "Failed" }
        $result.Items += New-UpdateItemResult -Name "Microsoft Store (StoreEdgeFD)" `
            -Status $(if ($result.Success) { "Installed" } else { "Failed" }) -ProviderCode $result.ExitCode
        $result.Message = "StoreEdgeFD upgrade completed with exit $($result.ExitCode)"
    } catch {
        $result.Success = $false; $result.Status = "Failed"; $result.Failed = 1
        $result.Message = "StoreEdgeFD upgrade failed: $($_.Exception.Message)"
        $result.Items += New-UpdateItemResult -Name "Microsoft Store (StoreEdgeFD)" -Status "Failed" -Message $result.Message
    }
    return $result
}

function Get-WSLPackageManagerPlan {
    $inWsl = -not [string]::IsNullOrWhiteSpace([string]$env:WSL_DISTRO_NAME)
    $gui = $inWsl -and (-not [string]::IsNullOrWhiteSpace([string]$env:DISPLAY) -or
        -not [string]::IsNullOrWhiteSpace([string]$env:WAYLAND_DISPLAY))
    $plans = [System.Collections.ArrayList]::new()
    foreach ($name in @("flatpak", "snap")) {
        $command = if ($gui) { Get-Command $name -ErrorAction SilentlyContinue | Select-Object -First 1 } else { $null }
        [void]$plans.Add([PSCustomObject][ordered]@{
            Name = $name; Path = if ($command) { [string]$command.Source } else { "" }
            Status = if (-not $inWsl) { "Skipped" } elseif (-not $gui) { "Skipped" } elseif ($command) { "Ready" } else { "Unavailable" }
            Reason = if (-not $inWsl) { "Flatpak and Snap are limited to WSL GUI environments" } elseif (-not $gui) { "WSL GUI display variables are not present" } elseif ($command) { "$name is available in the WSL GUI environment" } else { "$name is not installed in the WSL GUI environment" }
        })
    }
    return @($plans)
}

function Invoke-WSLPackageManagerUpdates {
    $result = @{
        Success = $true; Status = "Skipped"; RebootRequired = $false
        Available = 0; Attempted = 0; Installed = 0; Failed = 0; Skipped = 0
        ExitCode = $null; HResult = $null; Items = @(); Evidence = @(); Message = ""
    }
    foreach ($plan in @(Get-WSLPackageManagerPlan)) {
        if ($plan.Status -ne "Ready") {
            $result.Skipped++
            $result.Items += New-UpdateItemResult -Name "WSL $($plan.Name)" -Status "Skipped" -Message $plan.Reason
            continue
        }
        try {
            $arguments = if ($plan.Name -eq "flatpak") { @("update", "-y") } else { @("refresh") }
            if ($DryRun) {
                $probeArguments = if ($plan.Name -eq "flatpak") { @("update", "--assumeno") } else { @("refresh", "--list") }
                $output = @(& ([string]$plan.Path) @probeArguments 2>&1)
                $result.Available += @($output | Where-Object { [string]$_ -match "\S" }).Count
                $result.Items += New-UpdateItemResult -Name "WSL $($plan.Name)" -Status "Available" -Message "WSL GUI package source queried"
                continue
            }
            $output = @(& ([string]$plan.Path) @arguments 2>&1)
            $exitCode = $LASTEXITCODE
            $result.Attempted++
            $result.ExitCode = $exitCode
            if ($exitCode -eq 0) {
                $result.Installed++
                $result.Items += New-UpdateItemResult -Name "WSL $($plan.Name)" -Status "Installed" -ProviderCode $exitCode
            } else {
                $result.Failed++
                $result.Items += New-UpdateItemResult -Name "WSL $($plan.Name)" -Status "Failed" -ProviderCode $exitCode `
                    -Message ((@($output) | Select-Object -Last 1) -join "")
            }
        } catch {
            $result.Failed++
            $result.Items += New-UpdateItemResult -Name "WSL $($plan.Name)" -Status "Failed" -Message $_.Exception.Message
        }
    }
    $result.Success = ($result.Failed -eq 0)
    $result.Status = if ($result.Failed -gt 0 -and $result.Installed -gt 0) { "Partial" } elseif ($result.Failed -gt 0) { "Failed" } elseif ($result.Installed -gt 0) { "Succeeded" } else { "Skipped" }
    $result.Message = "WSL package sources: installed $($result.Installed), failed $($result.Failed), skipped $($result.Skipped)"
    return $result
}

function Invoke-PackageManagerUpgrades {
    $result = @{
        Success = $true; Status = "Succeeded"; RebootRequired = $false
        Available = 0; Attempted = 0; Installed = 0; Failed = 0; Skipped = 0
        ExitCode = $null; HResult = $null; Items = @(); Evidence = @(); Message = ""
        Managers = @()
    }
    $result.Managers = @(Get-PackageManagerAvailability)
    $storeResult = Invoke-WingetStoreUpgrade
    $wslResult = Invoke-WSLPackageManagerUpdates
    $result.Items += @($storeResult.Items) + @($wslResult.Items)
    foreach ($subResult in @($storeResult, $wslResult)) {
        $result.Available += [int]$subResult.Available
        $result.Attempted += [int]$subResult.Attempted
        $result.Installed += [int]$subResult.Installed
        $result.Failed += [int]$subResult.Failed
        $result.Skipped += [int]$subResult.Skipped
    }

    foreach ($manager in @($result.Managers)) {
        if ($manager.Status -ne "Ready") {
            $result.Skipped++
            $result.Items += New-UpdateItemResult -Name ([string]$manager.Name) -Status "Skipped" -Message ([string]$manager.Reason)
            continue
        }
        try {
            if ([string]$manager.Name -eq "Chocolatey") {
                $queryOutput = @(& ([string]$manager.Path) outdated --limit-output --no-color 2>&1)
                if ($DryRun) {
                    $packages = @($queryOutput | Where-Object { [string]$_ -match "\|" })
                    $result.Available += $packages.Count
                    foreach ($package in $packages) {
                        $result.Items += New-UpdateItemResult -Name ([string]$package) -Status "Available" -Message "Chocolatey outdated package"
                    }
                } else {
                    $output = @(& ([string]$manager.Path) upgrade all -y --no-progress --limit-output 2>&1)
                    $exitCode = $LASTEXITCODE
                    $result.Attempted++
                    $result.ExitCode = $exitCode
                    if ($exitCode -eq 0) { $result.Installed++; $result.Items += New-UpdateItemResult -Name "Chocolatey packages" -Status "Installed" -ProviderCode $exitCode }
                    else { $result.Failed++; $result.Items += New-UpdateItemResult -Name "Chocolatey packages" -Status "Failed" -ProviderCode $exitCode -Message ((@($output) | Select-Object -Last 1) -join "") }
                }
            } else {
                $queryOutput = @(& ([string]$manager.Path) status 2>&1)
                if ($DryRun) {
                    $packages = @($queryOutput | Where-Object { [string]$_ -match "\S" -and [string]$_ -notmatch "(?i)Everything is up to date|Name" })
                    $result.Available += $packages.Count
                    foreach ($package in $packages) { $result.Items += New-UpdateItemResult -Name ([string]$package) -Status "Available" -Message "Scoop outdated package" }
                } else {
                    $output = @(& ([string]$manager.Path) update * 2>&1)
                    $exitCode = $LASTEXITCODE
                    $result.Attempted++
                    $result.ExitCode = $exitCode
                    if ($exitCode -eq 0) { $result.Installed++; $result.Items += New-UpdateItemResult -Name "Scoop packages" -Status "Installed" -ProviderCode $exitCode }
                    else { $result.Failed++; $result.Items += New-UpdateItemResult -Name "Scoop packages" -Status "Failed" -ProviderCode $exitCode -Message ((@($output) | Select-Object -Last 1) -join "") }
                }
            }
        } catch {
            $result.Failed++
            $result.Items += New-UpdateItemResult -Name ([string]$manager.Name) -Status "Failed" -Message $_.Exception.Message
        }
    }
    $result.Success = ($result.Failed -eq 0)
    $result.Status = if ($result.Failed -gt 0 -and $result.Installed -gt 0) { "Partial" } elseif ($result.Failed -gt 0) { "Failed" } elseif ($result.Installed -gt 0) { "Succeeded" } else { "Succeeded" }
    $result.Message = "Package managers: installed $($result.Installed), available $($result.Available), failed $($result.Failed), skipped $($result.Skipped)"
    return $result
}

# ============================================================================
# POWERSHELL MODULE MANAGEMENT
# ============================================================================

function Get-DirectoryPayloadHash {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $resolvedRoot = (Resolve-Path -LiteralPath $Path -ErrorAction Stop).Path.TrimEnd("\", "/")
    $lines = [System.Collections.Generic.List[string]]::new()
    foreach ($file in @(Get-ChildItem -LiteralPath $resolvedRoot -File -Recurse -Force -ErrorAction Stop |
        Sort-Object FullName)) {
        $relativePath = $file.FullName.Substring($resolvedRoot.Length).TrimStart("\", "/").Replace("\", "/")
        $fileHash = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256 -ErrorAction Stop).Hash.ToUpperInvariant()
        $lines.Add("$relativePath|$fileHash")
    }

    $payload = [Text.Encoding]::UTF8.GetBytes(($lines -join "`n"))
    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        return [BitConverter]::ToString($sha.ComputeHash($payload)).Replace("-", "")
    } finally {
        $sha.Dispose()
    }
}

function Get-PSModuleInstallPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ModuleName,
        [Parameter(Mandatory = $true)]
        [string]$Version
    )

    if ([string]::IsNullOrWhiteSpace([string]$script:PSModuleInstallRoot)) {
        $script:PSModuleInstallRoot = Join-Path ([Environment]::GetFolderPath("ProgramFiles")) "WindowsPowerShell\Modules"
    }
    return Join-Path (Join-Path $script:PSModuleInstallRoot $ModuleName) $Version
}

function Test-InstalledPSModule {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ModuleName
    )

    try {
        $spec = Get-AcquisitionManifestEntry -Name $ModuleName
        if ($spec.Kind -ne "PowerShellModule") { throw "'$ModuleName' is not an approved PowerShell module" }
        $installPath = Get-PSModuleInstallPath -ModuleName $ModuleName -Version $spec.ExactVersion
        $manifestPath = Join-Path $installPath "$ModuleName.psd1"
        if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) {
            throw "The pinned module manifest is missing"
        }

        $moduleManifest = Import-PowerShellDataFile -LiteralPath $manifestPath -ErrorAction Stop
        if ([string]$moduleManifest.ModuleVersion -ne [string]$spec.ExactVersion) {
            throw "Installed module version '$($moduleManifest.ModuleVersion)' is not the pinned version $($spec.ExactVersion)"
        }
        $payloadHash = Get-DirectoryPayloadHash -Path $installPath
        if ($payloadHash -ne [string]$spec.PayloadTreeSha256) {
            throw "Installed module payload hash does not match the verified package"
        }

        return [PSCustomObject]@{
            Valid = $true; Version = [string]$moduleManifest.ModuleVersion
            Path = $manifestPath; InstallPath = $installPath; Sha256 = $payloadHash; Reason = ""
        }
    } catch {
        return [PSCustomObject]@{
            Valid = $false; Version = ""; Path = ""; InstallPath = ""; Sha256 = ""
            Reason = $_.Exception.Message
        }
    }
}

function Get-VerifiedPSModulePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ModuleName
    )

    $evidence = Test-InstalledPSModule -ModuleName $ModuleName
    if ($evidence.Valid) { return $evidence.Path }
    return $null
}

function Install-VerifiedPSModule {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ModuleName
    )

    $spec = Get-AcquisitionManifestEntry -Name $ModuleName
    if ($spec.Kind -ne "PowerShellModule") {
        throw "'$ModuleName' is not approved for module acquisition"
    }

    $existing = Test-InstalledPSModule -ModuleName $ModuleName
    if ($existing.Valid) {
        [void](Add-AcquisitionProvenance -Name $spec.Name -Version $existing.Version `
            -SourceUri $spec.Uri -Sha256 $spec.Sha256 -Architecture "neutral" `
            -InstallPath $existing.InstallPath -Status "VerifiedExisting" `
            -Evidence @("Installed payload SHA-256 $($existing.Sha256)"))
        return $true
    }

    $temporaryDirectory = New-SystemUpdateProTemporaryDirectory -Purpose $ModuleName
    $packagePath = Join-Path $temporaryDirectory "$ModuleName.nupkg"
    $extractPath = Join-Path $temporaryDirectory "payload"
    $moduleBase = Split-Path -Parent (Get-PSModuleInstallPath -ModuleName $ModuleName -Version $spec.ExactVersion)
    $destinationPath = Get-PSModuleInstallPath -ModuleName $ModuleName -Version $spec.ExactVersion
    $stagingPath = Join-Path $moduleBase ".$ModuleName.$([guid]::NewGuid().ToString('N')).staging"
    $backupPath = Join-Path $moduleBase ".$ModuleName.$([guid]::NewGuid().ToString('N')).untrusted"
    $committed = $false

    try {
        $download = Invoke-VerifiedDownload -Name "$ModuleName module" -Uri $spec.Uri `
            -AllowedHosts $spec.AllowedHosts -ExpectedSha256 $spec.Sha256 `
            -DestinationPath $packagePath

        New-Item -ItemType Directory -Path $extractPath -Force -ErrorAction Stop | Out-Null
        Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
        [IO.Compression.ZipFile]::ExtractToDirectory($packagePath, $extractPath)

        foreach ($metadataDirectory in @("_rels", "package")) {
            $metadataPath = Join-Path $extractPath $metadataDirectory
            Remove-Item -LiteralPath $metadataPath -Recurse -Force -ErrorAction SilentlyContinue
        }
        foreach ($metadataFile in @("[Content_Types].xml", "PSGetModuleInfo.xml")) {
            Remove-Item -LiteralPath (Join-Path $extractPath $metadataFile) -Force -ErrorAction SilentlyContinue
        }
        Get-ChildItem -LiteralPath $extractPath -Filter "*.nuspec" -File -ErrorAction SilentlyContinue |
            Remove-Item -Force -ErrorAction SilentlyContinue

        $manifestPath = Join-Path $extractPath "$ModuleName.psd1"
        if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) {
            throw "$ModuleName package does not contain its module manifest"
        }
        $moduleManifest = Import-PowerShellDataFile -LiteralPath $manifestPath -ErrorAction Stop
        if ([string]$moduleManifest.ModuleVersion -ne [string]$spec.ExactVersion) {
            throw "$ModuleName package declares version '$($moduleManifest.ModuleVersion)', expected $($spec.ExactVersion)"
        }
        $payloadHash = Get-DirectoryPayloadHash -Path $extractPath
        if ($payloadHash -ne [string]$spec.PayloadTreeSha256) {
            throw "$ModuleName extracted payload was rejected before installation: tree hash mismatch"
        }

        New-Item -ItemType Directory -Path $moduleBase -Force -ErrorAction Stop | Out-Null
        New-Item -ItemType Directory -Path $stagingPath -Force -ErrorAction Stop | Out-Null
        Get-ChildItem -LiteralPath $extractPath -Force -ErrorAction Stop |
            Copy-Item -Destination $stagingPath -Recurse -Force -ErrorAction Stop
        if ((Get-DirectoryPayloadHash -Path $stagingPath) -ne [string]$spec.PayloadTreeSha256) {
            throw "$ModuleName staging copy failed payload verification"
        }

        if (Test-Path -LiteralPath $destinationPath) {
            Move-Item -LiteralPath $destinationPath -Destination $backupPath -ErrorAction Stop
        }
        try {
            Move-Item -LiteralPath $stagingPath -Destination $destinationPath -ErrorAction Stop
            $installed = Test-InstalledPSModule -ModuleName $ModuleName
            if (-not $installed.Valid) {
                throw "$ModuleName failed installed verification: $($installed.Reason)"
            }
            $committed = $true
        } catch {
            Remove-Item -LiteralPath $destinationPath -Recurse -Force -ErrorAction SilentlyContinue
            if (Test-Path -LiteralPath $backupPath) {
                Move-Item -LiteralPath $backupPath -Destination $destinationPath -ErrorAction SilentlyContinue
            }
            throw
        }

        Remove-Item -LiteralPath $backupPath -Recurse -Force -ErrorAction SilentlyContinue
        [void](Add-AcquisitionProvenance -Name $spec.Name -Version $spec.ExactVersion `
            -SourceUri $spec.Uri -Sha256 $download.Sha256 -Architecture "neutral" `
            -InstallPath $destinationPath -Status "Installed" `
            -Evidence @("Installed payload SHA-256 $($spec.PayloadTreeSha256)"))
        Write-Log "$ModuleName $($spec.ExactVersion) installed from a verified package" "SUCCESS"
        return $true
    } finally {
        if (-not $committed) {
            Remove-Item -LiteralPath $stagingPath -Recurse -Force -ErrorAction SilentlyContinue
        }
        Remove-SystemUpdateProTemporaryDirectory -Path $temporaryDirectory
    }
}

function Install-PSModuleWithRetry {
    param(
        [string]$ModuleName
    )

    $spec = Get-AcquisitionManifestEntry -Name $ModuleName
    $existing = Test-InstalledPSModule -ModuleName $ModuleName
    if ($existing.Valid) {
        [void](Add-AcquisitionProvenance -Name $spec.Name -Version $existing.Version `
            -SourceUri $spec.Uri -Sha256 $spec.Sha256 -Architecture "neutral" `
            -InstallPath $existing.InstallPath -Status "VerifiedExisting" `
            -Evidence @("Installed payload SHA-256 $($existing.Sha256)"))
        Write-Log "$ModuleName v$($existing.Version) verified" "DEBUG"
        return $true
    }

    Write-Log "Installing verified $ModuleName $($spec.ExactVersion) module..." "INFO"

    if ($DryRun) {
        Write-Log "Would install $ModuleName from its pinned PowerShell Gallery package without changing repository trust" "INFO"
        return $true
    }

    return Invoke-WithRetry -OperationName "Install $ModuleName" -ScriptBlock {
        return Install-VerifiedPSModule -ModuleName $ModuleName
    }
}

# ============================================================================
# DELL COMMAND UPDATE
# ============================================================================

function Get-InstalledExecutableEvidence {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter(Mandatory = $true)]
        [string]$MinimumVersion,
        [Parameter(Mandatory = $true)]
        [string]$PublisherPattern,
        [string]$ExpectedSha256 = ""
    )

    try {
        if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
            throw "Executable is missing: $Path"
        }
        $file = Get-Item -LiteralPath $Path -ErrorAction Stop
        $rawVersion = if (-not [string]::IsNullOrWhiteSpace([string]$file.VersionInfo.ProductVersion)) {
            [string]$file.VersionInfo.ProductVersion
        } else {
            [string]$file.VersionInfo.FileVersion
        }
        $version = ConvertTo-SafeVersion -Value $rawVersion
        if (-not (Test-VersionAtLeast -Version $version -MinimumVersion $MinimumVersion)) {
            throw "Executable version '$rawVersion' is below the approved minimum $MinimumVersion"
        }

        $hash = (Get-FileHash -LiteralPath $Path -Algorithm SHA256 -ErrorAction Stop).Hash.ToUpperInvariant()
        if (-not [string]::IsNullOrWhiteSpace($ExpectedSha256) -and
            $hash -ne $ExpectedSha256.ToUpperInvariant()) {
            throw "Installed executable SHA-256 does not match the pinned payload"
        }
        $signature = Get-AuthenticodeEvidence -Path $Path -PublisherPattern $PublisherPattern
        if (-not $signature.Valid) { throw $signature.Reason }

        return [PSCustomObject]@{
            Valid = $true; Path = $Path; Version = [string]$version; Sha256 = $hash
            Publisher = $signature.Subject; Thumbprint = $signature.Thumbprint; Reason = ""
        }
    } catch {
        return [PSCustomObject]@{
            Valid = $false; Path = $Path; Version = ""; Sha256 = ""; Publisher = ""
            Thumbprint = ""; Reason = $_.Exception.Message
        }
    }
}

function Test-WingetPackageContractText {
    param(
        [Parameter(Mandatory = $true)]
        [string]$OutputText,
        [Parameter(Mandatory = $true)]
        [string]$ExpectedUri,
        [Parameter(Mandatory = $true)]
        [string]$ExpectedSha256,
        [Parameter(Mandatory = $true)]
        [string]$ExpectedPublisher
    )

    return (
        $OutputText.IndexOf($ExpectedUri, [StringComparison]::OrdinalIgnoreCase) -ge 0 -and
        $OutputText.IndexOf($ExpectedSha256, [StringComparison]::OrdinalIgnoreCase) -ge 0 -and
        $OutputText.IndexOf($ExpectedPublisher, [StringComparison]::OrdinalIgnoreCase) -ge 0
    )
}

function Test-DellWinGetPackageContract {
    $spec = Get-AcquisitionManifestEntry -Name "DellCommandUpdate"
    $output = & winget show --exact --id $spec.PackageId --version $spec.ExactVersion `
        --source $spec.PackageSource --accept-source-agreements --disable-interactivity 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "WinGet could not resolve $($spec.PackageId) $($spec.ExactVersion) (exit $LASTEXITCODE)"
    }
    $outputText = ($output | Out-String)
    if (-not (Test-WingetPackageContractText -OutputText $outputText `
        -ExpectedUri $spec.Uri -ExpectedSha256 $spec.Sha256 `
        -ExpectedPublisher $spec.PublisherDisplayName)) {
        throw "WinGet source metadata for $($spec.PackageId) $($spec.ExactVersion) does not match the pinned URI, SHA-256, and publisher"
    }
    return $true
}

function Get-DCUPath {
    $spec = Get-AcquisitionManifestEntry -Name "DellCommandUpdate"
    $candidates = @(
        "${env:ProgramFiles}\Dell\CommandUpdate\dcu-cli.exe",
        "${env:ProgramFiles(x86)}\Dell\CommandUpdate\dcu-cli.exe"
    )
    foreach ($path in $candidates) {
        $evidence = Get-InstalledExecutableEvidence -Path $path -MinimumVersion $spec.MinimumVersion `
            -PublisherPattern $spec.PublisherPattern
        if ($evidence.Valid) {
            [void](Add-AcquisitionProvenance -Name $spec.Name -Version $evidence.Version `
                -SourceUri $spec.Uri -Sha256 $evidence.Sha256 -Publisher $evidence.Publisher `
                -Thumbprint $evidence.Thumbprint -Architecture (Get-SystemArchitecture) `
                -InstallPath $path -Status "VerifiedExisting" `
                -Evidence @("WinGet package $($spec.PackageId) $($spec.ExactVersion)"))
            return $path
        }
    }
    return $null
}

function Get-DellInventoryCollectorEvidence {
    $spec = Get-AcquisitionManifestEntry -Name "DellCommandUpdate"
    return Get-InstalledExecutableEvidence -Path $spec.InventoryCollectorPath `
        -MinimumVersion $spec.InventoryCollectorMinimum `
        -PublisherPattern $spec.InventoryPublisherPattern
}

function Test-DellInventoryCollector {
    $spec = Get-AcquisitionManifestEntry -Name "DellCommandUpdate"
    $evidence = Get-DellInventoryCollectorEvidence
    if (-not $evidence.Valid) { return $false }

    [void](Add-AcquisitionProvenance -Name "Dell Inventory Collector" -Version $evidence.Version `
        -SourceUri $spec.Uri -Sha256 $evidence.Sha256 -Publisher $evidence.Publisher `
        -Thumbprint $evidence.Thumbprint -Architecture (Get-SystemArchitecture) `
        -InstallPath $evidence.Path -Status "VerifiedExisting" `
        -Evidence @("Minimum remediated version $($spec.InventoryCollectorMinimum)"))
    return $true
}

function Update-DellInventoryCollector {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseShouldProcessForStateChangingFunctions", "", Justification = "Performs Dell's documented signed inventory refresh prerequisite.")]
    param(
        [Parameter(Mandatory = $true)]
        [string]$DCUPath
    )

    if (Test-DellInventoryCollector) { return $true }
    if ($DryRun) { return $false }

    $spec = Get-AcquisitionManifestEntry -Name "DellCommandUpdate"
    $inventoryLog = Join-Path $LogPath "DCU_InventoryCollector_$((Get-Date).ToString('yyyyMMdd_HHmmss')).log"
    Write-Log "Refreshing Dell Inventory Collector before update execution..." "INFO"
    $process = Start-Process -FilePath $DCUPath `
        -ArgumentList "/scan -silent -outputLog=`"$inventoryLog`"" `
        -Wait -NoNewWindow -PassThru -ErrorAction Stop
    if ($process.ExitCode -notin @(0, 500)) {
        throw "Dell Inventory Collector refresh scan exited with $($process.ExitCode)"
    }

    $deadline = (Get-Date).AddSeconds(30)
    do {
        if (Test-DellInventoryCollector) { return $true }
        Start-Sleep -Seconds 2
    } while ((Get-Date) -lt $deadline)

    $evidence = Get-DellInventoryCollectorEvidence
    throw "Dell Inventory Collector was rejected before update execution: $($evidence.Reason). Version $($spec.InventoryCollectorMinimum) or later is required"
}

function Install-DellCommandUpdate {
    param([switch]$ForceRepair)

    $spec = Get-AcquisitionManifestEntry -Name "DellCommandUpdate"
    $architecture = Get-SystemArchitecture
    if ($architecture -notin @($spec.Architectures)) {
        Write-Log "Dell Command Update $($spec.ExactVersion) is not approved for architecture '$architecture'" "WARNING"
        return $false
    }

    $existingPath = Get-DCUPath
    if (-not $ForceRepair -and $existingPath -and (Test-DellInventoryCollector)) {
        return $true
    }

    Write-Log "Installing verified Dell Command Update $($spec.ExactVersion)..." "INFO"

    if ($DryRun) {
        Write-Log "Would install exact package $($spec.PackageId) $($spec.ExactVersion) and verify Inventory Collector $($spec.InventoryCollectorMinimum)+" "INFO"
        return $true
    }

    if (-not (Test-WingetInstalled)) {
        if (-not (Install-Winget)) { return $false }
    }

    return Invoke-WithRetry -OperationName "Install DCU" -ScriptBlock {
        & winget source update --disable-interactivity 2>&1 | Out-Null
        if ($LASTEXITCODE -ne 0) { throw "WinGet source update exited with $LASTEXITCODE" }
        if (-not (Test-DellWinGetPackageContract)) {
            throw "Dell Command Update source contract verification failed"
        }

        $wingetArgs = @(
            "install", "--exact", "--id", $spec.PackageId, "--version", $spec.ExactVersion,
            "--source", $spec.PackageSource, "--scope", "machine", "--silent", "--force",
            "--accept-package-agreements", "--accept-source-agreements", "--disable-interactivity"
        )

        $installProcess = Start-Process -FilePath "winget" -ArgumentList $wingetArgs `
            -Wait -NoNewWindow -PassThru -ErrorAction Stop
        if ($installProcess.ExitCode -ne 0) {
            throw "WinGet rejected Dell Command Update installation with exit $($installProcess.ExitCode)"
        }
        Start-Sleep -Seconds 3

        $dcuPath = Get-DCUPath
        if (-not $dcuPath) { throw "DCU failed version or publisher verification after installation" }
        if (-not (Update-DellInventoryCollector -DCUPath $dcuPath)) {
            throw "Dell Inventory Collector $($spec.InventoryCollectorMinimum)+ was not verified"
        }

        $dcuEvidence = Get-InstalledExecutableEvidence -Path $dcuPath `
            -MinimumVersion $spec.MinimumVersion -PublisherPattern $spec.PublisherPattern
        [void](Add-AcquisitionProvenance -Name $spec.Name -Version $dcuEvidence.Version `
            -SourceUri $spec.Uri -Sha256 $dcuEvidence.Sha256 -Publisher $dcuEvidence.Publisher `
            -Thumbprint $dcuEvidence.Thumbprint -Architecture $architecture -InstallPath $dcuPath `
            -Status "Installed" -Evidence @(
                "WinGet source contract SHA-256 $($spec.Sha256)",
                "Package $($spec.PackageId) $($spec.ExactVersion)"
            ))
        Write-Log "Dell Command Update and Inventory Collector passed provenance verification" "SUCCESS"
        return $true
    }
}

function Repair-DellServices {
    if ($DryRun) { return $true }

    $serviceName = "DellClientManagementService"
    $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue

    if (-not $service) {
        Write-Log "Dell service not found - repairing the pinned DCU installation" "WARNING"
        return (Install-DellCommandUpdate -ForceRepair)
    }

    if ($service.Status -ne "Running") {
        try {
            [void](Set-JournaledServiceState -ServiceName $serviceName -DesiredStatus "Running" `
                -StartupType "Automatic" -Scope "Dell" `
                -RecoveryAction "Restore the exact Dell Client Management Service status and startup mode")
            $service = Get-Service -Name $serviceName -ErrorAction Stop
        } catch {
            Write-Log "Dell Client Management Service repair failed: $($_.Exception.Message)" "WARNING"
            return $false
        }
    }

    return ($service.Status -eq "Running")
}

function Invoke-DellUpdate {
    param(
        [switch]$IncludeBIOS,
        [AllowNull()][object]$SystemInfo = $null,
        [AllowNull()][object]$FirmwarePrerequisites = $null
    )

    $result = @{
        Success = $false; Status = "Failed"; RebootRequired = $false
        UpdateCount = 0; Available = 0; Attempted = 0; Installed = 0; Failed = 0; Skipped = 0
        ExitCode = $null; HResult = $null; Items = @(); Evidence = @(); Message = ""
        FirmwareReadiness = $null
    }

    try {
        Write-Log "========== DELL COMMAND UPDATE ==========" "HEADER"

        $sysInfo = if ($null -ne $SystemInfo) { $SystemInfo } else { Get-SystemInfo }
        Add-SensitiveEvidenceValue -Value ([string]$sysInfo.SerialNumber)
        Write-Log "Service Tag: $($sysInfo.SerialNumber)" "INFO"

        $dcuPath = Get-DCUPath
        $inventoryCollectorReady = $dcuPath -and (Test-DellInventoryCollector)
        if (-not $dcuPath -or -not $inventoryCollectorReady) {
            if (-not (Install-DellCommandUpdate)) {
                $result.Message = "Dell Command Update or its remediated Inventory Collector could not be verified; repair WinGet/network access and rerun"
                $result.Failed = 1
            }
            $dcuPath = Get-DCUPath
            $inventoryCollectorReady = $dcuPath -and (Test-DellInventoryCollector)
        }

        if (-not $dcuPath -or -not $inventoryCollectorReady) {
            $readiness = Test-FirmwareReadiness -Requested:$IncludeBIOS -Provider "Dell" -SystemInfo $sysInfo `
                -ToolState "Unknown" -ToolMessage "Dell Command Update and Inventory Collector 13.8.0+ are not both verified. Install or repair DCU and rerun its applicability scan" `
                -SupportsBitLockerAutoSuspend $true -Prerequisites $FirmwarePrerequisites
            $result.FirmwareReadiness = $readiness
            $result.Items += New-FirmwareReadinessItem -Readiness $readiness
            $result.Skipped = $(if ($IncludeBIOS) { 1 } else { 0 })
            if ($DryRun) {
                $result.Success = $true
                $result.Status = $(if ($IncludeBIOS) { "Partial" } else { "Succeeded" })
                $result.Message = "DCU would be installed before scanning; $($readiness.Message)"
            } else {
                $result.Failed = [math]::Max(1, $result.Failed)
                $result.Message = "Dell Command Update dependency trust could not be established after installation"
            }
            Write-Log $result.Message $(if ($result.Success) { "WARNING" } else { "ERROR" })
            return $result
        }

        if (-not (Repair-DellServices)) {
            Write-Log "Dell Client Management Service was not verified; the applicability scan must succeed before any update is applied" "WARNING"
        }

        if (-not $DryRun) {
            # Disable nonessential Dell services while preserving the DCU service.
            Get-Service -ErrorAction SilentlyContinue | Where-Object {
                ($_.DisplayName -like "*Dell*" -or $_.Name -like "*DDV*" -or $_.Name -like "*SupportAssist*") -and
                $_.Name -ne "DellClientManagementService"
            } | ForEach-Object {
                [void](Set-JournaledServiceState -ServiceName $_.Name -DesiredStatus "Stopped" `
                    -StartupType "Disabled" -Scope "Dell" `
                    -RecoveryAction "Restore the exact pre-update Dell service status and startup mode")
            }
        }

        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $scanLog = Join-Path $LogPath "DCU_Scan_$timestamp.log"
        $scanTypes = if ($IncludeBIOS) {
            @("bios", "firmware", "driver", "application", "others")
        } else {
            @("driver", "application", "others")
        }
        $scanArgs = @(
            "/scan", "-silent", "-updateSeverity=security,critical,recommended",
            "-updateType=$($scanTypes -join ',')", "-outputLog=`"$scanLog`""
        )
        Write-Log "Verifying Dell model/tool applicability..." "INFO"
        $scanProcess = Start-Process -FilePath $dcuPath -ArgumentList ($scanArgs -join " ") `
            -Wait -NoNewWindow -PassThru -ErrorAction Stop
        [void](Protect-EvidenceFile -Path $scanLog `
            -SensitiveValues @([string]$sysInfo.SerialNumber))
        $result.Attempted++
        $result.ExitCode = $scanProcess.ExitCode
        $result.Evidence = @($scanLog)
        $scanSucceeded = $scanProcess.ExitCode -in @(0, 500)
        $readiness = Test-FirmwareReadiness -Requested:$IncludeBIOS -Provider "Dell" -SystemInfo $sysInfo `
            -ToolState $(if ($scanSucceeded) { "Ready" } else { "Unknown" }) `
            -ToolMessage $(if ($scanSucceeded) {
                "Dell Command Update completed an applicability scan for this model"
            } else {
                "DCU scan exit $($scanProcess.ExitCode) did not verify support. Review $scanLog, repair DCU, and rerun"
            }) -SupportsBitLockerAutoSuspend $true -Prerequisites $FirmwarePrerequisites
        $result.FirmwareReadiness = $readiness
        $result.Items += New-FirmwareReadinessItem -Readiness $readiness
        $policy = Get-FirmwareUpdatePolicy -Readiness $readiness

        if (-not $scanSucceeded) {
            $result.Failed = 1
            $result.Message = "DCU applicability scan failed with exit $($scanProcess.ExitCode); no updates were applied"
            Write-Log $result.Message "ERROR"
            return $result
        }

        if ($IncludeBIOS -and -not $readiness.Safe) {
            $result.Skipped++
            $result.Status = "Partial"
            Write-Log $readiness.Message "WARNING"
        }

        if ($DryRun) {
            $result.Success = $true
            if ($result.Status -ne "Partial") { $result.Status = "Succeeded" }
            $result.Message = "DCU applicability scan completed (exit: $($scanProcess.ExitCode)); no updates installed. $($readiness.Message)"
            Write-Log $result.Message $(if ($readiness.Safe -or -not $IncludeBIOS) { "INFO" } else { "WARNING" })
            return $result
        }

        if ($scanProcess.ExitCode -eq 500) {
            $result.Success = $true
            if ($result.Status -ne "Partial") { $result.Status = "Succeeded" }
            $result.Message = "No applicable Dell updates were found; $($readiness.Message)"
            Write-Log $result.Message "SUCCESS"
            return $result
        }

        $dcuLog = Join-Path $LogPath "DCU_Apply_$timestamp.log"
        $dcuArgs = @(
            "/applyUpdates", "-silent", "-updateSeverity=security,critical,recommended", "-reboot=disable",
            "-updateType=$($policy.DellUpdateTypes -join ',')", "-outputLog=`"$dcuLog`""
        )
        if ($policy.DellAutoSuspendBitLocker) {
            $dcuArgs += "-autoSuspendBitLocker=enable"
        }
        $result.Evidence += $dcuLog
        Write-Log "Applying Dell updates with types: $($policy.DellUpdateTypes -join ', ')..." "INFO"

        $attempts = 0
        while ($attempts -lt $MaxRetries) {
            $attempts++
            $process = Start-Process -FilePath $dcuPath -ArgumentList ($dcuArgs -join " ") `
                -Wait -NoNewWindow -PassThru -ErrorAction Stop
            $result.Attempted++
            $result.ExitCode = $process.ExitCode

            switch ($process.ExitCode) {
                0   { $result.Success = $true; $result.Message = "Updates applied"; break }
                1   { $result.Success = $true; $result.RebootRequired = $true; $result.Message = "Updates applied - reboot required"; break }
                500 { $result.Success = $true; $result.Message = "No updates available"; break }
                3000 {
                    if ($attempts -lt $MaxRetries -and (Repair-DellServices)) { continue }
                    $result.Message = "Dell Client Management Service is not running; repair DCU and rerun"
                    break
                }
                3003 {
                    if ($attempts -lt $MaxRetries) { Start-Sleep -Seconds 30; continue }
                    $result.Message = "DCU remained busy after $MaxRetries attempts; close other DCU sessions and rerun"
                    break
                }
                default { $result.Message = "DCU apply exit code $($process.ExitCode); review $dcuLog and rerun"; break }
            }
            break
        }
        [void](Protect-EvidenceFile -Path $dcuLog `
            -SensitiveValues @([string]$sysInfo.SerialNumber))

        if ($result.Success) {
            if ($result.Status -ne "Partial") { $result.Status = "Succeeded" }
            $result.Items += New-UpdateItemResult -Name "Dell update application" -Status "Succeeded" `
                -ProviderCode $result.ExitCode -RebootRequired $result.RebootRequired -Message $result.Message -Evidence @($dcuLog)
        } else {
            $result.Failed = 1
            $result.Status = "Failed"
            $result.Items += New-UpdateItemResult -Name "Dell update application" -Status "Failed" `
                -ProviderCode $result.ExitCode -Message $result.Message -Evidence @($dcuLog)
        }
        Write-Log $result.Message $(if ($result.Success) { "SUCCESS" } else { "WARNING" })
        $script:OEMUpdateCount = $result.UpdateCount
        return $result
    } catch {
        $result.Success = $false
        $result.Status = "Failed"
        $result.Failed = [math]::Max(1, $result.Failed)
        $result.HResult = $_.Exception.HResult
        $result.Message = "Dell update error: $($_.Exception.Message)"
        if ($null -eq $result.FirmwareReadiness) {
            $readiness = Test-FirmwareReadiness -Requested:$IncludeBIOS -Provider "Dell" -SystemInfo $SystemInfo `
                -ToolState "Unknown" -ToolMessage "Dell tooling failed before support could be verified: $($_.Exception.Message). Repair DCU and rerun" `
                -SupportsBitLockerAutoSuspend $true -Prerequisites $FirmwarePrerequisites
            $result.FirmwareReadiness = $readiness
            $result.Items += New-FirmwareReadinessItem -Readiness $readiness
        }
        Write-Log $result.Message "ERROR"
        return $result
    } finally {
        if (-not $DryRun -and -not (Restore-MutationJournalScope -Scope "Dell")) {
            $result.Success = $false
            $result.Status = $(if ($result.Installed -gt 0) { "Partial" } else { "Failed" })
            $result.Failed = [math]::Max(1, $result.Failed)
            $restoreMessage = "Dell service state restoration could not be verified"
            $result.Message = $(if ([string]::IsNullOrWhiteSpace($result.Message)) {
                $restoreMessage
            } else {
                "$($result.Message); $restoreMessage"
            })
            Write-Log $restoreMessage "ERROR"
        }
    }
}

# ============================================================================
# LENOVO LSUClient
# ============================================================================

function Invoke-LenovoUpdate {
    param(
        [switch]$IncludeBIOS,
        [AllowNull()][object]$SystemInfo = $null,
        [AllowNull()][object]$FirmwarePrerequisites = $null
    )

    $result = @{
        Success = $false; Status = "Failed"; RebootRequired = $false
        UpdateCount = 0; Available = 0; Attempted = 0; Installed = 0; Failed = 0; Skipped = 0
        ExitCode = $null; HResult = $null; Items = @(); Evidence = @(); Message = ""
        FirmwareReadiness = $null
    }

    Write-Log "========== LENOVO SYSTEM UPDATE ==========" "HEADER"

    $sysInfo = if ($null -ne $SystemInfo) { $SystemInfo } else { Get-SystemInfo }
    Add-SensitiveEvidenceValue -Value ([string]$sysInfo.SerialNumber)
    Write-Log "Serial: $($sysInfo.SerialNumber)" "INFO"

    if (-not (Install-PSModuleWithRetry -ModuleName "LSUClient")) {
        $readiness = Test-FirmwareReadiness -Requested:$IncludeBIOS -Provider "Lenovo" -SystemInfo $sysInfo `
            -ToolState "Unknown" -ToolMessage "LSUClient could not be installed. Repair PowerShell Gallery/module access and rerun" `
            -Prerequisites $FirmwarePrerequisites
        $result.FirmwareReadiness = $readiness
        $result.Items += New-FirmwareReadinessItem -Readiness $readiness
        $result.Message = "Failed to install LSUClient; no Lenovo updates were attempted"
        $result.Failed = 1
        Write-Log $result.Message "ERROR"
        return $result
    }

    $lsuClientPath = Get-VerifiedPSModulePath -ModuleName "LSUClient"
    if (-not $lsuClientPath) {
        $readiness = Test-FirmwareReadiness -Requested:$IncludeBIOS -Provider "Lenovo" -SystemInfo $sysInfo `
            -ToolState "Unknown" -ToolMessage "LSUClient is not installed, so model support cannot be scanned. Install the module and rerun" `
            -Prerequisites $FirmwarePrerequisites
        $result.FirmwareReadiness = $readiness
        $result.Items += New-FirmwareReadinessItem -Readiness $readiness
        if ($DryRun) {
            $result.Success = $true
            $result.Status = $(if ($IncludeBIOS) { "Partial" } else { "Succeeded" })
            $result.Skipped = $(if ($IncludeBIOS) { 1 } else { 0 })
            $result.Message = "LSUClient would be installed before scanning; $($readiness.Message)"
        } else {
            $result.Failed = 1
            $result.Message = "LSUClient was not found after installation"
        }
        Write-Log $result.Message $(if ($result.Success) { "WARNING" } else { "ERROR" })
        return $result
    }

    try {
        Import-Module $lsuClientPath -Force -ErrorAction Stop

        Write-Log "Scanning for updates..." "INFO"
        $allUpdates = @(Get-LSUpdate -ErrorAction Stop | Where-Object { $_.Installer.Unattended -eq $true })
        $readiness = Test-FirmwareReadiness -Requested:$IncludeBIOS -Provider "Lenovo" -SystemInfo $sysInfo `
            -ToolState "Ready" -ToolMessage "LSUClient completed an applicability scan for this model" `
            -Prerequisites $FirmwarePrerequisites
        $result.FirmwareReadiness = $readiness
        $result.Items += New-FirmwareReadinessItem -Readiness $readiness
        $policy = Get-FirmwareUpdatePolicy -Readiness $readiness

        $firmwareUpdates = @($allUpdates | Where-Object { Test-OEMUpdateIsFirmware -Update $_ })
        $updates = @(
            if ($policy.LenovoExcludeFirmware) {
                $allUpdates | Where-Object { -not (Test-OEMUpdateIsFirmware -Update $_) }
            } else {
                $allUpdates
            }
        )

        if ($policy.LenovoExcludeFirmware) {
            $blockedStatus = if (-not $IncludeBIOS) {
                "Skipped"
            } elseif ($readiness.Status -eq "Unknown") {
                "Unknown"
            } else {
                "Blocked"
            }
            foreach ($firmwareUpdate in $firmwareUpdates) {
                $firmwareName = [string](Get-ResultValue -Result $firmwareUpdate -Names @("Title", "Name") -Default "Lenovo firmware update")
                $result.Items += New-UpdateItemResult -Name $firmwareName `
                    -Id ([string](Get-ResultValue -Result $firmwareUpdate -Names @("ID", "Id") -Default "")) `
                    -Status $blockedStatus -Message $readiness.Message
                $result.Skipped++
            }
        }

        if ($IncludeBIOS -and -not $readiness.Safe) {
            $result.Status = "Partial"
            Write-Log $readiness.Message "WARNING"
        }

        if ($updates.Count -eq 0) {
            $result.Success = $true
            if ($result.Status -ne "Partial") { $result.Status = "Succeeded" }
            $result.Message = if ($firmwareUpdates.Count -gt 0 -and $policy.LenovoExcludeFirmware) {
                "No non-firmware Lenovo updates are available; $($firmwareUpdates.Count) firmware update(s) were not attempted. $($readiness.Message)"
            } else {
                "No applicable Lenovo updates are available; $($readiness.Message)"
            }
            Write-Log $result.Message $(if ($result.Status -eq "Partial") { "WARNING" } else { "SUCCESS" })
            $script:OEMUpdateCount = 0
            return $result
        }

        if ($DryRun) {
            $result.UpdateCount = $updates.Count
            $result.Available = $updates.Count
            $result.Success = $true
            if ($result.Status -ne "Partial") { $result.Status = "Succeeded" }
            $result.Message = "$($updates.Count) nonblocked Lenovo update(s) available (dry run - not installed); $($readiness.Message)"
            Write-Log "Available Lenovo updates:" "INFO"
            foreach ($u in $updates) {
                Write-Log "  -- $($u.Title) ($($u.Category))" "INFO"
                [void]$script:OEMUpdates.Add("$($u.Title) ($($u.Category))")
                $result.Items += New-UpdateItemResult -Name $u.Title -Id $u.ID -Status "Available" `
                    -Message ([string]$u.Category)
            }
            Write-Log $result.Message $(if ($result.Status -eq "Partial") { "WARNING" } else { "INFO" })
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

        $installationItems = @($result.Items | Where-Object { $_.Name -ne "Firmware safety" -and $_.Status -notin @("Skipped", "Blocked", "Unknown") })
        $result.Success = ($result.Failed -eq 0 -and $installationItems.Count -eq $result.Attempted)
        if (-not $result.Success) {
            $result.Status = $(if ($result.Installed -gt 0) { "Partial" } else { "Failed" })
        } elseif ($result.Status -ne "Partial") {
            $result.Status = "Succeeded"
        }
        $result.Message = "Installed: $($result.Installed), Failed: $($result.Failed)"
        Write-Log $result.Message $(if ($result.Success) { "SUCCESS" } else { "WARNING" })

    } catch {
        $result.Status = "Failed"
        $result.Message = "Lenovo error: $($_.Exception.Message)"
        $result.Failed = [math]::Max(1, $result.Failed)
        $result.HResult = $_.Exception.HResult
        if ($null -eq $result.FirmwareReadiness) {
            $readiness = Test-FirmwareReadiness -Requested:$IncludeBIOS -Provider "Lenovo" -SystemInfo $sysInfo `
                -ToolState "Unknown" -ToolMessage "LSUClient failed before support could be verified: $($_.Exception.Message). Repair the module and rerun" `
                -Prerequisites $FirmwarePrerequisites
            $result.FirmwareReadiness = $readiness
            $result.Items += New-FirmwareReadinessItem -Readiness $readiness
        }
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
    $spec = Get-AcquisitionManifestEntry -Name "HPIA"
    if ([string]::IsNullOrWhiteSpace([string]$script:HPIAInstallRoot)) {
        $script:HPIAInstallRoot = "C:\ProgramData\SystemUpdatePro\HPIA"
    }
    $searchPaths = @(
        $script:HPIAInstallRoot,
        "C:\SWSetup\SP*",
        "${env:ProgramFiles}\HP\HPIA"
    )

    foreach ($path in $searchPaths) {
        $candidates = @(Get-ChildItem -Path $path -Filter "HPImageAssistant.exe" -File -Recurse -ErrorAction SilentlyContinue)
        foreach ($found in $candidates) {
            $evidence = Get-InstalledExecutableEvidence -Path $found.FullName `
                -MinimumVersion $spec.InstalledMinimumVersion `
                -PublisherPattern $spec.InstalledPublisherPattern
            if ($evidence.Valid) {
                [void](Add-AcquisitionProvenance -Name $spec.Name -Version $evidence.Version `
                    -SourceUri $spec.Uri -Sha256 $evidence.Sha256 -Publisher $evidence.Publisher `
                    -Thumbprint $evidence.Thumbprint -Architecture (Get-SystemArchitecture) `
                    -InstallPath $found.FullName -Status "VerifiedExisting" `
                    -Evidence @("Minimum remediated version $($spec.MinimumVersion)"))
                return $found.FullName
            }
        }
    }
    return $null
}

function Install-HPIA {
    $spec = Get-AcquisitionManifestEntry -Name "HPIA"
    $architecture = Get-SystemArchitecture
    if ($architecture -notin @($spec.Architectures)) {
        Write-Log "HP Image Assistant $($spec.ExactVersion) is not approved for architecture '$architecture'" "WARNING"
        return $false
    }
    if (Get-HPIAPath) { return $true }

    Write-Log "Installing verified HP Image Assistant $($spec.ExactVersion)..." "INFO"

    if ($DryRun) {
        Write-Log "Would download HP Image Assistant from its pinned HP release and verify both signatures" "INFO"
        return $true
    }

    return Invoke-WithRetry -OperationName "Install HPIA" -ScriptBlock {
        $hpiaDir = $script:HPIAInstallRoot
        $hpiaParent = Split-Path -Parent $hpiaDir
        $temporaryDirectory = New-SystemUpdateProTemporaryDirectory -Purpose "HPIA"
        $extractPath = Join-Path $temporaryDirectory "payload"
        $installer = Join-Path $temporaryDirectory "hp-hpia-$($spec.ExactVersion).exe"
        $stagingPath = "$hpiaDir.staging.$([guid]::NewGuid().ToString('N'))"
        $backupPath = "$hpiaDir.untrusted.$([guid]::NewGuid().ToString('N'))"
        $committed = $false

        try {
            $download = Invoke-VerifiedDownload -Name $spec.Name -Uri $spec.Uri `
                -AllowedHosts $spec.AllowedHosts -ExpectedSha256 $spec.Sha256 `
                -PublisherPattern $spec.PublisherPattern -DestinationPath $installer
            New-Item -ItemType Directory -Path $extractPath -Force -ErrorAction Stop | Out-Null
            $extractProcess = Start-Process -FilePath $installer `
                -ArgumentList "/s /e /f `"$extractPath`"" -Wait -NoNewWindow -PassThru -ErrorAction Stop

            $extractedExecutable = Get-ChildItem -LiteralPath $extractPath -Filter "HPImageAssistant.exe" `
                -File -Recurse -ErrorAction Stop | Select-Object -First 1
            if (-not $extractedExecutable) {
                throw "HPIA extraction exited with $($extractProcess.ExitCode) without producing HPImageAssistant.exe"
            }
            $extractedEvidence = Get-InstalledExecutableEvidence -Path $extractedExecutable.FullName `
                -MinimumVersion $spec.InstalledMinimumVersion `
                -PublisherPattern $spec.InstalledPublisherPattern `
                -ExpectedSha256 $spec.InstalledSha256
            if (-not $extractedEvidence.Valid) {
                throw "HPIA payload was rejected before installation: $($extractedEvidence.Reason)"
            }

            New-Item -ItemType Directory -Path $hpiaParent -Force -ErrorAction Stop | Out-Null
            New-Item -ItemType Directory -Path $stagingPath -Force -ErrorAction Stop | Out-Null
            Get-ChildItem -LiteralPath $extractedExecutable.Directory.FullName -Force -ErrorAction Stop |
                Copy-Item -Destination $stagingPath -Recurse -Force -ErrorAction Stop
            $stagedExecutable = Join-Path $stagingPath "HPImageAssistant.exe"
            $stagedEvidence = Get-InstalledExecutableEvidence -Path $stagedExecutable `
                -MinimumVersion $spec.InstalledMinimumVersion `
                -PublisherPattern $spec.InstalledPublisherPattern `
                -ExpectedSha256 $spec.InstalledSha256
            if (-not $stagedEvidence.Valid) {
                throw "HPIA staging copy failed payload verification: $($stagedEvidence.Reason)"
            }

            if (Test-Path -LiteralPath $hpiaDir) {
                Move-Item -LiteralPath $hpiaDir -Destination $backupPath -ErrorAction Stop
            }
            try {
                Move-Item -LiteralPath $stagingPath -Destination $hpiaDir -ErrorAction Stop
                $installedExecutable = Join-Path $hpiaDir "HPImageAssistant.exe"
                $installedEvidence = Get-InstalledExecutableEvidence -Path $installedExecutable `
                    -MinimumVersion $spec.InstalledMinimumVersion `
                    -PublisherPattern $spec.InstalledPublisherPattern `
                    -ExpectedSha256 $spec.InstalledSha256
                if (-not $installedEvidence.Valid) {
                    throw "Installed HPIA failed provenance verification: $($installedEvidence.Reason)"
                }
                $committed = $true
            } catch {
                Remove-Item -LiteralPath $hpiaDir -Recurse -Force -ErrorAction SilentlyContinue
                if (Test-Path -LiteralPath $backupPath) {
                    Move-Item -LiteralPath $backupPath -Destination $hpiaDir -ErrorAction SilentlyContinue
                }
                throw
            }

            Remove-Item -LiteralPath $backupPath -Recurse -Force -ErrorAction SilentlyContinue
            [void](Add-AcquisitionProvenance -Name $spec.Name -Version $installedEvidence.Version `
                -SourceUri $spec.Uri -Sha256 $installedEvidence.Sha256 `
                -Publisher $installedEvidence.Publisher -Thumbprint $installedEvidence.Thumbprint `
                -Architecture $architecture -InstallPath $installedEvidence.Path -Status "Installed" `
                -Evidence @(
                    "Outer package SHA-256 $($download.Sha256)",
                    "Extractor exit $($extractProcess.ExitCode)"
                ))
            Write-Log "HP Image Assistant $($installedEvidence.Version) installed from verified HP artifacts" "SUCCESS"
            return $true
        } finally {
            if (-not $committed) {
                Remove-Item -LiteralPath $stagingPath -Recurse -Force -ErrorAction SilentlyContinue
            }
            Remove-SystemUpdateProTemporaryDirectory -Path $temporaryDirectory
        }
    }
}

function Invoke-HPUpdate {
    param(
        [switch]$IncludeBIOS,
        [AllowNull()][object]$SystemInfo = $null,
        [AllowNull()][object]$FirmwarePrerequisites = $null
    )

    $result = @{
        Success = $false; Status = "Failed"; RebootRequired = $false
        UpdateCount = 0; Available = 0; Attempted = 0; Installed = 0; Failed = 0; Skipped = 0
        ExitCode = $null; HResult = $null; Items = @(); Evidence = @(); Message = ""
        FirmwareReadiness = $null
    }

    try {
        Write-Log "========== HP IMAGE ASSISTANT ==========" "HEADER"

        $sysInfo = if ($null -ne $SystemInfo) { $SystemInfo } else { Get-SystemInfo }
        Add-SensitiveEvidenceValue -Value ([string]$sysInfo.SerialNumber)
        Write-Log "Serial: $($sysInfo.SerialNumber)" "INFO"

        $hpiaPath = Get-HPIAPath
        if (-not $hpiaPath) {
            if (-not (Install-HPIA)) {
                $result.Message = "HP Image Assistant could not be installed; repair download access and rerun"
                $result.Failed = 1
            }
            $hpiaPath = Get-HPIAPath
        }

        if (-not $hpiaPath) {
            $readiness = Test-FirmwareReadiness -Requested:$IncludeBIOS -Provider "HP" -SystemInfo $sysInfo `
                -ToolState "Unknown" -ToolMessage "HP Image Assistant is unavailable. Install or repair HPIA and rerun its applicability scan" `
                -Prerequisites $FirmwarePrerequisites
            $result.FirmwareReadiness = $readiness
            $result.Items += New-FirmwareReadinessItem -Readiness $readiness
            $result.Skipped = $(if ($IncludeBIOS) { 1 } else { 0 })
            if ($DryRun) {
                $result.Success = $true
                $result.Status = $(if ($IncludeBIOS) { "Partial" } else { "Succeeded" })
                $result.Message = "HPIA would be installed before scanning; $($readiness.Message)"
            } else {
                $result.Failed = [math]::Max(1, $result.Failed)
                $result.Message = "HP Image Assistant was not found after installation"
            }
            Write-Log $result.Message $(if ($result.Success) { "WARNING" } else { "ERROR" })
            return $result
        }

        if (-not $DryRun) {
            Get-Process -Name "HPImageAssistant*" -ErrorAction SilentlyContinue |
                Stop-Process -Force -ErrorAction SilentlyContinue
        }

        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $reportDir = Join-Path $LogPath "HPIA_$timestamp"
        $softpaqDir = Join-Path $env:TEMP "HPSoftpaqs"
        if (-not (New-ProtectedDirectory -Path $reportDir)) {
            throw "HPIA report directory could not be protected"
        }
        if (-not $DryRun) {
            New-Item -ItemType Directory -Path $softpaqDir -Force | Out-Null
        }

        $scanCategories = if ($IncludeBIOS) { @("Drivers", "Firmware", "BIOS") } else { @("Drivers") }
        $scanArgs = @(
            "/Operation:Analyze", "/Action:List", "/Selection:All",
            "/Category:$($scanCategories -join ',')", "/Silent", "/Noninteractive",
            "/ReportFolder:`"$reportDir`""
        )
        Write-Log "Verifying HP model/tool applicability..." "INFO"
        $scanProcess = Start-Process -FilePath $hpiaPath -ArgumentList ($scanArgs -join " ") `
            -Wait -NoNewWindow -PassThru -ErrorAction Stop
        [void](Protect-EvidenceTree -Path $reportDir `
            -SensitiveValues @([string]$sysInfo.SerialNumber))
        $result.Attempted++
        $result.ExitCode = $scanProcess.ExitCode
        $result.Evidence = @($reportDir)
        $scanSucceeded = $scanProcess.ExitCode -in @(0, 256, 257, 3010)
        $toolState = if ($scanSucceeded) { "Ready" } elseif ($scanProcess.ExitCode -eq 4104) { "Blocked" } else { "Unknown" }
        $toolMessage = if ($scanSucceeded) {
            "HP Image Assistant completed an applicability scan for this model and operating system"
        } elseif ($scanProcess.ExitCode -eq 4104) {
            "HPIA reports that this model/OS lacks a supported reference image (exit 4104). Check HP's supported-platform list"
        } else {
            "HPIA scan exit $($scanProcess.ExitCode) did not verify support. Review $reportDir, repair HPIA, and rerun"
        }
        $readiness = Test-FirmwareReadiness -Requested:$IncludeBIOS -Provider "HP" -SystemInfo $sysInfo `
            -ToolState $toolState -ToolMessage $toolMessage -Prerequisites $FirmwarePrerequisites
        $result.FirmwareReadiness = $readiness
        $result.Items += New-FirmwareReadinessItem -Readiness $readiness
        $policy = Get-FirmwareUpdatePolicy -Readiness $readiness

        if (-not $scanSucceeded) {
            $result.Failed = 1
            $result.Message = "HPIA applicability scan failed with exit $($scanProcess.ExitCode); no updates were applied. $toolMessage"
            Write-Log $result.Message "ERROR"
            return $result
        }

        if ($IncludeBIOS -and -not $readiness.Safe) {
            $result.Skipped++
            $result.Status = "Partial"
            Write-Log $readiness.Message "WARNING"
        }

        if ($DryRun) {
            $result.Success = $true
            if ($result.Status -ne "Partial") { $result.Status = "Succeeded" }
            $result.Message = "HPIA applicability scan completed (exit: $($scanProcess.ExitCode)); no updates installed. $($readiness.Message)"
            Write-Log $result.Message $(if ($readiness.Safe -or -not $IncludeBIOS) { "INFO" } else { "WARNING" })
            $script:OEMUpdateCount = $result.UpdateCount
            return $result
        }

        if ($scanProcess.ExitCode -eq 256) {
            $result.Success = $true
            if ($result.Status -ne "Partial") { $result.Status = "Succeeded" }
            $result.Message = "No applicable HP updates were found; $($readiness.Message)"
            Write-Log $result.Message "SUCCESS"
            return $result
        }

        $hpiaArgs = @(
            "/Operation:Analyze", "/Action:Install", "/Selection:All",
            "/Category:$($policy.HPCategories -join ',')", "/Silent", "/Noninteractive",
            "/ReportFolder:`"$reportDir`"", "/SoftpaqDownloadFolder:`"$softpaqDir`""
        )
        Write-Log "Applying HP update categories: $($policy.HPCategories -join ', ')..." "INFO"
        $process = Start-Process -FilePath $hpiaPath -ArgumentList ($hpiaArgs -join " ") `
            -Wait -NoNewWindow -PassThru -ErrorAction Stop
        [void](Protect-EvidenceTree -Path $reportDir `
            -SensitiveValues @([string]$sysInfo.SerialNumber))
        $result.Attempted++
        $result.ExitCode = $process.ExitCode

        Get-Process -Name "HPImageAssistant*" -ErrorAction SilentlyContinue |
            Stop-Process -Force -ErrorAction SilentlyContinue

        switch ($process.ExitCode) {
            0    { $result.Success = $true; $result.Message = "Updates applied" }
            256  { $result.Success = $true; $result.Message = "No updates needed" }
            257  { $result.Success = $true; $result.RebootRequired = $true; $result.Message = "Updates applied - reboot required" }
            3010 { $result.Success = $true; $result.RebootRequired = $true; $result.Message = "Updates applied - reboot required" }
            default { $result.Success = $false; $result.Message = "HPIA exit $($process.ExitCode); review $reportDir and rerun" }
        }
        if ($result.Success) {
            if ($result.Status -ne "Partial") { $result.Status = "Succeeded" }
            $result.Items += New-UpdateItemResult -Name "HP update application" -Status "Succeeded" `
                -ProviderCode $result.ExitCode -RebootRequired $result.RebootRequired -Message $result.Message -Evidence @($reportDir)
        } else {
            $result.Status = "Failed"
            $result.Failed = 1
            $result.Items += New-UpdateItemResult -Name "HP update application" -Status "Failed" `
                -ProviderCode $result.ExitCode -Message $result.Message -Evidence @($reportDir)
        }

        Write-Log $result.Message $(if ($result.Success) { "SUCCESS" } else { "WARNING" })
        Remove-Item $softpaqDir -Recurse -Force -ErrorAction SilentlyContinue
        $script:OEMUpdateCount = $result.UpdateCount
        return $result
    } catch {
        $result.Success = $false
        $result.Status = "Failed"
        $result.Failed = [math]::Max(1, $result.Failed)
        $result.HResult = $_.Exception.HResult
        $result.Message = "HP update error: $($_.Exception.Message)"
        if ($null -eq $result.FirmwareReadiness) {
            $readiness = Test-FirmwareReadiness -Requested:$IncludeBIOS -Provider "HP" -SystemInfo $SystemInfo `
                -ToolState "Unknown" -ToolMessage "HPIA failed before support could be verified: $($_.Exception.Message). Repair HPIA and rerun" `
                -Prerequisites $FirmwarePrerequisites
            $result.FirmwareReadiness = $readiness
            $result.Items += New-FirmwareReadinessItem -Readiness $readiness
        }
        Write-Log $result.Message "ERROR"
        return $result
    }
}

# ============================================================================
# ADDITIONAL OEM AND GPU PROVIDERS
# ============================================================================

function Get-GPUProviderPlan {
    param(
        [AllowNull()][object[]]$VideoControllers = @()
    )

    if (@($VideoControllers).Count -eq 0) {
        try {
            $VideoControllers = @(Get-CimInstance -ClassName Win32_VideoController -ErrorAction SilentlyContinue)
        } catch {
            $VideoControllers = @()
        }
    }

    $plans = [System.Collections.ArrayList]::new()
    $seen = @{}
    foreach ($controller in @($VideoControllers)) {
        $name = [string](Get-ResultValue -Result $controller -Names @("Name", "Description", "AdapterCompatibility") -Default "")
        $vendor = if ($name -match "(?i)NVIDIA|GeForce|Quadro|RTX|Tesla") {
            "NVIDIA"
        } elseif ($name -match "(?i)AMD|Radeon|ATI") {
            "AMD"
        } elseif ($name -match "(?i)Intel|Iris|UHD|Arc") {
            "Intel"
        } else {
            ""
        }
        if ([string]::IsNullOrWhiteSpace($vendor) -or $seen.ContainsKey($vendor)) { continue }
        $seen[$vendor] = $true

        $candidatePaths = switch ($vendor) {
            "Intel" {
                @(
                    "${env:ProgramFiles}\Intel\Driver and Support Assistant\DSAService.exe",
                    "${env:ProgramFiles(x86)}\Intel\Driver and Support Assistant\DSAService.exe"
                )
            }
            "AMD" {
                @(
                    "${env:ProgramFiles}\AMD\CIM\Bin64\AMDInstallManager.exe",
                    "${env:ProgramFiles}\AMD\CNext\CNext\RadeonSoftware.exe"
                )
            }
            "NVIDIA" {
                @(
                    "${env:ProgramFiles}\NVIDIA Corporation\NVIDIA GeForce Experience\NVIDIA GeForce Experience.exe",
                    "${env:ProgramFiles}\NVIDIA Corporation\NVIDIA app\NVIDIA app.exe"
                )
            }
        }
        $source = switch ($vendor) {
            "Intel" { "https://www.intel.com/content/www/us/en/download-center/home.html" }
            "AMD" { "https://www.amd.com/en/support/download/drivers.html" }
            "NVIDIA" { "https://www.nvidia.com/Download/index.aspx" }
        }
        $arguments = switch ($vendor) {
            "Intel" { @("/update", "/silent") }
            "AMD" { @("--update", "--silent") }
            "NVIDIA" { @("--update", "--silent") }
        }
        $installedPath = @($candidatePaths | Where-Object {
            -not [string]::IsNullOrWhiteSpace([string]$_) -and (Test-Path -LiteralPath $_ -PathType Leaf)
        } | Select-Object -First 1)
        [void]$plans.Add([PSCustomObject][ordered]@{
            Provider = $vendor
            Category = "GPU"
            DisplayName = "$vendor GPU driver"
            SourceUri = $source
            CandidatePaths = @($candidatePaths)
            ExecutablePath = if ($installedPath.Count) { [string]$installedPath[0] } else { "" }
            Arguments = @($arguments)
            Status = if ($installedPath.Count) { "Ready" } else { "AcquisitionRequired" }
            Applicable = $true
            Reason = if ($installedPath.Count) {
                "A signed $vendor update client is installed and can be invoked with bounded arguments"
            } else {
                "$vendor GPU detected; use the vendor's public driver endpoint or install its signed update client"
            }
        })
    }
    return @($plans)
}

function Get-AdditionalOEMProviderPlan {
    param(
        [AllowNull()][System.Collections.IDictionary]$SystemInfo = $null,
        [switch]$IncludeGPU
    )

    if ($null -eq $SystemInfo) { $SystemInfo = $script:CurrentSystemInfo }
    if ($null -eq $SystemInfo) { $SystemInfo = Get-SystemInfo }
    $manufacturer = [string]$SystemInfo.Manufacturer
    $model = [string]$SystemInfo.Model
    $plans = [System.Collections.ArrayList]::new()

    $definitions = @(
        [ordered]@{
            Provider = "ASUS"; Pattern = "(?i)ASUSTeK|ASUS"
            DisplayName = "ASUS MyASUS / Armoury Crate"
            SourceUri = "https://www.asus.com/support/"
            CandidatePaths = @(
                "${env:ProgramFiles}\ASUS\ASUS Smart Display Control\AsusSoftwareManager.exe",
                "${env:ProgramFiles}\ASUS\Armoury Crate Service\ArmouryCrate.Service.exe",
                "${env:ProgramFiles(x86)}\ASUS\ASUS Live Update\LiveUpdate.exe"
            )
            Arguments = @("/update", "/silent")
        },
        [ordered]@{
            Provider = "Acer"; Pattern = "(?i)ACER"
            DisplayName = "Acer Care Center"
            SourceUri = "https://www.acer.com/us-en/support/drivers-and-manuals"
            CandidatePaths = @(
                "${env:ProgramFiles}\Acer\Care Center\CareCenter.exe",
                "${env:ProgramFiles(x86)}\Acer\Care Center\CareCenter.exe"
            )
            Arguments = @("/update", "/silent")
        },
        [ordered]@{
            Provider = "MSI"; Pattern = "(?i)MICRO-STAR|MSI"
            DisplayName = "MSI Center / Dragon Center"
            SourceUri = "https://www.msi.com/Landing/MSI-Center"
            CandidatePaths = @(
                "${env:ProgramFiles}\MSI\MSI Center\MSI.CentralServer.exe",
                "${env:ProgramFiles(x86)}\MSI\Dragon Center\Dragon Center.exe"
            )
            Arguments = @("/update", "/silent")
        },
        [ordered]@{
            Provider = "Surface"; Pattern = "(?i)MICROSOFT"
            ModelPattern = "(?i)SURFACE"
            DisplayName = "Microsoft Surface firmware and driver pack"
            SourceUri = "https://www.microsoft.com/download/details.aspx?id=100440"
            CandidatePaths = @(
                "${env:ProgramFiles}\Microsoft Surface\SurfaceUpdate.exe",
                "${env:ProgramFiles(x86)}\Microsoft Surface\SurfaceUpdate.exe"
            )
            Arguments = @("/update", "/silent", "/norestart")
        },
        [ordered]@{
            Provider = "Framework"; Pattern = "(?i)FRAMEWORK"
            DisplayName = "Framework laptop firmware"
            SourceUri = "https://knowledgebase.frame.work/en_us/categories/firmware-updates"
            CandidatePaths = @("${env:ProgramFiles}\Framework\fwupdmgr.exe", "${env:SystemRoot}\System32\fwupdmgr.exe")
            Arguments = @("update", "--assumeyes")
        },
        [ordered]@{
            Provider = "Panasonic"; Pattern = "(?i)PANASONIC|MATSUSHITA"
            DisplayName = "Panasonic Toughbook drivers"
            SourceUri = "https://na.panasonic.com/us/support"
            CandidatePaths = @(
                "${env:ProgramFiles}\Panasonic\Common Components\SetAutoUpdate.exe",
                "${env:ProgramFiles(x86)}\Panasonic\Common Components\SetAutoUpdate.exe"
            )
            Arguments = @("/update", "/silent")
        }
    )

    foreach ($definition in $definitions) {
        if ([string]$manufacturer -notmatch [string]$definition.Pattern) { continue }
        if ($definition.Contains("ModelPattern") -and [string]$model -notmatch [string]$definition.ModelPattern) { continue }
        $installedPath = @($definition.CandidatePaths | Where-Object {
            -not [string]::IsNullOrWhiteSpace([string]$_) -and (Test-Path -LiteralPath $_ -PathType Leaf)
        } | Select-Object -First 1)
        [void]$plans.Add([PSCustomObject][ordered]@{
            Provider = [string]$definition.Provider
            Category = "OEM"
            DisplayName = [string]$definition.DisplayName
            SourceUri = [string]$definition.SourceUri
            CandidatePaths = @($definition.CandidatePaths)
            ExecutablePath = if ($installedPath.Count) { [string]$installedPath[0] } else { "" }
            Arguments = @($definition.Arguments)
            Status = if ($installedPath.Count) { "Ready" } else { "AcquisitionRequired" }
            Applicable = $true
            Reason = if ($installedPath.Count) {
                "A signed vendor updater is installed and can be invoked with bounded arguments"
            } else {
                "The model matches this provider, but no approved local updater was found"
            }
        })
    }

    if ($IncludeGPU) {
        foreach ($gpuPlan in @(Get-GPUProviderPlan)) { [void]$plans.Add($gpuPlan) }
    }
    return @($plans)
}

function Invoke-AdditionalOEMUpdate {
    param(
        [AllowNull()][System.Collections.IDictionary]$SystemInfo = $null,
        [switch]$IncludeGPU
    )

    $result = @{
        Success = $true; Status = "Succeeded"; RebootRequired = $false
        UpdateCount = 0; Available = 0; Attempted = 0; Installed = 0; Failed = 0; Skipped = 0
        ExitCode = $null; HResult = $null; Items = @(); Evidence = @(); Message = ""
        Plans = @()
    }
    $plans = @(Get-AdditionalOEMProviderPlan -SystemInfo $SystemInfo -IncludeGPU:$IncludeGPU)
    $result.Plans = $plans
    if ($plans.Count -eq 0) {
        $result.Message = "No additional OEM or GPU providers apply to this system"
        return $result
    }

    foreach ($plan in $plans) {
        $status = [string]$plan.Status
        if ($status -eq "Ready" -and -not [string]::IsNullOrWhiteSpace([string]$plan.ExecutablePath)) {
            if ($DryRun) {
                $result.Available++
                $result.Items += New-UpdateItemResult -Name ([string]$plan.DisplayName) -Status "Available" `
                    -Message "Would invoke $($plan.ExecutablePath) $($plan.Arguments -join ' ')" `
                    -Evidence @("source:$($plan.SourceUri)")
                continue
            }
            try {
                $process = Start-Process -FilePath ([string]$plan.ExecutablePath) `
                    -ArgumentList @($plan.Arguments) -Wait -NoNewWindow -PassThru -ErrorAction Stop
                $result.Attempted++
                $result.ExitCode = $process.ExitCode
                if ($process.ExitCode -in @(0, 3010, 1641)) {
                    $result.Installed++
                    $result.UpdateCount++
                    if ($process.ExitCode -in @(3010, 1641)) { $result.RebootRequired = $true }
                    $result.Items += New-UpdateItemResult -Name ([string]$plan.DisplayName) -Status "Installed" `
                        -ProviderCode $process.ExitCode -RebootRequired ($process.ExitCode -in @(3010, 1641)) `
                        -Message "Vendor updater completed" -Evidence @("source:$($plan.SourceUri)")
                } else {
                    $result.Failed++
                    $result.Items += New-UpdateItemResult -Name ([string]$plan.DisplayName) -Status "Failed" `
                        -ProviderCode $process.ExitCode -Message "Vendor updater exited with $($process.ExitCode)"
                }
            } catch {
                $result.Attempted++
                $result.Failed++
                $result.Items += New-UpdateItemResult -Name ([string]$plan.DisplayName) -Status "Failed" `
                    -Message "Vendor updater failed: $($_.Exception.Message)"
            }
        } else {
            $result.Skipped++
            $result.Items += New-UpdateItemResult -Name ([string]$plan.DisplayName) -Status "Skipped" `
                -Message ([string]$plan.Reason) -Evidence @("source:$($plan.SourceUri)")
        }
    }

    $result.Success = ($result.Failed -eq 0)
    $result.Status = if (-not $result.Success -and $result.Installed -gt 0) { "Partial" } elseif (-not $result.Success) { "Failed" } else { "Succeeded" }
    $result.Message = if ($DryRun) {
        "$($result.Available) additional OEM/GPU updater(s) are available (dry run)"
    } else {
        "Installed: $($result.Installed), Failed: $($result.Failed), Skipped: $($result.Skipped)"
    }
    return $result
}

# ============================================================================
# WINDOWS UPDATE
# ============================================================================

function Get-WindowsUpdatePolicy {
    param(
        [string]$Path = [string]$script:PolicyPath,
        [int]$FeatureDeferralDays = [int]$script:FeatureDeferralDays,
        [bool]$SecurityOnly = [bool]$script:SecurityOnly,
        [bool]$PreStage = [bool]$script:PreStage
    )

    $policy = [ordered]@{
        SchemaVersion = 1
        FeatureDeferralDays = [math]::Max(0, [math]::Min(3650, $FeatureDeferralDays))
        SecurityOnly = $SecurityOnly
        PreStage = $PreStage
        AllowFeatureUpdates = $false
        DriverAllow = @()
        DriverDeny = @()
        CatalogFallback = $true
        ADMXSnapshot = $true
        Sources = @()
        Errors = @()
    }
    try {
        $document = Get-PolicyDocument -Path $Path
        if ($document.Contains("windows_update") -and $document.windows_update -is [System.Collections.IDictionary]) {
            $document = $document.windows_update
        }
        if ($document.Contains("feature_deferral_days")) {
            $parsedDays = 0
            if ([int]::TryParse([string]$document.feature_deferral_days, [ref]$parsedDays)) {
                $policy.FeatureDeferralDays = [math]::Max(0, [math]::Min(3650, $parsedDays))
            }
        }
        if ($document.Contains("security_only")) { $policy.SecurityOnly = [bool]$document.security_only }
        if ($document.Contains("pre_stage")) { $policy.PreStage = [bool]$document.pre_stage }
        if ($document.Contains("allow_feature_updates")) { $policy.AllowFeatureUpdates = [bool]$document.allow_feature_updates }
        if ($document.Contains("driver_allow")) { $policy.DriverAllow = @($document.driver_allow | ForEach-Object { [string]$_ } | Where-Object { $_ }) }
        if ($document.Contains("driver_deny")) { $policy.DriverDeny = @($document.driver_deny | ForEach-Object { [string]$_ } | Where-Object { $_ }) }
        if ($document.Contains("catalog_fallback")) { $policy.CatalogFallback = [bool]$document.catalog_fallback }
        if ($document.Contains("admx_snapshot")) { $policy.ADMXSnapshot = [bool]$document.admx_snapshot }
        if (-not [string]::IsNullOrWhiteSpace($Path)) { $policy.Sources += "policy:$Path" }
    } catch {
        $policy.Errors += Protect-EvidenceText -Text $_.Exception.Message
    }
    if ([int]$policy.FeatureDeferralDays -gt 0) { $policy.AllowFeatureUpdates = $true }
    $policy.DriverAllow = @($policy.DriverAllow | Select-Object -Unique)
    $policy.DriverDeny = @($policy.DriverDeny | Select-Object -Unique)
    return [PSCustomObject]$policy
}

function Get-WindowsUpdateItemIdentity {
    param([Parameter(Mandatory = $true)][object]$Item)

    $id = [string](Get-ResultValue -Result $Item -Names @("UpdateID", "UpdateId", "Id", "Identity") -Default "")
    if ($Item.Identity) { $id = [string](Get-ResultValue -Result $Item.Identity -Names @("UpdateID", "UpdateId") -Default $id) }
    $kb = Get-ResultValue -Result $Item -Names @("KB", "KBArticleIDs", "ArticleID", "ArticleIds") -Default ""
    $kbText = (@($kb) | ForEach-Object { [string]$_ }) -join ","
    $title = [string](Get-ResultValue -Result $Item -Names @("Title", "Name") -Default $kbText)
    if ([string]::IsNullOrWhiteSpace($id)) { $id = $kbText }
    if ([string]::IsNullOrWhiteSpace($id)) { $id = $title }
    return [PSCustomObject][ordered]@{ Id = $id; KB = $kbText; Title = $title }
}

function Test-WindowsUpdateItemPolicy {
    param(
        [Parameter(Mandatory = $true)][object]$Item,
        [AllowNull()][object]$Policy = $null
    )

    if ($null -eq $Policy) { $Policy = Get-WindowsUpdatePolicy }
    $identity = Get-WindowsUpdateItemIdentity -Item $Item
    $categories = @($Item.Categories | ForEach-Object {
        [string](Get-ResultValue -Result $_ -Names @("Name", "Title") -Default $_)
    })
    $categoryText = $categories -join "; "
    $title = [string]$identity.Title
    $isDriver = ([bool](Get-ResultValue -Result $Item -Names @("IsDriver") -Default $false)) -or
        $categoryText -match "(?i)driver" -or $title -match "(?i)driver"
    $isFeature = ([bool](Get-ResultValue -Result $Item -Names @("IsFeature") -Default $false)) -or
        $categoryText -match "(?i)feature\s*packs?" -or $title -match "(?i)feature update|upgrade to windows|enablement package"
    $isSecurity = ([bool](Get-ResultValue -Result $Item -Names @("IsSecurity", "Security") -Default $false)) -or
        $categoryText -match "(?i)security|critical" -or $title -match "(?i)security|critical|defender|malicious software removal"
    $releaseValue = Get-ResultValue -Result $Item -Names @("ReleaseDate", "Date", "LastDeploymentChangeTime") -Default $null
    $releaseDate = [datetime]::MinValue
    $hasReleaseDate = $null -ne $releaseValue -and [datetime]::TryParse([string]$releaseValue, [ref]$releaseDate)

    $matchesPattern = {
        param([object[]]$Patterns)
        foreach ($pattern in @($Patterns)) {
            if ([string]::IsNullOrWhiteSpace([string]$pattern)) { continue }
            if ($identity.Id -like [string]$pattern -or $identity.KB -like [string]$pattern -or
                $title -like [string]$pattern -or $categoryText -like [string]$pattern) { return $true }
        }
        return $false
    }

    $status = "Allowed"
    $reason = "Update is allowed by the Windows Update policy"
    if ($title -match "(?i)preview") {
        $status = "Blocked"
        $reason = "Preview updates are excluded by the enterprise-safe policy"
    } elseif ($Policy.SecurityOnly -and -not $isSecurity) {
        $status = "Blocked"
        $reason = "Security-only mode excludes non-critical and non-security updates"
    } elseif ($isDriver) {
        $denied = & $matchesPattern @($Policy.DriverDeny)
        $allowed = @($Policy.DriverAllow).Count -gt 0 -and (& $matchesPattern @($Policy.DriverAllow))
        if ($denied) {
            $status = "Blocked"
            $reason = "Driver matched the administrator driver deny list"
        } elseif (-not $allowed) {
            $status = "Blocked"
            $reason = "Drivers remain blocked unless matched by the administrator driver allow list"
        }
    } elseif ($isFeature -and -not [bool]$Policy.AllowFeatureUpdates) {
        $status = "Blocked"
        $reason = "Feature updates are disabled by the default enterprise-safe policy"
    } elseif ($isFeature -and [int]$Policy.FeatureDeferralDays -gt 0) {
        if (-not $hasReleaseDate) {
            $status = "Deferred"
            $reason = "Feature update release age is unknown; it is deferred fail-closed"
        } elseif (((Get-Date) - $releaseDate).TotalDays -lt [int]$Policy.FeatureDeferralDays) {
            $status = "Deferred"
            $reason = "Feature update is younger than the configured $($Policy.FeatureDeferralDays)-day deferral"
        }
    }

    return [PSCustomObject][ordered]@{
        Id = [string]$identity.Id; KB = [string]$identity.KB; Title = $title
        Categories = @($categories); IsDriver = $isDriver; IsFeature = $isFeature; IsSecurity = $isSecurity
        ReleaseDate = if ($hasReleaseDate) { $releaseDate.ToString("o") } else { "" }
        Status = $status; Allowed = ($status -eq "Allowed"); Deferred = ($status -eq "Deferred")
        Reason = $reason
    }
}

function Get-WindowsUpdatePlan {
    param(
        [AllowEmptyCollection()][object[]]$Updates = @(),
        [AllowNull()][object]$Policy = $null
    )

    if ($null -eq $Policy) { $Policy = Get-WindowsUpdatePolicy }
    $allowed = [System.Collections.ArrayList]::new()
    $items = [System.Collections.ArrayList]::new()
    foreach ($update in @($Updates)) {
        $decision = Test-WindowsUpdateItemPolicy -Item $update -Policy $Policy
        [void]$items.Add([PSCustomObject][ordered]@{
            Update = $update; Decision = $decision
        })
        if ($decision.Allowed) { [void]$allowed.Add($update) }
    }
    return [PSCustomObject][ordered]@{
        SchemaVersion = 1
        Policy = $Policy
        Items = @($items)
        AllowedUpdates = @($allowed)
        Blocked = @($items | Where-Object { $_.Decision.Status -eq "Blocked" })
        Deferred = @($items | Where-Object { $_.Decision.Status -eq "Deferred" })
    }
}

function Test-WindowsUpdatePrestageDocument {
    param([AllowNull()][object]$Document)
    if ($Document -isnot [System.Collections.IDictionary] -or [int]$Document.SchemaVersion -ne 1) {
        return [PSCustomObject]@{ Valid = $false; Reason = "Windows Update pre-stage schema is invalid" }
    }
    if ($Document.Items -isnot [System.Collections.IEnumerable]) {
        return [PSCustomObject]@{ Valid = $false; Reason = "Windows Update pre-stage items are missing" }
    }
    return [PSCustomObject]@{ Valid = $true; Reason = "" }
}

function Get-WindowsUpdatePrestagePath {
    return Join-Path $script:DataPath "WindowsUpdatePrestage.json"
}

function Get-WindowsUpdatePrestagePlan {
    $path = Get-WindowsUpdatePrestagePath
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        return [PSCustomObject]@{ Exists = $false; Valid = $true; Path = $path; Items = @(); Reason = "No pre-stage plan exists" }
    }
    $read = Read-ProtectedJsonFile -Path $path -ValidationScript ${function:Test-WindowsUpdatePrestageDocument}
    if (-not $read.Success) {
        return [PSCustomObject]@{ Exists = $true; Valid = $false; Path = $path; Items = @(); Reason = [string]$read.Error }
    }
    return [PSCustomObject]@{ Exists = $true; Valid = $true; Path = $path; Items = @($read.Data.Items); Reason = "Pre-stage plan loaded" }
}

function Save-WindowsUpdatePrestagePlan {
    param(
        [Parameter(Mandatory = $true)][object[]]$Items,
        [bool]$DryRunMode = [bool]$script:DryRun
    )

    $path = Get-WindowsUpdatePrestagePath
    $document = [ordered]@{
        SchemaVersion = 1
        RunId = [string]$script:RunId
        CreatedAt = (Get-Date).ToUniversalTime().ToString("o")
        Items = @($Items)
    }
    if ($DryRunMode) {
        return [PSCustomObject]@{ Success = $true; Persisted = $false; Path = $path; Reason = "Dry run did not persist a pre-stage plan" }
    }
    try {
        if (-not (Write-ProtectedAtomicJson -Path $path -Data $document -Depth 20 `
            -DataValidationScript ${function:Test-WindowsUpdatePrestageDocument})) {
            throw $script:LastEvidenceWriteError
        }
        return [PSCustomObject]@{ Success = $true; Persisted = $true; Path = $path; Reason = "Pre-stage plan persisted atomically" }
    } catch {
        return [PSCustomObject]@{ Success = $false; Persisted = $false; Path = $path; Reason = "Pre-stage plan could not be persisted: $($_.Exception.Message)" }
    }
}

function Clear-WindowsUpdatePrestagePlan {
    $path = Get-WindowsUpdatePrestagePath
    if ($script:DryRun) { return $true }
    try {
        Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath "$path.previous" -Force -ErrorAction SilentlyContinue
        return -not (Test-Path -LiteralPath $path)
    } catch { return $false }
}

function Get-WindowsPolicySnapshot {
    param(
        [string[]]$Paths = @(
            "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate",
            "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"
        )
    )

    $entries = [System.Collections.ArrayList]::new()
    $missing = [System.Collections.ArrayList]::new()
    foreach ($path in @($Paths)) {
        $keys = @()
        try { $keys = @(Get-ChildItem -LiteralPath $path -Recurse -ErrorAction SilentlyContinue) } catch { $keys = @() }
        if (Test-Path -LiteralPath $path) { $keys = @((Get-Item -LiteralPath $path -ErrorAction SilentlyContinue)) + $keys }
        $keys = @($keys | Where-Object { $null -ne $_ } | Sort-Object PSPath -Unique)
        if ($keys.Count -eq 0) { [void]$missing.Add($path); continue }
        foreach ($key in $keys) {
            try {
                foreach ($name in @($key.Property)) {
                    $snapshot = Get-RegistryValueSnapshot -Path ([string]$key.PSPath) -Name ([string]$name)
                    [void]$entries.Add($snapshot)
                }
            } catch {
                [void]$missing.Add("${path}: $($_.Exception.Message)")
            }
        }
    }
    return [PSCustomObject][ordered]@{
        SchemaVersion = 1; CapturedAt = (Get-Date).ToUniversalTime().ToString("o")
        Paths = @($Paths); Entries = @($entries); Missing = @($missing)
        Status = if ($missing.Count -eq 0) { "Ready" } else { "Partial" }
    }
}

function Save-WindowsPolicySnapshot {
    param(
        [AllowNull()][object]$Snapshot = $null,
        [bool]$DryRunMode = [bool]$script:DryRun
    )

    if ($null -eq $Snapshot) { $Snapshot = Get-WindowsPolicySnapshot }
    $path = Join-Path $script:DataPath "WindowsPolicySnapshot.json"
    if ($DryRunMode) {
        return [PSCustomObject]@{ Success = $true; Persisted = $false; Path = $path; Snapshot = $Snapshot; Reason = "Dry run retained the ADMX policy snapshot in memory only" }
    }
    try {
        if (-not (Write-ProtectedAtomicJson -Path $path -Data $Snapshot -Depth 30)) {
            throw $script:LastEvidenceWriteError
        }
        return [PSCustomObject]@{ Success = $true; Persisted = $true; Path = $path; Snapshot = $Snapshot; Reason = "Windows Update ADMX snapshot persisted atomically" }
    } catch {
        return [PSCustomObject]@{ Success = $false; Persisted = $false; Path = $path; Snapshot = $Snapshot; Reason = "Windows Update ADMX snapshot failed: $($_.Exception.Message)" }
    }
}

function Invoke-WindowsPolicySnapshot {
    param([bool]$DryRunMode = [bool]$script:DryRun)
    $snapshot = Get-WindowsPolicySnapshot
    return Save-WindowsPolicySnapshot -Snapshot $snapshot -DryRunMode $DryRunMode
}

function Invoke-WindowsUpdateCatalogFallback {
    param(
        [AllowEmptyCollection()][object[]]$Items = @(),
        [string]$Reason = "Windows Update providers failed"
    )

    $result = @{
        Success = $false; Status = "Failed"; RebootRequired = $false
        Available = 0; Attempted = 0; Installed = 0; Failed = 0; Skipped = 0
        ExitCode = $null; HResult = $null; Items = @(); Evidence = @(); Message = ""
    }
    $identities = @($Items | Where-Object { $null -ne $_ } | ForEach-Object { Get-WindowsUpdateItemIdentity -Item $_ } | Where-Object {
        $_.KB -match "(?i)KB\d+" -or $_.Title -match "(?i)KB\d+"
    })
    if ($identities.Count -eq 0) {
        $result.Failed = 1
        $result.Message = "Catalog fallback could not identify a KB from the failed provider result"
        return $result
    }
    if (-not (Test-DownloadAllowed)) {
        $result.Success = $true; $result.Status = "Succeeded"; $result.Skipped = $identities.Count
        $result.Message = "Catalog fallback deferred by the provider download policy"
        foreach ($identity in $identities) {
            $result.Items += New-UpdateItemResult -Name $identity.Title -Id $identity.KB -Status "Skipped" -Message $result.Message
        }
        return $result
    }

    $seen = @{}
    foreach ($identity in $identities) {
        $kbMatches = [regex]::Matches("$($identity.KB) $($identity.Title)", "(?i)KB\d+") | ForEach-Object { $_.Value.ToUpperInvariant() }
        foreach ($kb in @($kbMatches | Select-Object -Unique)) {
            if ($seen.ContainsKey($kb)) { continue }
            $seen[$kb] = $true
            $queryUri = "https://www.catalog.update.microsoft.com/Search.aspx?q=$kb"
            try {
                $page = Invoke-WebRequest -Uri $queryUri -UseBasicParsing -TimeoutSec ([int]$script:SourceTimeoutSeconds) -ErrorAction Stop
                $content = [string](Get-ResultValue -Result $page -Names @("Content", "RawContent") -Default "")
                $directUri = @([regex]::Matches($content, 'https://catalog\.s\.download\.windowsupdate\.com/[^"''<>\s]+\.(?:msu|cab)', "IgnoreCase") | ForEach-Object { $_.Value } | Select-Object -First 1)
                if ($script:DryRun) {
                    $result.Available++
                    $result.Items += New-UpdateItemResult -Name $identity.Title -Id $kb -Status "Available" `
                        -Message "Catalog checked; no package was downloaded in dry-run mode" -Evidence @($queryUri)
                    continue
                }
                if ($directUri.Count -eq 0 -or -not (Test-AcquisitionUri -Uri $directUri[0] -AllowedHosts @("catalog.s.download.windowsupdate.com"))) {
                    $result.Skipped++
                    $result.Items += New-UpdateItemResult -Name $identity.Title -Id $kb -Status "Skipped" `
                        -Message "Catalog search completed but no approved Microsoft package URL was exposed" -Evidence @($queryUri)
                    continue
                }
                $temporaryDirectory = New-SystemUpdateProTemporaryDirectory -Purpose "Catalog"
                $packagePath = Join-Path $temporaryDirectory "$kb.msu"
                try {
                    Invoke-WebRequest -Uri $directUri[0] -OutFile $packagePath -UseBasicParsing `
                        -TimeoutSec ([int]$script:SourceTimeoutSeconds) -ErrorAction Stop
                    $process = Start-Process -FilePath (Join-Path $script:WindowsRoot "System32\wusa.exe") `
                        -ArgumentList @($packagePath, "/quiet", "/norestart") -Wait -NoNewWindow -PassThru -ErrorAction Stop
                    $result.Attempted++
                    $result.ExitCode = $process.ExitCode
                    if ($process.ExitCode -in @(0, 3010, 1641)) {
                        $result.Installed++
                        $result.Success = $true
                        if ($process.ExitCode -in @(3010, 1641)) { $result.RebootRequired = $true }
                        $result.Items += New-UpdateItemResult -Name $identity.Title -Id $kb -Status "Installed" `
                            -ProviderCode $process.ExitCode -RebootRequired ($process.ExitCode -in @(3010, 1641)) `
                            -Message "Installed from Microsoft Update Catalog" -Evidence @($queryUri, $directUri[0])
                    } else {
                        $result.Failed++
                        $result.Items += New-UpdateItemResult -Name $identity.Title -Id $kb -Status "Failed" `
                            -ProviderCode $process.ExitCode -Message "Catalog package installer exited with $($process.ExitCode)"
                    }
                } finally { Remove-SystemUpdateProTemporaryDirectory -Path $temporaryDirectory }
            } catch {
                $result.Failed++
                $result.Items += New-UpdateItemResult -Name $identity.Title -Id $kb -Status "Failed" `
                    -Message "Catalog fallback failed: $($_.Exception.Message)" -Evidence @($queryUri)
            }
        }
    }
    if ($script:DryRun) { $result.Success = ($result.Failed -eq 0); $result.Status = "Succeeded" }
    elseif ($result.Failed -eq 0 -and $result.Installed -gt 0) { $result.Success = $true; $result.Status = "Succeeded" }
    elseif ($result.Installed -gt 0) { $result.Success = $false; $result.Status = "Partial" }
    $result.Message = if ($script:DryRun) {
        "Catalog fallback identified $($result.Available) package(s) after provider failure"
    } else {
        "Catalog fallback: installed $($result.Installed), failed $($result.Failed), skipped $($result.Skipped)"
    }
    return $result
}

function Invoke-WindowsUpdateWUA {
    $result = @{
        Success = $false; RebootRequired = $false
        Available = 0; Attempted = 0; Installed = 0; Failed = 0; Skipped = 0
        ExitCode = $null; HResult = $null; Items = @(); Evidence = @(); Message = ""
        CandidateUpdates = @(); Policy = $null; PreStage = $null
    }

    try {
        $policy = Get-WindowsUpdatePolicy
        $result.Policy = $policy
        $session = New-Object -ComObject Microsoft.Update.Session
        $searcher = $session.CreateUpdateSearcher()

        $searchResult = $searcher.Search("IsInstalled=0 and Type='Software' and IsHidden=0")

        if ($searchResult.Updates.Count -eq 0) {
            $result.Success = $true
            $result.Message = "No updates available"
            return $result
        }

        $allUpdates = @($searchResult.Updates)
        $result.CandidateUpdates = @($allUpdates)
        $plan = Get-WindowsUpdatePlan -Updates $allUpdates -Policy $policy
        foreach ($planned in @($plan.Items | Where-Object { -not $_.Decision.Allowed })) {
            $result.Skipped++
            $result.Items += New-UpdateItemResult -Name $planned.Decision.Title -Id $planned.Decision.Id -Status "Skipped" `
                -Message $planned.Decision.Reason -Evidence @("policy-status:$($planned.Decision.Status)")
        }

        $allowedUpdates = @($plan.AllowedUpdates)
        $preStagePlan = Get-WindowsUpdatePrestagePlan
        if ($preStagePlan.Exists -and $preStagePlan.Valid -and -not $policy.PreStage) {
            $pendingIds = @($preStagePlan.Items | ForEach-Object { [string](Get-ResultValue -Result $_ -Names @("Id", "UpdateID") -Default "") })
            if ($pendingIds.Count -gt 0) {
                $allowedUpdates = @($allowedUpdates | Where-Object {
                    $identity = Get-WindowsUpdateItemIdentity -Item $_
                    $pendingIds -contains [string]$identity.Id
                })
                $result.Evidence += $preStagePlan.Path
            }
        }

        $updatesToInstall = New-Object -ComObject Microsoft.Update.UpdateColl
        foreach ($update in @($allowedUpdates)) { [void]$updatesToInstall.Add($update) }

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
                $result.Items += New-UpdateItemResult -Name $uTitle -Id $update.Identity.UpdateID -Status "Available" `
                    -Message $(if ($policy.PreStage) { "Would download and persist a pre-stage plan" } else { "Policy-approved update" })
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

        if ($policy.PreStage) {
            $stageItems = @($allowedUpdates | ForEach-Object {
                $identity = Get-WindowsUpdateItemIdentity -Item $_
                [ordered]@{ Id = $identity.Id; KB = $identity.KB; Title = $identity.Title; Provider = "WUA" }
            })
            $saved = Save-WindowsUpdatePrestagePlan -Items $stageItems
            $result.PreStage = $saved
            if (-not $saved.Success) {
                $result.Failed = $updatesToInstall.Count
                $result.Message = $saved.Reason
                return $result
            }
            $result.Attempted = $updatesToInstall.Count
            $result.Success = $true
            foreach ($update in @($allowedUpdates)) {
                $identity = Get-WindowsUpdateItemIdentity -Item $update
                $result.Items += New-UpdateItemResult -Name $identity.Title -Id $identity.Id -Status "Available" `
                    -Message "Downloaded and staged for a later install window"
            }
            $result.Message = "Downloaded and staged: $($updatesToInstall.Count)"
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
        if ($result.Success -and $preStagePlan.Exists) { [void](Clear-WindowsUpdatePrestagePlan) }
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
        CandidateUpdates = @(); Policy = $null; PreStage = $null
    }

    try {
        $policy = Get-WindowsUpdatePolicy
        $result.Policy = $policy
        $psWindowsUpdatePath = Get-VerifiedPSModulePath -ModuleName "PSWindowsUpdate"
        if (-not $psWindowsUpdatePath) {
            throw "The pinned PSWindowsUpdate module failed installed-payload verification"
        }
        Import-Module $psWindowsUpdatePath -Force -ErrorAction Stop

        $updates = @(Get-WindowsUpdate -MicrosoftUpdate -ErrorAction Stop)

        if (-not $updates -or $updates.Count -eq 0) {
            $result.Success = $true
            $result.Message = "No updates available"
            return $result
        }

        $result.CandidateUpdates = @($updates)
        $plan = Get-WindowsUpdatePlan -Updates $updates -Policy $policy
        foreach ($planned in @($plan.Items | Where-Object { -not $_.Decision.Allowed })) {
            $result.Skipped++
            $result.Items += New-UpdateItemResult -Name $planned.Decision.Title -Id $planned.Decision.Id -Status "Skipped" `
                -Message $planned.Decision.Reason -Evidence @("policy-status:$($planned.Decision.Status)")
        }
        $allowedUpdates = @($plan.AllowedUpdates)
        $result.Available = $allowedUpdates.Count

        if ($DryRun) {
            $result.Success = ($result.Available -ge 0)
            $result.Message = "$($result.Available) updates available (dry run - not installed)"
            foreach ($u in $allowedUpdates) {
                $uTitle = if ($u.Title) { $u.Title } else { $u.KB }
                Write-Log "  -- $uTitle" "INFO"
                [void]$script:WindowsUpdates.Add($uTitle)
                $identity = Get-WindowsUpdateItemIdentity -Item $u
                $result.Items += New-UpdateItemResult -Name $uTitle -Id $identity.Id -Status "Available" `
                    -Message $(if ($policy.PreStage) { "Would download and persist a pre-stage plan" } else { "Policy-approved update" })
            }
            return $result
        }

        if ($allowedUpdates.Count -eq 0) {
            $result.Success = $true
            $result.Message = "No updates remain after Windows Update policy filtering"
            return $result
        }

        if ($policy.PreStage) {
            $downloadResults = @(Get-WindowsUpdate -Update $allowedUpdates -Download -AcceptAll -IgnoreReboot -Confirm:$false -ErrorAction Stop)
            $stageItems = @($allowedUpdates | ForEach-Object {
                $identity = Get-WindowsUpdateItemIdentity -Item $_
                [ordered]@{ Id = $identity.Id; KB = $identity.KB; Title = $identity.Title; Provider = "PSWindowsUpdate" }
            })
            $saved = Save-WindowsUpdatePrestagePlan -Items $stageItems
            $result.PreStage = $saved
            if (-not $saved.Success) {
                $result.Failed = $allowedUpdates.Count
                $result.Message = $saved.Reason
                return $result
            }
            $result.Attempted = $allowedUpdates.Count
            $result.Success = $true
            $result.Message = "Downloaded and staged: $($allowedUpdates.Count)"
            foreach ($u in $allowedUpdates) {
                $identity = Get-WindowsUpdateItemIdentity -Item $u
                $result.Items += New-UpdateItemResult -Name $identity.Title -Id $identity.Id -Status "Available" `
                    -Message "Downloaded and staged for a later install window"
            }
            return $result
        }

        $preStagePlan = Get-WindowsUpdatePrestagePlan
        if ($preStagePlan.Exists -and $preStagePlan.Valid) {
            $pendingIds = @($preStagePlan.Items | ForEach-Object { [string](Get-ResultValue -Result $_ -Names @("Id", "UpdateID") -Default "") })
            $allowedUpdates = @($allowedUpdates | Where-Object {
                $identity = Get-WindowsUpdateItemIdentity -Item $_
                $pendingIds -contains [string]$identity.Id
            })
            $result.Available = $allowedUpdates.Count
            if ($allowedUpdates.Count -eq 0) {
                $result.Success = $true
                $result.Message = "No updates from the persisted pre-stage plan remain applicable"
                return $result
            }
            $result.Evidence += $preStagePlan.Path
        }

        $result.Attempted = $allowedUpdates.Count
        $installResults = @(Install-WindowsUpdate -Update $allowedUpdates -MicrosoftUpdate -AcceptAll -IgnoreReboot -Confirm:$false -ErrorAction Stop)

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
            foreach ($u in $allowedUpdates) {
                $itemName = if ($u.Title) { $u.Title } else { [string]($u.KB -join ",") }
                $result.Items += New-UpdateItemResult -Name $itemName -Id ([string]($u.KB -join ",")) `
                    -Status "Failed" -Message "PSWindowsUpdate returned no installation result"
            }
        }

        $installationItems = @($result.Items | Where-Object { $_.Status -in @("Installed", "Failed") })
        $result.Success = ($result.Failed -eq 0 -and $installationItems.Count -eq $result.Attempted)
        $result.RebootRequired = (Get-WURebootStatus -Silent -ErrorAction SilentlyContinue)
        if ($result.Success -and $preStagePlan.Exists) { [void](Clear-WindowsUpdatePrestagePlan) }
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
    $policy = Get-WindowsUpdatePolicy
    $script:WindowsUpdatePolicy = $policy
    $providerSucceeded = $true

    for ($pass = 1; $pass -le $MaxPasses; $pass++) {
        Write-Log "Pass $pass of $MaxPasses" "INFO"
        $result.Passes = $pass

        $providerNames = if ($usePSWU) { @("PSWindowsUpdate", "WUA") } else { @("WUA") }
        $passResult = $null
        $providerFailureMessages = [System.Collections.ArrayList]::new()
        foreach ($providerName in $providerNames) {
            $candidateResult = if ($providerName -eq "PSWindowsUpdate") {
                Invoke-WindowsUpdatePSWU
            } else {
                Invoke-WindowsUpdateWUA
            }
            if ($candidateResult.Success) {
                $passResult = $candidateResult
                break
            }
            $passResult = $candidateResult
            [void]$providerFailureMessages.Add("${providerName}: $($candidateResult.Message)")
            if ($providerName -eq "PSWindowsUpdate") {
                $hasCandidateContract = ($candidateResult -is [System.Collections.IDictionary] -and $candidateResult.Contains("CandidateUpdates")) -or
                    ($candidateResult.PSObject.Properties["CandidateUpdates"] -and $null -ne $candidateResult.CandidateUpdates)
                if (-not $hasCandidateContract) { break }
            }
        }

        if ($null -eq $passResult) {
            $passResult = @{ Success = $false; Available = 0; Attempted = 0; Installed = 0; Failed = 1; Skipped = 0; Items = @(); Evidence = @(); Message = "No Windows Update provider returned a result" }
        }
        if (-not $passResult.Success -and [bool]$policy.CatalogFallback) {
            $fallbackResult = Invoke-WindowsUpdateCatalogFallback -Items @($passResult.CandidateUpdates) `
                -Reason (@($providerFailureMessages) -join "; ")
            if ($fallbackResult.Success) {
                $passResult = $fallbackResult
                $passResult.Message = "$($fallbackResult.Message); provider failures: $(@($providerFailureMessages) -join '; ')"
            } else {
                $passResult.Items = @($passResult.Items) + @($fallbackResult.Items)
                $passResult.Failed = [int]$passResult.Failed + [int]$fallbackResult.Failed
                $passResult.Message = "Provider failures: $(@($providerFailureMessages) -join '; '); $($fallbackResult.Message)"
            }
        }

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

    $dependencies = @($RunData.Dependencies)
    $mutationEvents = @($RunData.MutationRecovery)
    $mutationRecoveryCount = @($mutationEvents | Where-Object {
        [string]$_.Action -in @("Restored", "Committed")
    }).Count
    $retention = Get-ResultValue -Result $RunData -Names @("Retention") -Default $null
    $retentionDisplay = if ($null -eq $retention) {
        "Not evaluated"
    } else {
        $retainedDeletedFiles = [int](Get-ResultValue -Result $retention `
            -Names @("DeletedFiles") -Default 0)
        $retainedDeletedDirectories = [int](Get-ResultValue -Result $retention `
            -Names @("DeletedDirectories") -Default 0)
        $retainedBytes = [long](Get-ResultValue -Result $retention `
            -Names @("BytesFreed") -Default 0)
        "$retainedDeletedFiles files · $retainedDeletedDirectories directories · $retainedBytes bytes freed"
    }
    $capabilities = Get-ResultValue -Result $RunData -Names @("Capabilities") -Default $null
    $platformCapabilityDisplay = "Not assessed"
    $providerCapabilityDisplay = "Not assessed"
    $runtimeCapabilityDisplay = "Not assessed"
    if ($null -ne $capabilities) {
        $platformCapability = Get-ResultValue -Result $capabilities -Names @("Platform") -Default $null
        if ($null -ne $platformCapability) {
            $platformCapabilityDisplay = "{0} · {1} build {2} · {3} · {4}" -f @(
                (& $displayValue (Get-ResultValue -Result $platformCapability -Names @("Status") -Default "Unknown")),
                (& $displayValue (Get-ResultValue -Result $platformCapability -Names @("InstallationType") -Default "Unknown installation")),
                (& $displayValue (Get-ResultValue -Result $platformCapability -Names @("OSBuild") -Default "Unknown")),
                (& $displayValue (Get-ResultValue -Result $platformCapability -Names @("Architecture") -Default "Unknown architecture")),
                (& $displayValue (Get-ResultValue -Result $platformCapability -Names @("ExecutionContext") -Default "Unknown context"))
            )
            $runtimeCapabilityDisplay = "{0} {1}" -f @(
                (& $displayValue (Get-ResultValue -Result $platformCapability -Names @("PowerShellEdition") -Default "PowerShell")),
                (& $displayValue (Get-ResultValue -Result $platformCapability -Names @("PowerShellVersion") -Default "Unknown version"))
            )
        }

        $providerCapabilities = Get-ResultValue -Result $capabilities -Names @("Providers") -Default $null
        $providerValues = @()
        if ($providerCapabilities -is [System.Collections.IDictionary]) {
            $providerValues = @($providerCapabilities.GetEnumerator() | ForEach-Object { $_.Value })
        } elseif ($null -ne $providerCapabilities) {
            $providerValues = @($providerCapabilities.PSObject.Properties | ForEach-Object { $_.Value })
        }
        if ($providerValues.Count -gt 0) {
            $readyProviders = @($providerValues | Where-Object { [string]$_.Status -eq "Ready" }).Count
            $acquisitionProviders = @($providerValues | Where-Object {
                [string]$_.Status -eq "RequiresAcquisition"
            }).Count
            $unsupportedProviders = @($providerValues | Where-Object {
                [string]$_.Status -eq "Unsupported"
            }).Count
            $providerCapabilityDisplay = (
                "$readyProviders ready · $acquisitionProviders require acquisition · " +
                "$unsupportedProviders unsupported"
            )
        }
    }
    $dependencyRows = ""
    foreach ($dependency in $dependencies) {
        $dependencyName = & $displayValue $dependency.Name "Unnamed dependency"
        $dependencyVersion = & $displayValue $dependency.Version "Unknown version"
        $dependencyStatus = & $displayValue $dependency.Status "Verified"
        $dependencyPublisher = & $displayValue $dependency.Publisher "Exact package hash"
        $dependencyArchitecture = & $displayValue $dependency.Architecture "neutral"
        $dependencyPath = & $displayValue $dependency.InstallPath "No install path recorded"
        $dependencyHash = [string]$dependency.Sha256
        $dependencyHashDisplay = if ([string]::IsNullOrWhiteSpace($dependencyHash)) {
            "Publisher verified"
        } else {
            $shortHashLength = [math]::Min(16, $dependencyHash.Length)
            $shortHash = & $encode ($dependencyHash.Substring(0, $shortHashLength))
            "SHA-256 $shortHash..."
        }
        $dependencyRows += @"
<li class="dependency">
  <div><strong>$dependencyName <span>v$dependencyVersion</span></strong><p>$dependencyStatus · $dependencyPublisher · $dependencyArchitecture · $dependencyHashDisplay</p></div>
  <code>$dependencyPath</code>
</li>
"@
    }
    if ([string]::IsNullOrWhiteSpace($dependencyRows)) {
        $dependencyRows = @"
<li class="dependency dependency--empty">
  <div><strong>No acquired dependencies</strong><p>This run did not need to install or load an external update provider.</p></div>
</li>
"@
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
    $serialNumber = & $displayValue (
        Protect-EvidenceText -Text $SysInfo.SerialNumber `
            -SensitiveValues @([string]$SysInfo.SerialNumber)
    )
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

  .dependencies {
    grid-column: 1 / -1;
  }

  .dependency-list {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(260px, 1fr));
    gap: 10px;
    margin: 0;
    padding: 14px 18px 18px;
    list-style: none;
  }

  .dependency {
    min-width: 0;
    padding: 13px 14px;
    background: rgba(54, 197, 240, 0.045);
    border: 1px solid rgba(54, 197, 240, 0.18);
    border-radius: 9px;
  }

  .dependency strong,
  .dependency p,
  .dependency code {
    display: block;
    overflow-wrap: anywhere;
  }

  .dependency strong {
    margin-bottom: 4px;
    color: var(--text);
    font-size: 13px;
  }

  .dependency strong span {
    color: var(--cyan);
    font-family: "Cascadia Mono", "SFMono-Regular", Consolas, monospace;
    font-size: 11px;
  }

  .dependency p {
    margin: 0 0 7px;
    color: var(--muted-strong);
    font-size: 11px;
    line-height: 1.55;
  }

  .dependency code {
    color: var(--muted);
    font-size: 10px;
  }

  .dependency--empty {
    color: var(--muted);
    background: transparent;
    border-color: var(--line-soft);
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
    .dependencies,
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
        <div class="profile-row"><dt>Installation</dt><dd>$(& $displayValue $SysInfo.InstallationType)</dd></div>
        <div class="profile-row"><dt>Architecture</dt><dd>$(& $displayValue $SysInfo.Architecture)</dd></div>
        <div class="profile-row"><dt>Run context</dt><dd>$(& $displayValue $SysInfo.ExecutionContext)</dd></div>
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

    <section class="panel dependencies" aria-labelledby="dependencies-heading">
      <div class="panel-heading">
        <div>
          <h2 id="dependencies-heading">Dependency provenance</h2>
          <p>Version, publisher, digest, architecture, and installed location verified during this run</p>
        </div>
        <span class="section-count">$($dependencies.Count) verified</span>
      </div>
      <ul class="dependency-list">$dependencyRows</ul>
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
        <div class="detail-item"><dt>Platform capability</dt><dd>$platformCapabilityDisplay</dd></div>
        <div class="detail-item"><dt>Provider capability</dt><dd>$providerCapabilityDisplay</dd></div>
        <div class="detail-item"><dt>PowerShell runtime</dt><dd>$runtimeCapabilityDisplay</dd></div>
        <div class="detail-item"><dt>Component rollback</dt><dd>$componentRollbackDisplay</dd></div>
        <div class="detail-item"><dt>Mutation recovery</dt><dd>$mutationRecoveryCount verified actions</dd></div>
        <div class="detail-item"><dt>Evidence retention</dt><dd>$retentionDisplay</dd></div>
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
        $safeHtml = Protect-EvidenceText -Text $html `
            -SensitiveValues @([string]$SysInfo.SerialNumber)
        $htmlValidation = {
            param([string]$CandidatePath)
            try {
                $candidateHtml = [IO.File]::ReadAllText($CandidatePath)
                return (
                    $candidateHtml -match "<!DOCTYPE html>" -and
                    $candidateHtml -match "</html>\s*$"
                )
            } catch {
                return $false
            }
        }
        if (-not (Write-ProtectedAtomicFile -Path $reportFile -Content $safeHtml `
            -ValidationScript $htmlValidation)) {
            throw "HTML report could not be committed: $($script:LastEvidenceWriteError)"
        }
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

function Get-WebhookIdempotencyKey {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RunId
    )

    $hashAlgorithm = [Security.Cryptography.SHA256]::Create()
    try {
        $contractKey = "{0}|webhook-v{1}|{2}|terminal" -f @(
            $script:ProductName,
            $script:WebhookPayloadSchemaVersion,
            $RunId
        )
        $digest = $hashAlgorithm.ComputeHash(
            [Text.Encoding]::UTF8.GetBytes($contractKey)
        )
        return ([BitConverter]::ToString($digest) -replace "-", "").ToLowerInvariant()
    } finally {
        $hashAlgorithm.Dispose()
    }
}

function Get-WebhookEvidenceUri {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$RunData
    )

    $delivery = Get-ResultValue -Result $RunData -Names @("EvidenceDelivery") -Default $null
    $report = if ($null -ne $delivery) {
        Get-ResultValue -Result $delivery -Names @("Report") -Default $null
    } else {
        $null
    }
    $detail = if ($null -ne $report) {
        [string](Get-ResultValue -Result $report -Names @("Detail") -Default "")
    } else {
        ""
    }
    if ([string]::IsNullOrWhiteSpace($detail)) { return "" }
    try {
        if ([IO.Path]::IsPathRooted($detail)) {
            return (New-Object Uri([IO.Path]::GetFullPath($detail))).AbsoluteUri
        }
        $uri = $null
        if ([uri]::TryCreate($detail, [UriKind]::Absolute, [ref]$uri)) {
            return $uri.AbsoluteUri
        }
    } catch {
        return ""
    }
    return ""
}

function New-AzureMonitorEvent {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseShouldProcessForStateChangingFunctions", "", Justification = "Creates an in-memory Azure Monitor/Sentinel event contract.")]
    param(
        [Parameter(Mandatory = $true)][hashtable]$RunData
    )
    $completedAt = [string](Get-ResultValue -Result $RunData -Names @("CompletedAt", "completed_at") -Default (Get-Date).ToUniversalTime().ToString("o"))
    return [ordered]@{
        time = $completedAt
        resourceId = "SystemUpdatePro/$env:COMPUTERNAME"
        category = "SystemUpdatePro"
        operationName = "system_update.completed"
        resultType = [string]$RunData.Status
        resultSignature = [string]$RunData.ExitCode
        correlationId = [string]$RunData.RunId
        dataVersion = "1.0"
        properties = [ordered]@{
            run_id = [string]$RunData.RunId
            hostname = [string]$env:COMPUTERNAME
            status = [string]$RunData.Status
            dry_run = [bool]$DryRun
            exit_code = [int]$RunData.ExitCode
            total_installed = [int]$RunData.TotalInstalled
            total_available = [int]$RunData.TotalAvailable
            total_failed = [int]$RunData.TotalFailed
            reboot_required = [bool]$RunData.RebootRequired
            duration_seconds = [int]$RunData.DurationSeconds
            stages = @($RunData.Stages)
            errors = @($RunData.Errors)
            warnings = @($RunData.Warnings)
        }
    }
}

function New-StructuredEventLogXml {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseShouldProcessForStateChangingFunctions", "", Justification = "Creates an in-memory structured event payload.")]
    param(
        [Parameter(Mandatory = $true)][hashtable]$RunData
    )
    try {
        $document = New-Object System.Xml.XmlDocument
        $root = $document.CreateElement("SystemUpdateProEvent")
        [void]$document.AppendChild($root)
        foreach ($field in @(
            @{ Name = "SchemaVersion"; Value = "1" },
            @{ Name = "RunId"; Value = [string]$RunData.RunId },
            @{ Name = "Status"; Value = [string]$RunData.Status },
            @{ Name = "ExitCode"; Value = [string]$RunData.ExitCode },
            @{ Name = "DryRun"; Value = [string][bool]$DryRun },
            @{ Name = "TotalInstalled"; Value = [string]$RunData.TotalInstalled },
            @{ Name = "TotalAvailable"; Value = [string]$RunData.TotalAvailable },
            @{ Name = "TotalFailed"; Value = [string]$RunData.TotalFailed },
            @{ Name = "RebootRequired"; Value = [string]$RunData.RebootRequired }
        )) {
            $node = $document.CreateElement($field.Name)
            $node.InnerText = [string]$field.Value
            [void]$root.AppendChild($node)
        }
        $stagesNode = $document.CreateElement("Stages")
        foreach ($stage in @($RunData.Stages)) {
            $stageNode = $document.CreateElement("Stage")
            foreach ($field in @("Name", "Provider", "Status", "Attempted", "Available", "Installed", "Failed", "Skipped")) {
                $attribute = $document.CreateAttribute($field)
                $attribute.Value = [string](Get-ResultValue -Result $stage -Names @($field) -Default "")
                [void]$stageNode.Attributes.Append($attribute)
            }
            [void]$stagesNode.AppendChild($stageNode)
        }
        [void]$root.AppendChild($stagesNode)
        return $document.OuterXml
    } catch {
        return ""
    }
}

function Write-PrometheusMetrics {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseShouldProcessForStateChangingFunctions", "", Justification = "Writes an explicitly scoped Prometheus textfile artifact.")]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseSingularNouns", "", Justification = "Prometheus uses the standard plural metrics artifact terminology.")]
    param(
        [Parameter(Mandatory = $true)][hashtable]$RunData,
        [string]$Path = (Join-Path $script:DataPath "metrics.prom"),
        [bool]$DryRunMode = [bool]$script:DryRun
    )
    $result = [ordered]@{ Success = $false; Persisted = $false; Path = $Path; Reason = "" }
    if ($DryRunMode) {
        $result.Success = $true
        $result.Reason = "Dry run did not persist Prometheus metrics"
        return [PSCustomObject]$result
    }
    try {
        $root = [IO.Path]::GetFullPath($script:DataPath).TrimEnd("\", "/")
        $target = [IO.Path]::GetFullPath($Path).TrimEnd("\", "/")
        if (-not $target.StartsWith("$root$([IO.Path]::DirectorySeparatorChar)", [StringComparison]::OrdinalIgnoreCase)) {
            throw "Prometheus path must remain under the protected data directory"
        }
        $statusValue = if ([int]$RunData.ExitCode -le 1) { 1 } else { 0 }
        $labels = 'computer="' + ($env:COMPUTERNAME -replace '[^A-Za-z0-9_.-]', '_') + '"'
        $lines = @(
            "# HELP systemupdatepro_run_success Whether the latest run completed without a failure.",
            "# TYPE systemupdatepro_run_success gauge",
            "systemupdatepro_run_success{$labels} $statusValue",
            "# HELP systemupdatepro_updates_installed Updates installed by the latest run.",
            "# TYPE systemupdatepro_updates_installed gauge",
            "systemupdatepro_updates_installed{$labels} $([int]$RunData.TotalInstalled)",
            "# HELP systemupdatepro_updates_available Updates discovered by the latest run.",
            "# TYPE systemupdatepro_updates_available gauge",
            "systemupdatepro_updates_available{$labels} $([int]$RunData.TotalAvailable)",
            "# HELP systemupdatepro_updates_failed Updates that failed in the latest run.",
            "# TYPE systemupdatepro_updates_failed gauge",
            "systemupdatepro_updates_failed{$labels} $([int]$RunData.TotalFailed)",
            "# HELP systemupdatepro_reboot_required Whether the latest run requires a reboot.",
            "# TYPE systemupdatepro_reboot_required gauge",
            "systemupdatepro_reboot_required{$labels} $([int][bool]$RunData.RebootRequired)"
        )
        if (-not (Write-ProtectedAtomicFile -Path $target -Content (($lines -join "`r`n") + "`r`n") -KeepLastKnownGood:$true)) {
            throw $script:LastEvidenceWriteError
        }
        $result.Success = $true; $result.Persisted = $true; $result.Reason = "Prometheus textfile written atomically"
    } catch {
        $result.Reason = "Prometheus textfile failed: $($_.Exception.Message)"
    }
    return [PSCustomObject]$result
}

function New-WebhookPayload {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseShouldProcessForStateChangingFunctions", "", Justification = "Creates an in-memory webhook contract only.")]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$RunData
    )

    $runId = [string](Get-ResultValue -Result $RunData -Names @("RunId", "run_id") -Default "")
    $idempotencyKey = Get-WebhookIdempotencyKey -RunId $runId
    $overallStatus = switch ([int]$RunData.ExitCode) {
        0 { "success" }
        1 { "success" }
        2 { "partial" }
        default { "failed" }
    }
    $stageSummary = @($RunData.Stages | ForEach-Object {
        [ordered]@{
            name = [string](Get-ResultValue -Result $_ -Names @("Name") -Default "")
            provider = [string](Get-ResultValue -Result $_ -Names @("Provider") -Default "")
            status = [string](Get-ResultValue -Result $_ -Names @("Status") -Default "")
            attempted = [int](Get-ResultValue -Result $_ -Names @("Attempted") -Default 0)
            available = [int](Get-ResultValue -Result $_ -Names @("Available") -Default 0)
            installed = [int](Get-ResultValue -Result $_ -Names @("Installed") -Default 0)
            failed = [int](Get-ResultValue -Result $_ -Names @("Failed") -Default 0)
            skipped = [int](Get-ResultValue -Result $_ -Names @("Skipped") -Default 0)
            reboot_required = [bool](Get-ResultValue -Result $_ -Names @("RebootRequired") -Default $false)
            provider_exit_code = [int](Get-ResultValue -Result $_ -Names @("ProviderExitCode") -Default 0)
        }
    })

    return Protect-EvidenceObject -InputObject ([ordered]@{
        schema_version = $script:WebhookPayloadSchemaVersion
        event_type = "system_update.completed"
        run_id = $runId
        idempotency_key = $idempotencyKey
        started_at = [string](Get-ResultValue -Result $RunData -Names @("StartedAt") -Default "")
        completed_at = [string](Get-ResultValue -Result $RunData -Names @("CompletedAt") -Default "")
        hostname = $env:COMPUTERNAME
        status = $overallStatus
        dry_run = $DryRun.IsPresent
        evidence_uri = Get-WebhookEvidenceUri -RunData $RunData
        oem_updates = [int]$RunData.OEMUpdates
        windows_updates = [int]$RunData.WindowsUpdates
        winget_updates = [int]$RunData.WingetUpdates
        total_installed = [int]$RunData.TotalInstalled
        total_available = [int]$RunData.TotalAvailable
        total_failed = [int]$RunData.TotalFailed
        reboot_required = [bool]$RunData.RebootRequired
        exit_code = [int]$RunData.ExitCode
        runtime_seconds = [int]$RunData.DurationSeconds
        stage_summary = $stageSummary
        errors = @($RunData.Errors)
        warnings = @($RunData.Warnings)
        dependencies = @($RunData.Dependencies)
        mutation_recovery = @($RunData.MutationRecovery)
        capabilities = $RunData.Capabilities
        retention = $RunData.Retention
        dependency_readiness = $RunData.DependencyReadiness
        download_policy = $RunData.DownloadPolicy
        winget_scopes = @($RunData.WingetScopes)
        rollout_decision = $RunData.RolloutDecision
        windows_update_policy = $RunData.WindowsUpdatePolicy
        maintenance_decision = $RunData.MaintenanceDecision
        power_plan_state = $RunData.PowerPlanState
        health = $RunData.Health
        metrics = $RunData.Metrics
        azure_monitor_event = New-AzureMonitorEvent -RunData $RunData
    })
}

function Get-WebhookChannel {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Url
    )

    try {
        $uri = [uri]$Url
        $hostName = $uri.DnsSafeHost.ToLowerInvariant()
        if ($hostName -eq "hooks.slack.com") { return "Slack" }
        if ($hostName.EndsWith(".logic.azure.com") -and
            $uri.AbsolutePath -match "(?i)/workflows/") {
            return "TeamsWorkflow"
        }
        if ($hostName -match "(?i)(^|\.)webhook\.office\.com$" -or
            $hostName -match "(?i)(^|\.)outlook\.office\.com$") {
            return "TeamsConnector"
        }
    } catch {
        return "Generic"
    }
    return "Generic"
}

function ConvertTo-WebhookRequest {
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.IDictionary]$Payload,
        [Parameter(Mandatory = $true)]
        [ValidateSet("Slack", "TeamsWorkflow", "TeamsConnector", "Generic")]
        [string]$Channel
    )

    $statusLabel = ([string]$Payload.status).ToUpperInvariant()
    $modeLabel = if ($Payload.dry_run) { " [DRY RUN]" } else { "" }
    if ($Channel -eq "Slack") {
        $bodyObject = [ordered]@{
            text = (
                "SystemUpdatePro$modeLabel - $($Payload.hostname)`n" +
                "Status: $statusLabel | Installed: $($Payload.total_installed) | " +
                "Failed: $($Payload.total_failed) | Reboot: $($Payload.reboot_required)`n" +
                "Run: $($Payload.run_id) | Evidence: $($Payload.evidence_uri)"
            )
        }
    } elseif ($Channel -eq "TeamsWorkflow") {
        $cardColor = switch ([string]$Payload.status) {
            "success" { "Good" }
            "partial" { "Warning" }
            default { "Attention" }
        }
        $bodyObject = [ordered]@{
            type = "message"
            attachments = @(
                [ordered]@{
                    contentType = "application/vnd.microsoft.card.adaptive"
                    contentUrl = $null
                    content = [ordered]@{
                        '$schema' = "http://adaptivecards.io/schemas/adaptive-card.json"
                        type = "AdaptiveCard"
                        version = "1.4"
                        body = @(
                            [ordered]@{
                                type = "TextBlock"
                                text = "SystemUpdatePro$modeLabel · $statusLabel"
                                size = "Large"
                                weight = "Bolder"
                                color = $cardColor
                                wrap = $true
                            },
                            [ordered]@{
                                type = "FactSet"
                                facts = @(
                                    @{ title = "Endpoint"; value = [string]$Payload.hostname },
                                    @{ title = "Installed"; value = [string]$Payload.total_installed },
                                    @{ title = "Failed"; value = [string]$Payload.total_failed },
                                    @{ title = "Reboot"; value = [string]$Payload.reboot_required },
                                    @{ title = "Run ID"; value = [string]$Payload.run_id },
                                    @{ title = "Evidence"; value = [string]$Payload.evidence_uri }
                                )
                            },
                            [ordered]@{
                                type = "TextBlock"
                                text = "Idempotency: $($Payload.idempotency_key)"
                                isSubtle = $true
                                spacing = "Small"
                                wrap = $true
                            }
                        )
                    }
                }
            )
        }
    } elseif ($Channel -eq "TeamsConnector") {
        $bodyObject = [ordered]@{
            "@type" = "MessageCard"
            "@context" = "http://schema.org/extensions"
            summary = "SystemUpdatePro $statusLabel"
            title = "SystemUpdatePro$modeLabel - $($Payload.hostname)"
            themeColor = switch ([string]$Payload.status) {
                "success" { "00A36C" }
                "partial" { "E0A800" }
                default { "C62828" }
            }
            sections = @(
                @{
                    facts = @(
                        @{ name = "Status"; value = $statusLabel },
                        @{ name = "Installed"; value = [string]$Payload.total_installed },
                        @{ name = "Failed"; value = [string]$Payload.total_failed },
                        @{ name = "Run ID"; value = [string]$Payload.run_id },
                        @{ name = "Evidence"; value = [string]$Payload.evidence_uri },
                        @{ name = "Idempotency"; value = [string]$Payload.idempotency_key }
                    )
                }
            )
        }
    } else {
        $bodyObject = $Payload
    }
    return [PSCustomObject]@{
        Channel = $Channel
        ContentType = "application/json"
        Body = ($bodyObject | ConvertTo-Json -Depth 24 -Compress)
    }
}

function Get-WebhookRetryDelay {
    param(
        [AllowNull()][object]$Headers
    )

    if ($null -eq $Headers) { return $null }
    $value = $null
    try {
        $value = $Headers["Retry-After"]
    } catch {
        $property = $Headers.PSObject.Properties["Retry-After"]
        if ($null -ne $property) { $value = $property.Value }
    }
    if ($value -is [System.Collections.IEnumerable] -and $value -isnot [string]) {
        $value = @($value | Select-Object -First 1)
        if ($value.Count -gt 0) { $value = $value[0] }
    }
    if ([string]::IsNullOrWhiteSpace([string]$value)) { return $null }

    $seconds = 0
    if ([int]::TryParse([string]$value, [ref]$seconds)) {
        return [math]::Max(0, $seconds)
    }
    $retryAt = [DateTimeOffset]::MinValue
    if ([DateTimeOffset]::TryParse([string]$value, [ref]$retryAt)) {
        return [math]::Max(
            0,
            [int][math]::Ceiling(($retryAt - [DateTimeOffset]::UtcNow).TotalSeconds)
        )
    }
    return $null
}

function Invoke-WebhookRequest {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseShouldProcessForStateChangingFunctions", "", Justification = "Sends an explicitly configured HTTPS webhook request and returns transport evidence.")]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Url,
        [Parameter(Mandatory = $true)]
        [string]$Body,
        [Parameter(Mandatory = $true)]
        [System.Collections.IDictionary]$Headers,
        [string]$ContentType = "application/json",
        [ValidateRange(1, 120)]
        [int]$TimeoutSeconds = 30
    )

    $statusCode = 0
    $responseHeaders = $null
    try {
        $response = Invoke-WebRequest -Uri $Url -Method Post -Body $Body `
            -ContentType $ContentType -Headers $Headers -TimeoutSec $TimeoutSeconds `
            -MaximumRedirection 0 -UseBasicParsing -ErrorAction Stop
        $statusCode = [int]$response.StatusCode
        $responseHeaders = $response.Headers
        return [PSCustomObject]@{
            Success = ($statusCode -ge 200 -and $statusCode -lt 300)
            StatusCode = $statusCode
            RetryAfterSeconds = Get-WebhookRetryDelay -Headers $responseHeaders
            Error = ""
        }
    } catch {
        try {
            if ($null -ne $_.Exception.Response) {
                $statusCode = [int]$_.Exception.Response.StatusCode
                $responseHeaders = $_.Exception.Response.Headers
            }
        } catch {
            $statusCode = 0
        }
        return [PSCustomObject]@{
            Success = $false
            StatusCode = $statusCode
            RetryAfterSeconds = Get-WebhookRetryDelay -Headers $responseHeaders
            Error = Protect-EvidenceText -Text $_.Exception.Message
        }
    }
}

function Get-WebhookDeliveryPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RunId
    )

    $safeRunId = [regex]::Replace($RunId, "[^A-Za-z0-9._-]", "_")
    if ([string]::IsNullOrWhiteSpace($safeRunId)) {
        $safeRunId = Get-WebhookIdempotencyKey -RunId $RunId
    }
    if ($safeRunId.Length -gt 128) { $safeRunId = $safeRunId.Substring(0, 128) }
    return Join-Path $script:WebhookDeliveryDirectory "$safeRunId.json"
}

function Test-WebhookDeliveryResult {
    param(
        [AllowNull()][object]$Delivery
    )

    if ($Delivery -isnot [System.Collections.IDictionary] -or
        [int]$Delivery.SchemaVersion -ne $script:WebhookDeliverySchemaVersion) {
        return [PSCustomObject]@{ Valid = $false; Reason = "Webhook delivery schema is invalid" }
    }
    if ([string]::IsNullOrWhiteSpace([string]$Delivery.RunId) -or
        [string]::IsNullOrWhiteSpace([string]$Delivery.IdempotencyKey)) {
        return [PSCustomObject]@{ Valid = $false; Reason = "Webhook delivery correlation is missing" }
    }
    if ([string]$Delivery.TerminalStatus -notin @(
        "Pending", "Retrying", "Succeeded", "Failed", "Rejected"
    )) {
        return [PSCustomObject]@{ Valid = $false; Reason = "Webhook terminal status is invalid" }
    }
    if ([int]$Delivery.AttemptCount -ne @($Delivery.Attempts).Count) {
        return [PSCustomObject]@{ Valid = $false; Reason = "Webhook attempt count is inconsistent" }
    }
    if ([int]$Delivery.MaximumAttempts -lt 1 -or [int]$Delivery.MaximumAttempts -gt 10) {
        return [PSCustomObject]@{ Valid = $false; Reason = "Webhook attempt bound is invalid" }
    }
    return [PSCustomObject]@{ Valid = $true; Reason = "" }
}

function Save-WebhookDeliveryResult {
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.IDictionary]$Delivery
    )

    try {
        if (-not (New-ProtectedDirectory -Path $script:WebhookDeliveryDirectory)) {
            throw "Webhook delivery directory could not be protected"
        }
        $Delivery["LastUpdatedAt"] = (Get-Date).ToUniversalTime().ToString("o")
        $Delivery["LocalRecordPath"] = Get-WebhookDeliveryPath -RunId ([string]$Delivery.RunId)
        $Delivery["LocalRecordStatus"] = "Succeeded"
        if (-not (Write-ProtectedAtomicJson -Path $Delivery.LocalRecordPath `
            -Data $Delivery -Depth 24 `
            -DataValidationScript ${function:Test-WebhookDeliveryResult})) {
            throw $script:LastEvidenceWriteError
        }
        return $true
    } catch {
        $Delivery["LocalRecordStatus"] = "Failed"
        $Delivery["LocalRecordError"] = Protect-EvidenceText -Text $_.Exception.Message
        $script:LastWebhookDeliveryError = $_.Exception.Message
        return $false
    }
}

function Send-WebhookNotification {
    param(
        [string]$Url,
        [hashtable]$RunData,
        [ValidateRange(1, 10)]
        [int]$MaximumAttempts = 3
    )

    $runId = [string](Get-ResultValue -Result $RunData -Names @("RunId", "run_id") -Default "")
    $idempotencyKey = Get-WebhookIdempotencyKey -RunId $runId
    $delivery = [ordered]@{
        SchemaVersion = $script:WebhookDeliverySchemaVersion
        RunId = $runId
        PayloadSchemaVersion = $script:WebhookPayloadSchemaVersion
        IdempotencyKey = $idempotencyKey
        Channel = "Unknown"
        EvidenceUri = ""
        StartedAt = (Get-Date).ToUniversalTime().ToString("o")
        LastUpdatedAt = ""
        CompletedAt = ""
        MaximumAttempts = $MaximumAttempts
        AttemptCount = 0
        Attempts = @()
        TerminalStatus = "Pending"
        LocalRecordStatus = "Pending"
        LocalRecordPath = ""
        LocalRecordError = ""
        Error = ""
    }
    Add-SensitiveEvidenceValue -Value $Url
    $endpointValidation = Test-HttpsWebhookEndpoint -Endpoint $Url
    if (-not $endpointValidation.Valid) {
        $delivery["TerminalStatus"] = "Rejected"
        $delivery["CompletedAt"] = (Get-Date).ToUniversalTime().ToString("o")
        $delivery["Error"] = $endpointValidation.Reason
        [void](Save-WebhookDeliveryResult -Delivery $delivery)
        Write-Log "Webhook notification rejected: $($endpointValidation.Reason)" "WARNING"
        return [PSCustomObject]$delivery
    }

    try {
        $payload = New-WebhookPayload -RunData $RunData
        $channel = Get-WebhookChannel -Url $Url
        $request = ConvertTo-WebhookRequest -Payload $payload -Channel $channel
        $delivery["Channel"] = $channel
        $delivery["EvidenceUri"] = [string]$payload.evidence_uri
    } catch {
        $delivery["TerminalStatus"] = "Failed"
        $delivery["CompletedAt"] = (Get-Date).ToUniversalTime().ToString("o")
        $delivery["Error"] = Protect-EvidenceText -Text $_.Exception.Message
        [void](Save-WebhookDeliveryResult -Delivery $delivery)
        Write-Log "Webhook payload construction failed: $($delivery.Error)" "WARNING"
        return [PSCustomObject]$delivery
    }
    $requestHeaders = [ordered]@{
        "Idempotency-Key" = $idempotencyKey
        "X-SystemUpdatePro-Run-Id" = $runId
    }
    Write-Log "Sending $channel webhook notification..." "DEBUG"

    for ($attemptNumber = 1; $attemptNumber -le $MaximumAttempts; $attemptNumber++) {
        $attemptStartedAt = Get-Date
        $response = Invoke-WebhookRequest -Url $Url -Body $request.Body `
            -ContentType $request.ContentType -Headers $requestHeaders
        $attemptRecord = [ordered]@{
            Attempt = $attemptNumber
            StartedAt = $attemptStartedAt.ToUniversalTime().ToString("o")
            CompletedAt = (Get-Date).ToUniversalTime().ToString("o")
            StatusCode = [int]$response.StatusCode
            Outcome = $(if ($response.Success) { "Succeeded" } else { "Failed" })
            RetryAfterSeconds = $response.RetryAfterSeconds
            DelayBeforeNextSeconds = 0
            Error = Protect-EvidenceText -Text $response.Error
        }

        if ($response.Success) {
            $delivery["Attempts"] = @($delivery.Attempts) + @($attemptRecord)
            $delivery["AttemptCount"] = @($delivery.Attempts).Count
            $delivery["TerminalStatus"] = "Succeeded"
            $delivery["CompletedAt"] = (Get-Date).ToUniversalTime().ToString("o")
            $delivery["Error"] = ""
            [void](Save-WebhookDeliveryResult -Delivery $delivery)
            Write-Log "Webhook notification sent on attempt $attemptNumber" "SUCCESS"
            return [PSCustomObject]$delivery
        }

        $responseStatusCode = [int]$response.StatusCode
        $transientStatus = (
            $responseStatusCode -in @(0, 408, 425, 429) -or
            ($responseStatusCode -ge 500 -and $responseStatusCode -le 599)
        )
        $canRetry = ($transientStatus -and $attemptNumber -lt $MaximumAttempts)
        if ($canRetry) {
            $delaySeconds = if ($null -ne $response.RetryAfterSeconds) {
                [int]$response.RetryAfterSeconds
            } else {
                [int][math]::Pow(2, $attemptNumber)
            }
            $delaySeconds = [math]::Min(60, [math]::Max(0, $delaySeconds))
            $attemptRecord["Outcome"] = "Retrying"
            $attemptRecord["DelayBeforeNextSeconds"] = $delaySeconds
        }
        $delivery["Attempts"] = @($delivery.Attempts) + @($attemptRecord)
        $delivery["AttemptCount"] = @($delivery.Attempts).Count

        if (-not $canRetry) {
            $delivery["TerminalStatus"] = "Failed"
            $delivery["CompletedAt"] = (Get-Date).ToUniversalTime().ToString("o")
            $delivery["Error"] = if (-not [string]::IsNullOrWhiteSpace($response.Error)) {
                Protect-EvidenceText -Text $response.Error
            } else {
                "Webhook returned HTTP status $($response.StatusCode)"
            }
            [void](Save-WebhookDeliveryResult -Delivery $delivery)
            Write-Log "Webhook notification failed after $attemptNumber attempt(s): $($delivery.Error)" "WARNING"
            return [PSCustomObject]$delivery
        }

        $delivery["TerminalStatus"] = "Retrying"
        [void](Save-WebhookDeliveryResult -Delivery $delivery)
        if ($attemptRecord.DelayBeforeNextSeconds -gt 0) {
            Start-Sleep -Seconds $attemptRecord.DelayBeforeNextSeconds
        }
    }
    return [PSCustomObject]$delivery
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
        $eventMessage = New-StructuredEventLogXml -RunData $RunData
        if ([string]::IsNullOrWhiteSpace($eventMessage)) {
            $eventMessage = "SystemUpdatePro completed; Run ID: $($RunData.RunId); Status: $($RunData.Status); Exit code: $($RunData.ExitCode)"
        }
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
            $webhookResult = Send-WebhookNotification -Url $WebhookEndpoint `
                -RunData $RunData -MaximumAttempts $script:MaxRetries
            $delivery.Webhook["SchemaVersion"] = $webhookResult.SchemaVersion
            $delivery.Webhook["PayloadSchemaVersion"] = $webhookResult.PayloadSchemaVersion
            $delivery.Webhook["Channel"] = $webhookResult.Channel
            $delivery.Webhook["IdempotencyKey"] = $webhookResult.IdempotencyKey
            $delivery.Webhook["EvidenceUri"] = $webhookResult.EvidenceUri
            $delivery.Webhook["MaximumAttempts"] = $webhookResult.MaximumAttempts
            $delivery.Webhook["AttemptCount"] = $webhookResult.AttemptCount
            $delivery.Webhook["Attempts"] = @($webhookResult.Attempts)
            $delivery.Webhook["TerminalStatus"] = $webhookResult.TerminalStatus
            $delivery.Webhook["LocalRecordStatus"] = $webhookResult.LocalRecordStatus
            $delivery.Webhook["LocalRecordPath"] = $webhookResult.LocalRecordPath
            if ($webhookResult.TerminalStatus -eq "Succeeded") {
                $delivery.Webhook.Status = "Succeeded"
            } else {
                $delivery.Webhook.Status = "Failed"
                $delivery.Webhook.Detail = [string]$webhookResult.Error
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
# REDACTED DIAGNOSTIC AND RECOVERY BUNDLE
# ============================================================================

function ConvertTo-ProcessArgument {
    param(
        [AllowNull()][string]$Argument
    )

    if ([string]::IsNullOrEmpty($Argument)) { return '""' }
    if ($Argument -notmatch '[\s"]') { return $Argument }

    $builder = New-Object Text.StringBuilder
    $backslash = [char]92
    [void]$builder.Append('"')
    $backslashCount = 0
    foreach ($character in $Argument.ToCharArray()) {
        if ($character -eq $backslash) {
            $backslashCount++
            continue
        }
        if ($character -eq '"') {
            [void]$builder.Append($backslash, (($backslashCount * 2) + 1))
            [void]$builder.Append('"')
            $backslashCount = 0
            continue
        }
        if ($backslashCount -gt 0) {
            [void]$builder.Append($backslash, $backslashCount)
            $backslashCount = 0
        }
        [void]$builder.Append($character)
    }
    if ($backslashCount -gt 0) {
        [void]$builder.Append($backslash, ($backslashCount * 2))
    }
    [void]$builder.Append('"')
    return $builder.ToString()
}

function Invoke-CapturedCommand {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseShouldProcessForStateChangingFunctions", "", Justification = "Runs a caller-supplied diagnostic command without a shell and returns bounded output.")]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,
        [AllowEmptyCollection()][string[]]$ArgumentList = @(),
        [string]$WorkingDirectory = "",
        [ValidateRange(1, 900)]
        [int]$TimeoutSeconds = 120,
        [ValidateRange(1024, 4194304)]
        [int]$MaximumOutputCharacters = 262144
    )

    $startedAt = Get-Date
    $process = $null
    $standardOutput = ""
    $standardError = ""
    $result = [ordered]@{
        SchemaVersion = 1
        Command = [IO.Path]::GetFileName($FilePath)
        StartedAt = $startedAt.ToUniversalTime().ToString("o")
        DurationSeconds = 0
        Success = $false
        TimedOut = $false
        ExitCode = -1
        StandardOutput = ""
        StandardError = ""
        Error = ""
    }
    try {
        if (-not (Test-Path -LiteralPath $FilePath -PathType Leaf)) {
            throw "Diagnostic executable was not found"
        }
        $startInfo = New-Object Diagnostics.ProcessStartInfo
        $startInfo.FileName = $FilePath
        $startInfo.Arguments = (@($ArgumentList | ForEach-Object {
            ConvertTo-ProcessArgument -Argument ([string]$_)
        }) -join " ")
        $startInfo.UseShellExecute = $false
        $startInfo.CreateNoWindow = $true
        $startInfo.RedirectStandardOutput = $true
        $startInfo.RedirectStandardError = $true
        if (-not [string]::IsNullOrWhiteSpace($WorkingDirectory)) {
            $startInfo.WorkingDirectory = $WorkingDirectory
        }

        $process = New-Object Diagnostics.Process
        $process.StartInfo = $startInfo
        if (-not $process.Start()) { throw "Diagnostic process could not be started" }
        $outputTask = $process.StandardOutput.ReadToEndAsync()
        $errorTask = $process.StandardError.ReadToEndAsync()
        if (-not $process.WaitForExit($TimeoutSeconds * 1000)) {
            $result["TimedOut"] = $true
            try { $process.Kill() } catch {
                $result["Error"] = "Timed-out diagnostic process could not be terminated"
            }
            [void]$process.WaitForExit(5000)
        }
        if ($process.HasExited) {
            $standardOutput = [string]$outputTask.Result
            $standardError = [string]$errorTask.Result
            $result["ExitCode"] = [int]$process.ExitCode
        } else {
            $result["Error"] = "Diagnostic process did not terminate after timeout"
        }
        $result["Success"] = (-not $result.TimedOut -and $result.ExitCode -eq 0)
    } catch {
        $result["Error"] = Protect-EvidenceText -Text $_.Exception.Message
    } finally {
        if ($null -ne $process) { $process.Dispose() }
        $result["DurationSeconds"] = [math]::Max(
            0,
            [int]((Get-Date) - $startedAt).TotalSeconds
        )
        foreach ($streamName in @("StandardOutput", "StandardError")) {
            $streamValue = if ($streamName -eq "StandardOutput") {
                $standardOutput
            } else {
                $standardError
            }
            if ($streamValue.Length -gt $MaximumOutputCharacters) {
                $streamValue = (
                    "[TRUNCATED TO LAST $MaximumOutputCharacters CHARACTERS]`r`n" +
                    $streamValue.Substring($streamValue.Length - $MaximumOutputCharacters)
                )
            }
            $result[$streamName] = Protect-EvidenceText -Text $streamValue
        }
    }
    return [PSCustomObject]$result
}

function Read-BoundedEvidenceText {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [ValidateRange(1024, 16777216)]
        [int]$MaximumBytes = 4194304
    )

    $stream = $null
    try {
        $stream = [IO.FileStream]::new(
            $Path,
            [IO.FileMode]::Open,
            [IO.FileAccess]::Read,
            [IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete
        )
        $originalBytes = [long]$stream.Length
        $readLength = [int][math]::Min([long]$MaximumBytes, $originalBytes)
        $offset = [math]::Max(0L, $originalBytes - $readLength)
        [void]$stream.Seek($offset, [IO.SeekOrigin]::Begin)
        $buffer = New-Object byte[] $readLength
        $totalRead = 0
        while ($totalRead -lt $readLength) {
            $read = $stream.Read($buffer, $totalRead, $readLength - $totalRead)
            if ($read -le 0) { break }
            $totalRead += $read
        }
        if ($totalRead -lt $buffer.Length) {
            $trimmed = New-Object byte[] $totalRead
            [Array]::Copy($buffer, $trimmed, $totalRead)
            $buffer = $trimmed
        }

        $encoding = New-Object Text.UTF8Encoding($false, $false)
        if ($offset -eq 0 -and $buffer.Length -ge 2 -and
            $buffer[0] -eq 0xFF -and $buffer[1] -eq 0xFE) {
            $encoding = [Text.Encoding]::Unicode
        } elseif ($offset -eq 0 -and $buffer.Length -ge 2 -and
            $buffer[0] -eq 0xFE -and $buffer[1] -eq 0xFF) {
            $encoding = [Text.Encoding]::BigEndianUnicode
        } else {
            $sampleLength = [math]::Min(4096, $buffer.Length)
            $zeroBytes = 0
            for ($index = 0; $index -lt $sampleLength; $index++) {
                if ($buffer[$index] -eq 0) { $zeroBytes++ }
            }
            if ($sampleLength -gt 0 -and $zeroBytes -gt ($sampleLength / 8)) {
                $encoding = [Text.Encoding]::Unicode
            }
        }
        $text = $encoding.GetString($buffer)
        $truncated = $originalBytes -gt $readLength
        if ($truncated) {
            $text = "[TRUNCATED TO LAST $readLength OF $originalBytes BYTES]`r`n$text"
        }
        return [PSCustomObject]@{
            Success = $true
            Text = Protect-EvidenceText -Text $text
            OriginalBytes = $originalBytes
            IncludedSourceBytes = [long]$readLength
            Truncated = $truncated
            Error = ""
        }
    } catch {
        return [PSCustomObject]@{
            Success = $false
            Text = ""
            OriginalBytes = 0L
            IncludedSourceBytes = 0L
            Truncated = $false
            Error = Protect-EvidenceText -Text $_.Exception.Message
        }
    } finally {
        if ($null -ne $stream) { $stream.Dispose() }
    }
}

function Add-DiagnosticBundleEntry {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseShouldProcessForStateChangingFunctions", "", Justification = "Writes a redacted entry inside a generated diagnostic staging directory.")]
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.IDictionary]$Context,
        [Parameter(Mandatory = $true)]
        [string]$RelativePath,
        [Parameter(Mandatory = $true)]
        [string]$Category,
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()][string]$Content,
        [long]$OriginalBytes = -1,
        [bool]$Truncated = $false,
        [switch]$AllowTruncate
    )

    try {
        $normalizedRelativePath = $RelativePath.Replace("\", "/").TrimStart("/")
        if ([string]::IsNullOrWhiteSpace($normalizedRelativePath) -or
            @($normalizedRelativePath.Split("/") | Where-Object { $_ -eq ".." }).Count -gt 0) {
            throw "Diagnostic entry path is invalid"
        }
        $safeContent = Protect-EvidenceText -Text $Content
        $encoding = New-Object Text.UTF8Encoding($false)
        $contentBytes = $encoding.GetBytes($safeContent)
        if ($OriginalBytes -lt 0) { $OriginalBytes = [long]$contentBytes.Length }

        $remainingBytes = [long]$Context.RemainingBytes
        if ($contentBytes.Length -gt $remainingBytes) {
            if (-not $AllowTruncate -or $remainingBytes -lt 1024) {
                [void]$Context.Omissions.Add([ordered]@{
                    path = $normalizedRelativePath
                    category = $Category
                    reason = "Bundle size budget exhausted"
                    original_bytes = $OriginalBytes
                })
                return $false
            }
            $prefix = "[TRUNCATED BY BUNDLE SIZE BUDGET]`r`n"
            $prefixBytes = $encoding.GetBytes($prefix)
            $tailLength = [math]::Max(
                0,
                [int][math]::Min(
                    [long]$contentBytes.Length,
                    $remainingBytes - $prefixBytes.Length
                )
            )
            $tail = New-Object byte[] $tailLength
            if ($tailLength -gt 0) {
                [Array]::Copy(
                    $contentBytes,
                    $contentBytes.Length - $tailLength,
                    $tail,
                    0,
                    $tailLength
                )
            }
            $safeContent = $prefix + $encoding.GetString($tail)
            $contentBytes = $encoding.GetBytes($safeContent)
            while ($contentBytes.Length -gt $remainingBytes -and $safeContent.Length -gt $prefix.Length) {
                $removeCount = [math]::Min(1024, $safeContent.Length - $prefix.Length)
                $safeContent = $prefix + $safeContent.Substring($prefix.Length + $removeCount)
                $contentBytes = $encoding.GetBytes($safeContent)
            }
            $Truncated = $true
        }

        $rootPath = [IO.Path]::GetFullPath([string]$Context.Root).TrimEnd("\", "/")
        $destinationPath = [IO.Path]::GetFullPath(
            (Join-Path $rootPath $normalizedRelativePath.Replace("/", "\"))
        )
        if (-not $destinationPath.StartsWith(
            "$rootPath$([IO.Path]::DirectorySeparatorChar)",
            [StringComparison]::OrdinalIgnoreCase
        )) {
            throw "Diagnostic entry escaped the staging directory"
        }
        if (-not (Write-ProtectedAtomicFile -Path $destinationPath -Content $safeContent `
            -KeepLastKnownGood:$false)) {
            throw "Diagnostic entry could not be written: $($script:LastEvidenceWriteError)"
        }
        $writtenBytes = [long](Get-Item -LiteralPath $destinationPath).Length
        $Context["RemainingBytes"] = [math]::Max(0L, $remainingBytes - $writtenBytes)
        [void]$Context.Entries.Add([ordered]@{
            path = $normalizedRelativePath
            category = $Category
            bytes = $writtenBytes
            original_bytes = [long]$OriginalBytes
            truncated = [bool]$Truncated
            sha256 = (Get-FileHash -LiteralPath $destinationPath -Algorithm SHA256).Hash.ToLowerInvariant()
        })
        return $true
    } catch {
        [void]$Context.Errors.Add(
            "$(Protect-EvidenceText -Text $RelativePath): $(Protect-EvidenceText -Text $_.Exception.Message)"
        )
        return $false
    }
}

function Add-DiagnosticBundleFile {
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.IDictionary]$Context,
        [Parameter(Mandatory = $true)]
        [string]$SourcePath,
        [Parameter(Mandatory = $true)]
        [string]$RelativePath,
        [Parameter(Mandatory = $true)]
        [string]$Category,
        [ValidateRange(1024, 16777216)]
        [int]$MaximumBytes = 4194304
    )

    $readLimit = [int][math]::Max(
        1024,
        [math]::Min(
            [long]$MaximumBytes,
            [math]::Max(1024L, [long]$Context.RemainingBytes)
        )
    )
    $read = Read-BoundedEvidenceText -Path $SourcePath -MaximumBytes $readLimit
    if (-not $read.Success) {
        [void]$Context.Errors.Add(
            "$(Protect-EvidenceText -Text $RelativePath): $($read.Error)"
        )
        return $false
    }
    return Add-DiagnosticBundleEntry -Context $Context -RelativePath $RelativePath `
        -Category $Category -Content $read.Text -OriginalBytes $read.OriginalBytes `
        -Truncated $read.Truncated -AllowTruncate
}

function Read-DiagnosticJsonSnapshot {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [AllowNull()][scriptblock]$MigrationScript = $null,
        [AllowNull()][scriptblock]$ValidationScript = $null
    )

    $errors = [System.Collections.ArrayList]::new()
    foreach ($candidate in @($Path, "$Path.previous")) {
        if (-not (Test-Path -LiteralPath $candidate -PathType Leaf)) { continue }
        try {
            $access = Test-EvidencePathAccess -Path $candidate
            if (-not $access.Valid) {
                $writeSafety = Test-EvidencePathWriteSafety -Path $candidate -AllowCurrentIdentity
                if (-not $writeSafety.Valid) { throw $writeSafety.Reason }
                [void]$errors.Add(
                    "$([IO.Path]::GetFileName($candidate)) used a legacy but write-safe ACL"
                )
            }
            $data = ConvertTo-Hashtable -InputObject (
                [IO.File]::ReadAllText($candidate) | ConvertFrom-Json -ErrorAction Stop
            )
            if ($null -ne $MigrationScript) { $data = & $MigrationScript $data }
            if ($null -ne $ValidationScript) {
                $validation = & $ValidationScript $data
                if (($validation -is [bool] -and -not $validation) -or
                    ($validation.PSObject.Properties["Valid"] -and -not [bool]$validation.Valid)) {
                    $reason = if ($validation.PSObject.Properties["Reason"]) {
                        [string]$validation.Reason
                    } else {
                        "Diagnostic JSON validation failed"
                    }
                    throw $reason
                }
            }
            return [PSCustomObject]@{
                Success = $true
                Data = Protect-EvidenceObject -InputObject $data
                Recovered = ($candidate -ne $Path)
                Errors = @($errors)
            }
        } catch {
            [void]$errors.Add(
                "$([IO.Path]::GetFileName($candidate)): $(Protect-EvidenceText -Text $_.Exception.Message)"
            )
        }
    }
    return [PSCustomObject]@{
        Success = $false
        Data = $null
        Recovered = $false
        Errors = @($errors)
    }
}

function Get-DiagnosticEvidenceFile {
    $items = [System.Collections.ArrayList]::new()
    if (Test-Path -LiteralPath $LogPath -PathType Container) {
        $specifications = @(
            @{ Pattern = "SystemUpdatePro_Transcript_*.log"; Category = "Transcript"; Count = 2 },
            @{
                Pattern = "SystemUpdatePro_*.log"
                Exclude = "^SystemUpdatePro_Transcript_"
                Category = "RunLog"
                Count = 2
            },
            @{ Pattern = "SystemUpdatePro_Report_*.html"; Category = "Report"; Count = 1 },
            @{ Pattern = "DCU_*.log"; Category = "Dell"; Count = 4 }
        )
        $seenPaths = @{}
        foreach ($specification in $specifications) {
            foreach ($file in @(Get-ChildItem -LiteralPath $LogPath -File -Force `
                -Filter $specification.Pattern -ErrorAction SilentlyContinue | Where-Object {
                    [string]::IsNullOrWhiteSpace([string]$specification.Exclude) -or
                    $_.Name -notmatch [string]$specification.Exclude
                } |
                Sort-Object LastWriteTime -Descending |
                Select-Object -First $specification.Count)) {
                if ($seenPaths.ContainsKey($file.FullName)) { continue }
                $seenPaths[$file.FullName] = $true
                [void]$items.Add([PSCustomObject]@{
                    Path = $file.FullName
                    RelativePath = "evidence/$($specification.Category.ToLowerInvariant())/$($file.Name)"
                    Category = $specification.Category
                    MaximumBytes = 4194304
                })
            }
        }

        foreach ($directory in @(Get-ChildItem -LiteralPath $LogPath -Directory -Force `
            -Filter "HPIA_*" -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending |
            Select-Object -First 2)) {
            foreach ($file in @(Get-ChildItem -LiteralPath $directory.FullName -File -Force `
                -Recurse -ErrorAction SilentlyContinue | Where-Object {
                    $_.Extension.ToLowerInvariant() -in @(
                        ".log", ".txt", ".xml", ".json", ".csv", ".html", ".htm"
                    )
                } | Sort-Object LastWriteTime -Descending | Select-Object -First 8)) {
                $relativeWithinDirectory = $file.FullName.Substring(
                    $directory.FullName.Length
                ).TrimStart("\", "/").Replace("\", "/")
                [void]$items.Add([PSCustomObject]@{
                    Path = $file.FullName
                    RelativePath = "evidence/hp/$($directory.Name)/$relativeWithinDirectory"
                    Category = "HP"
                    MaximumBytes = 4194304
                })
            }
        }
    }
    if (Test-Path -LiteralPath $script:WebhookDeliveryDirectory -PathType Container) {
        foreach ($file in @(Get-ChildItem -LiteralPath $script:WebhookDeliveryDirectory `
            -File -Force -Filter "*.json" -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending | Select-Object -First 3)) {
            [void]$items.Add([PSCustomObject]@{
                Path = $file.FullName
                RelativePath = "evidence/webhook/$($file.Name)"
                Category = "WebhookDelivery"
                MaximumBytes = 1048576
            })
        }
    }
    return @($items | Select-Object -First 30)
}

function Get-DiagnosticWindowsEvidence {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RawDirectory,
        [ValidateRange(1, 30)]
        [int]$SinceDays = 7,
        [ValidateRange(10, 1000)]
        [int]$MaximumEvents = 250
    )

    $errors = [System.Collections.ArrayList]::new()
    $eventSets = [ordered]@{}
    $eventChannels = @(
        "Microsoft-Windows-WindowsUpdateClient/Operational",
        "Microsoft-Windows-UpdateOrchestrator/Operational"
    )
    foreach ($channel in $eventChannels) {
        try {
            $eventSets[$channel] = @(
                Get-WinEvent -FilterHashtable @{
                    LogName = $channel
                    StartTime = (Get-Date).AddDays(-$SinceDays)
                } -MaxEvents $MaximumEvents -ErrorAction Stop | ForEach-Object {
                    $eventMessage = Protect-EvidenceText -Text $_.Message
                    if ($eventMessage.Length -gt 8192) {
                        $eventMessage = $eventMessage.Substring(0, 8192) +
                            "`r`n[TRUNCATED EVENT MESSAGE]"
                    }
                    [ordered]@{
                        time_created = $_.TimeCreated.ToUniversalTime().ToString("o")
                        id = $_.Id
                        level = $_.LevelDisplayName
                        provider = $_.ProviderName
                        message = $eventMessage
                    }
                }
            )
        } catch {
            $eventSets[$channel] = @()
            [void]$errors.Add(
                "$channel`: $(Protect-EvidenceText -Text $_.Exception.Message)"
            )
        }
    }
    try {
        $eventSets["System/WindowsUpdateClient"] = @(
            Get-WinEvent -FilterHashtable @{
                LogName = "System"
                ProviderName = "Microsoft-Windows-WindowsUpdateClient"
                StartTime = (Get-Date).AddDays(-$SinceDays)
            } -MaxEvents $MaximumEvents -ErrorAction Stop | ForEach-Object {
                $eventMessage = Protect-EvidenceText -Text $_.Message
                if ($eventMessage.Length -gt 8192) {
                    $eventMessage = $eventMessage.Substring(0, 8192) +
                        "`r`n[TRUNCATED EVENT MESSAGE]"
                }
                [ordered]@{
                    time_created = $_.TimeCreated.ToUniversalTime().ToString("o")
                    id = $_.Id
                    level = $_.LevelDisplayName
                    provider = $_.ProviderName
                    message = $eventMessage
                }
            }
        )
    } catch {
        $eventSets["System/WindowsUpdateClient"] = @()
        [void]$errors.Add(
            "System/WindowsUpdateClient: $(Protect-EvidenceText -Text $_.Exception.Message)"
        )
    }

    $files = [System.Collections.ArrayList]::new()
    $windowsUpdateLogPath = Join-Path $RawDirectory "WindowsUpdate.log"
    $windowsPowerShellPath = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"
    $commandResult = $null
    if (Test-Path -LiteralPath $windowsPowerShellPath -PathType Leaf) {
        $escapedLogPath = $windowsUpdateLogPath.Replace("'", "''")
        $commandText = (
            "`$ErrorActionPreference='Stop'; " +
            "Get-WindowsUpdateLog -LogPath '$escapedLogPath' | Out-Null"
        )
        $encodedCommand = [Convert]::ToBase64String(
            [Text.Encoding]::Unicode.GetBytes($commandText)
        )
        $commandResult = Invoke-CapturedCommand -FilePath $windowsPowerShellPath `
            -ArgumentList @(
                "-NoLogo", "-NoProfile", "-NonInteractive", "-EncodedCommand", $encodedCommand
            ) -TimeoutSeconds 180
        if ($commandResult.Success -and
            (Test-Path -LiteralPath $windowsUpdateLogPath -PathType Leaf)) {
            [void]$files.Add([PSCustomObject]@{
                Path = $windowsUpdateLogPath
                RelativePath = "windows/WindowsUpdate.log"
                Category = "WindowsUpdate"
                MaximumBytes = 8388608
            })
        } else {
            $detail = if ($null -ne $commandResult -and
                -not [string]::IsNullOrWhiteSpace($commandResult.Error)) {
                $commandResult.Error
            } elseif ($null -ne $commandResult) {
                $commandResult.StandardError
            } else {
                "Windows Update log command did not return evidence"
            }
            [void]$errors.Add("Get-WindowsUpdateLog: $detail")
        }
    } else {
        [void]$errors.Add("Get-WindowsUpdateLog: Windows PowerShell was not found")
    }

    $windowsLogSpecifications = @(
        @{
            Path = Join-Path $script:WindowsRoot "Logs\CBS\CBS.log"
            RelativePath = "windows/CBS.log"
            Category = "CBS"
            MaximumBytes = 4194304
        },
        @{
            Path = Join-Path $script:WindowsRoot "Logs\DISM\dism.log"
            RelativePath = "windows/DISM.log"
            Category = "DISM"
            MaximumBytes = 2097152
        },
        @{
            Path = Join-Path $script:WindowsRoot "SoftwareDistribution\ReportingEvents.log"
            RelativePath = "windows/ReportingEvents.log"
            Category = "WindowsUpdate"
            MaximumBytes = 2097152
        }
    )
    foreach ($specification in $windowsLogSpecifications) {
        if (Test-Path -LiteralPath $specification.Path -PathType Leaf) {
            [void]$files.Add([PSCustomObject]$specification)
        }
    }

    return [PSCustomObject]@{
        SchemaVersion = 1
        SinceDays = $SinceDays
        MaximumEventsPerChannel = $MaximumEvents
        Events = $eventSets
        WindowsUpdateLogCommand = $commandResult
        Files = @($files)
        Errors = @($errors)
    }
}

function Get-DiagnosticRuntimeSnapshot {
    param(
        [AllowNull()][object]$FallbackCapabilities = $null,
        [AllowEmptyCollection()][object[]]$Dependencies = @(),
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Collections.ArrayList]$Errors
    )

    $systemInfo = $null
    $providerVersions = $null
    $capabilities = $FallbackCapabilities
    try {
        $systemInfo = Get-SystemInfo
        Add-SensitiveEvidenceValue -Value ([string]$systemInfo.SerialNumber)
    } catch {
        [void]$Errors.Add("System inventory: $(Protect-EvidenceText -Text $_.Exception.Message)")
    }
    if ($null -ne $systemInfo) {
        try {
            $providerVersions = Get-ProviderVersionInventory -SystemInfo $systemInfo
            $capabilities = Get-CapabilityAssessment -SystemInfo $systemInfo `
                -VersionInventory $providerVersions
        } catch {
            [void]$Errors.Add(
                "Provider inventory: $(Protect-EvidenceText -Text $_.Exception.Message)"
            )
        }
    }

    return Protect-EvidenceObject -InputObject ([ordered]@{
        schema_version = 1
        collected_at = (Get-Date).ToUniversalTime().ToString("o")
        product = $script:ProductName
        product_version = $script:Version
        computer_name = $env:COMPUTERNAME
        powershell = [ordered]@{
            version = $PSVersionTable.PSVersion.ToString()
            edition = [string]$PSVersionTable.PSEdition
            process_architecture = [string]$env:PROCESSOR_ARCHITECTURE
        }
        system = $systemInfo
        provider_versions = $providerVersions
        capabilities = $capabilities
        dependency_provenance = @($Dependencies)
    })
}

function Get-DiagnosticRecoverySnapshot {
    param(
        [AllowEmptyCollection()][object[]]$LastRunActions = @(),
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Collections.ArrayList]$Errors
    )

    $journals = [System.Collections.ArrayList]::new()
    if (Test-Path -LiteralPath $script:MutationJournalDirectory -PathType Container) {
        foreach ($file in @(Get-ChildItem -LiteralPath $script:MutationJournalDirectory `
            -File -Force -Filter "*.json" -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -notmatch "\.previous$" } |
            Sort-Object LastWriteTime -Descending |
            Select-Object -First 20)) {
            $read = Read-DiagnosticJsonSnapshot -Path $file.FullName `
                -ValidationScript ${function:Test-MutationJournal}
            foreach ($errorMessage in @($read.Errors)) {
                [void]$Errors.Add("Journal $($file.Name): $errorMessage")
            }
            if (-not $read.Success) { continue }
            $journalEntries = @($read.Data.Entries | ForEach-Object {
                [ordered]@{
                    id = $_.Id
                    sequence = $_.Sequence
                    type = $_.Type
                    target = Protect-EvidenceText -Text $_.Target
                    scope = $_.Scope
                    state = $_.State
                    restore_on_finalize = $_.RestoreOnFinalize
                    recovery_action = Protect-EvidenceText -Text $_.RecoveryAction
                    created_at = $_.CreatedAt
                    applied_at = $_.AppliedAt
                    recovered_at = $_.RecoveredAt
                    error = Protect-EvidenceText -Text $_.Error
                }
            })
            [void]$journals.Add([ordered]@{
                schema_version = $read.Data.SchemaVersion
                run_id = $read.Data.RunId
                status = $read.Data.Status
                created_at = $read.Data.CreatedAt
                last_updated_at = $read.Data.LastUpdatedAt
                recovered_from_last_known_good = $read.Recovered
                entries = $journalEntries
            })
        }
    }
    return Protect-EvidenceObject -InputObject ([ordered]@{
        schema_version = 1
        collected_at = (Get-Date).ToUniversalTime().ToString("o")
        journals = @($journals)
        last_run_actions = @($LastRunActions)
    })
}

function New-DiagnosticZipArchive {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseShouldProcessForStateChangingFunctions", "", Justification = "Creates a new diagnostic archive at an explicitly supplied temporary path.")]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceDirectory,
        [Parameter(Mandatory = $true)]
        [string]$DestinationPath
    )

    $archiveStream = $null
    $archive = $null
    try {
        Add-Type -AssemblyName System.IO.Compression -ErrorAction Stop
        Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
        $sourceRoot = [IO.Path]::GetFullPath($SourceDirectory).TrimEnd("\", "/")
        $archiveStream = [IO.FileStream]::new(
            $DestinationPath,
            [IO.FileMode]::CreateNew,
            [IO.FileAccess]::ReadWrite,
            [IO.FileShare]::None
        )
        $archive = [IO.Compression.ZipArchive]::new(
            $archiveStream,
            [IO.Compression.ZipArchiveMode]::Create,
            $true
        )
        foreach ($file in @(Get-ChildItem -LiteralPath $sourceRoot -File -Force `
            -Recurse -ErrorAction Stop | Sort-Object FullName)) {
            $relativePath = $file.FullName.Substring($sourceRoot.Length).TrimStart("\", "/")
            $entryName = $relativePath.Replace("\", "/")
            $entry = $archive.CreateEntry(
                $entryName,
                [IO.Compression.CompressionLevel]::Optimal
            )
            $entry.LastWriteTime = $file.LastWriteTime
            $inputStream = $null
            $entryStream = $null
            try {
                $inputStream = [IO.FileStream]::new(
                    $file.FullName,
                    [IO.FileMode]::Open,
                    [IO.FileAccess]::Read,
                    [IO.FileShare]::Read
                )
                $entryStream = $entry.Open()
                $inputStream.CopyTo($entryStream)
            } finally {
                if ($null -ne $entryStream) { $entryStream.Dispose() }
                if ($null -ne $inputStream) { $inputStream.Dispose() }
            }
        }
        $archive.Dispose()
        $archive = $null
        $archiveStream.Flush($true)
        $archiveStream.Dispose()
        $archiveStream = $null
        return $true
    } catch {
        $script:LastEvidenceWriteError = $_.Exception.Message
        return $false
    } finally {
        if ($null -ne $archive) { $archive.Dispose() }
        if ($null -ne $archiveStream) { $archiveStream.Dispose() }
    }
}

function Install-ProtectedAtomicArtifact {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseShouldProcessForStateChangingFunctions", "", Justification = "Atomically publishes a generated and validated binary evidence artifact.")]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourcePath,
        [Parameter(Mandatory = $true)]
        [string]$DestinationPath,
        [AllowNull()][scriptblock]$ValidationScript = $null
    )

    $backupPath = "$DestinationPath.previous"
    $stream = $null
    try {
        if (-not (Test-Path -LiteralPath $SourcePath -PathType Leaf)) {
            throw "Generated artifact does not exist"
        }
        if ($null -ne $ValidationScript -and -not (& $ValidationScript $SourcePath)) {
            throw "Generated artifact validation failed"
        }
        $stream = [IO.FileStream]::new(
            $SourcePath,
            [IO.FileMode]::Open,
            [IO.FileAccess]::ReadWrite,
            [IO.FileShare]::None
        )
        $stream.Flush($true)
        $stream.Dispose()
        $stream = $null
        if (-not (Set-EvidencePathAccess -Path $SourcePath)) {
            throw "Generated artifact ACL could not be protected"
        }
        if (Test-Path -LiteralPath $DestinationPath) {
            Remove-Item -LiteralPath $backupPath -Force -ErrorAction SilentlyContinue
            [IO.File]::Replace($SourcePath, $DestinationPath, $backupPath, $true)
            Remove-Item -LiteralPath $backupPath -Force -ErrorAction SilentlyContinue
        } else {
            [IO.File]::Move($SourcePath, $DestinationPath)
        }
        if (-not (Set-EvidencePathAccess -Path $DestinationPath)) {
            throw "Published artifact ACL could not be protected"
        }
        return $true
    } catch {
        if ($null -ne $stream) { $stream.Dispose() }
        $script:LastEvidenceWriteError = $_.Exception.Message
        return $false
    }
}

function New-DiagnosticBundle {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseShouldProcessForStateChangingFunctions", "", Justification = "Creates a bounded support archive from local product and Windows evidence.")]
    param(
        [ValidateRange(5, 512)]
        [int]$MaximumSizeMB = $DiagnosticBundleMaxSizeMB
    )

    $bundleId = [guid]::NewGuid().ToString()
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $bundleRoot = Join-Path $script:DataPath "Bundles"
    $stagingRoot = Join-Path $bundleRoot ".staging_$($bundleId.Replace('-', ''))"
    $payloadRoot = Join-Path $stagingRoot "payload"
    $rawRoot = Join-Path $stagingRoot "raw"
    $temporaryArchive = Join-Path $bundleRoot ".bundle_$($bundleId.Replace('-', '')).tmp"
    $archivePath = Join-Path $bundleRoot "SystemUpdatePro_Diagnostic_${timestamp}_$($bundleId.Substring(0, 8)).zip"
    $maximumBytes = [long]$MaximumSizeMB * 1MB
    $previousRedactionMode = [string]$script:RedactionMode
    $context = $null
    try {
        # Diagnostic bundles are always fully de-identified, independent of
        # the normal local evidence policy.
        $script:RedactionMode = "SecretsAndSerials"
        foreach ($directory in @($bundleRoot, $stagingRoot, $payloadRoot, $rawRoot)) {
            if (-not (New-ProtectedDirectory -Path $directory)) {
                throw "Diagnostic directory could not be protected: $($script:LastEvidenceAccessError)"
            }
        }
        $context = [ordered]@{
            Root = $payloadRoot
            RemainingBytes = [long][math]::Floor($maximumBytes * 0.75)
            Entries = [System.Collections.ArrayList]::new()
            Omissions = [System.Collections.ArrayList]::new()
            Errors = [System.Collections.ArrayList]::new()
        }

        $historyRead = Read-DiagnosticJsonSnapshot -Path $script:HistoryFile `
            -MigrationScript ${function:Convert-HistorySchema} `
            -ValidationScript ${function:Test-HistoryDocument}
        foreach ($errorMessage in @($historyRead.Errors)) {
            [void]$context.Errors.Add("History: $errorMessage")
        }
        $latestRun = $null
        if ($historyRead.Success -and @($historyRead.Data.entries).Count -gt 0) {
            $latestRun = @($historyRead.Data.entries)[0]
            [void](Add-DiagnosticBundleEntry -Context $context -RelativePath "run/latest_run.json" `
                -Category "RunData" -Content (
                    $latestRun | ConvertTo-Json -Depth 24
                ))
        } else {
            [void]$context.Omissions.Add([ordered]@{
                path = "run/latest_run.json"
                category = "RunData"
                reason = "No validated run history was available"
                original_bytes = 0
            })
        }

        $stateRead = Read-DiagnosticJsonSnapshot -Path $script:StateFile `
            -MigrationScript ${function:Convert-ContinuationStateSchema} `
            -ValidationScript ${function:Test-ContinuationState}
        foreach ($errorMessage in @($stateRead.Errors)) {
            [void]$context.Errors.Add("Continuation state: $errorMessage")
        }
        if ($stateRead.Success) {
            [void](Add-DiagnosticBundleEntry -Context $context `
                -RelativePath "run/active_continuation.json" -Category "RunPolicy" `
                -Content ($stateRead.Data | ConvertTo-Json -Depth 24))
        }

        $policySource = if ($stateRead.Success) {
            $stateRead.Data.Parameters
        } elseif ($null -ne $latestRun) {
            $latestRun.parameters
        } else {
            $null
        }
        $policyDocument = Protect-EvidenceObject -InputObject ([ordered]@{
            schema_version = 1
            source = $(if ($stateRead.Success) {
                "active_continuation"
            } elseif ($null -ne $latestRun) {
                "latest_history"
            } else {
                "unavailable"
            })
            parameters = $policySource
        })
        [void](Add-DiagnosticBundleEntry -Context $context -RelativePath "run/policy.json" `
            -Category "RunPolicy" -Content ($policyDocument | ConvertTo-Json -Depth 16))

        $fallbackCapabilities = if ($null -ne $latestRun) {
            $latestRun.capabilities
        } elseif ($stateRead.Success) {
            $stateRead.Data.Capabilities
        } else {
            $null
        }
        $dependencies = if ($null -ne $latestRun) {
            @($latestRun.dependencies)
        } elseif ($stateRead.Success) {
            @($stateRead.Data.AcquisitionProvenance)
        } else {
            @()
        }
        $inventory = Get-DiagnosticRuntimeSnapshot `
            -FallbackCapabilities $fallbackCapabilities -Dependencies $dependencies `
            -Errors $context.Errors
        [void](Add-DiagnosticBundleEntry -Context $context `
            -RelativePath "inventory/runtime.json" -Category "Inventory" `
            -Content ($inventory | ConvertTo-Json -Depth 24))

        $lastRunRecovery = if ($null -ne $latestRun) {
            @($latestRun.mutation_recovery)
        } else {
            @()
        }
        $recovery = Get-DiagnosticRecoverySnapshot -LastRunActions $lastRunRecovery `
            -Errors $context.Errors
        [void](Add-DiagnosticBundleEntry -Context $context `
            -RelativePath "recovery/status.json" -Category "Recovery" `
            -Content ($recovery | ConvertTo-Json -Depth 20))

        $windowsEvidence = Get-DiagnosticWindowsEvidence -RawDirectory $rawRoot
        foreach ($errorMessage in @($windowsEvidence.Errors)) {
            [void]$context.Errors.Add("Windows evidence: $errorMessage")
        }
        $windowsIndex = Protect-EvidenceObject -InputObject ([ordered]@{
            schema_version = $windowsEvidence.SchemaVersion
            since_days = $windowsEvidence.SinceDays
            maximum_events_per_channel = $windowsEvidence.MaximumEventsPerChannel
            events = $windowsEvidence.Events
            windows_update_log_command = $windowsEvidence.WindowsUpdateLogCommand
        })
        [void](Add-DiagnosticBundleEntry -Context $context `
            -RelativePath "windows/events.json" -Category "WindowsEvents" `
            -Content ($windowsIndex | ConvertTo-Json -Depth 20))
        foreach ($file in @($windowsEvidence.Files)) {
            [void](Add-DiagnosticBundleFile -Context $context -SourcePath $file.Path `
                -RelativePath $file.RelativePath -Category $file.Category `
                -MaximumBytes $file.MaximumBytes)
        }

        foreach ($file in @(Get-DiagnosticEvidenceFile)) {
            $access = Test-EvidencePathAccess -Path $file.Path
            if (-not $access.Valid) {
                $writeSafety = Test-EvidencePathWriteSafety -Path $file.Path -AllowCurrentIdentity
                if (-not $writeSafety.Valid) {
                    [void]$context.Errors.Add(
                        "$($file.RelativePath): source ACL was not trusted"
                    )
                    continue
                }
                [void]$context.Errors.Add(
                    "$($file.RelativePath): source used a legacy but write-safe ACL"
                )
            }
            [void](Add-DiagnosticBundleFile -Context $context -SourcePath $file.Path `
                -RelativePath $file.RelativePath -Category $file.Category `
                -MaximumBytes $file.MaximumBytes)
        }

        $latestRunId = if ($null -ne $latestRun) { [string]$latestRun.run_id } else { "" }
        $latestRunStatus = if ($null -ne $latestRun) { [string]$latestRun.status } else { "" }
        $manifest = Protect-EvidenceObject -InputObject ([ordered]@{
            schema_version = $script:DiagnosticBundleSchemaVersion
            bundle_id = $bundleId
            created_at = (Get-Date).ToUniversalTime().ToString("o")
            product = $script:ProductName
            product_version = $script:Version
            computer_name = $env:COMPUTERNAME
            maximum_archive_bytes = $maximumBytes
            latest_run = [ordered]@{
                run_id = $latestRunId
                status = $latestRunStatus
            }
            files = @($context.Entries)
            omissions = @($context.Omissions)
            collection_errors = @($context.Errors | Select-Object -First 100)
        })
        if (-not (Write-ProtectedAtomicJson -Path (Join-Path $payloadRoot "manifest.json") `
            -Data $manifest -Depth 24 -KeepLastKnownGood:$false)) {
            throw "Diagnostic manifest could not be written: $($script:LastEvidenceWriteError)"
        }

        if (-not (New-DiagnosticZipArchive -SourceDirectory $payloadRoot `
            -DestinationPath $temporaryArchive)) {
            throw "Diagnostic ZIP creation failed: $($script:LastEvidenceWriteError)"
        }
        $zipValidation = {
            param([string]$CandidatePath)
            $archive = $null
            try {
                $archive = [IO.Compression.ZipFile]::OpenRead($CandidatePath)
                return @($archive.Entries | Where-Object {
                    $_.FullName -eq "manifest.json"
                }).Count -eq 1
            } catch {
                return $false
            } finally {
                if ($null -ne $archive) { $archive.Dispose() }
            }
        }
        if (-not (Install-ProtectedAtomicArtifact -SourcePath $temporaryArchive `
            -DestinationPath $archivePath -ValidationScript $zipValidation)) {
            throw "Diagnostic archive could not be published: $($script:LastEvidenceWriteError)"
        }
        $archiveBytes = [long](Get-Item -LiteralPath $archivePath).Length
        if ($archiveBytes -gt $maximumBytes) {
            Remove-Item -LiteralPath $archivePath -Force -ErrorAction SilentlyContinue
            throw "Diagnostic archive exceeded its $MaximumSizeMB MB limit"
        }
        return [PSCustomObject]@{
            Success = $true
            Path = $archivePath
            Bytes = $archiveBytes
            MaximumBytes = $maximumBytes
            EntryCount = @($context.Entries).Count + 1
            OmissionCount = @($context.Omissions).Count
            ErrorCount = @($context.Errors).Count
            Error = ""
        }
    } catch {
        return [PSCustomObject]@{
            Success = $false
            Path = ""
            Bytes = 0L
            MaximumBytes = $maximumBytes
            EntryCount = 0
            OmissionCount = $(if ($null -ne $context) { @($context.Omissions).Count } else { 0 })
            ErrorCount = $(if ($null -ne $context) { @($context.Errors).Count } else { 0 })
            Error = Protect-EvidenceText -Text $_.Exception.Message
        }
    } finally {
        $script:RedactionMode = $previousRedactionMode
        $bundleRootFullPath = [IO.Path]::GetFullPath($bundleRoot).TrimEnd("\", "/")
        if (Test-Path -LiteralPath $stagingRoot -PathType Container) {
            $stagingFullPath = [IO.Path]::GetFullPath($stagingRoot)
            if ($stagingFullPath.StartsWith(
                "$bundleRootFullPath$([IO.Path]::DirectorySeparatorChar).staging_",
                [StringComparison]::OrdinalIgnoreCase
            )) {
                Remove-Item -LiteralPath $stagingFullPath -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
        if (Test-Path -LiteralPath $temporaryArchive -PathType Leaf) {
            Remove-Item -LiteralPath $temporaryArchive -Force -ErrorAction SilentlyContinue
        }
    }
}

# ============================================================================
# MAIN EXECUTION
# ============================================================================

# Handle standalone diagnostic collection before update initialization.
if ($CreateDiagnosticBundle) {
    if (-not (Test-Administrator)) {
        [Console]::Error.WriteLine(
            "[X] Diagnostic bundle creation requires administrator privileges."
        )
        exit 3
    }
    $diagnosticBundle = New-DiagnosticBundle -MaximumSizeMB $DiagnosticBundleMaxSizeMB
    if (-not $diagnosticBundle.Success) {
        [Console]::Error.WriteLine(
            "[X] Diagnostic bundle failed: $($diagnosticBundle.Error)"
        )
        exit 3
    }
    Write-Output $diagnosticBundle.Path
    exit 0
}

# Handle -ShowHistory early exit
if ($ShowHistory) {
    Show-UpdateHistory -Count $HistoryCount
    exit 0
}

$webhookResolution = Resolve-WebhookSecretReference -Reference $WebhookSecretReference
if (-not $webhookResolution.Success) {
    [Console]::Error.WriteLine(
        "[X] Webhook configuration is invalid: $($webhookResolution.Error)"
    )
    exit 3
}
$script:WebhookUrl = [string]$webhookResolution.Url

$scriptStart = $script:RunStartedAt
$sysInfo = @{
    Manufacturer = "Unknown"
    Model = "Unknown"
    SerialNumber = ""
    BIOSVersion = ""
    BIOSDate = $null
    OSName = ""
    OSBuild = ""
    ProductType = 0
    OperatingSystemSKU = 0
    InstallationType = ""
    EditionID = ""
    DisplayVersion = ""
    Architecture = "unknown"
    PowerShellVersion = $PSVersionTable.PSVersion.ToString()
    PowerShellEdition = [string]$PSVersionTable.PSEdition
    ExecutionContext = "Unknown"
    IsServer = $false
    IsServerCore = $false
    Processor = ""
    TotalRAM = 0
}
$lockAcquired = $false
$wsusBypassApplied = $false
$shutdownRequested = $false
$runData = $null
$state = @{}
$oemCapabilityName = ""
$oemCapability = $null
$windowsCapability = $null
$servicingCapability = $null
$wingetCapability = $null
$parallelSafety = $null
$parallelPlanUsed = $false
$parallelPlan = $null

if ($DryRun) {
    try { $script:DryRunMutationBaseline = Get-DryRunMutationSnapshot } catch { $script:DryRunMutationBaseline = $null }
}

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

        $recoveryStart = Get-Date
        $recovery = Invoke-UnfinishedMutationRecovery -ExcludeRunId $(if ($script:ContinuationActive) {
            $script:RunId
        } else {
            ""
        })
        if (($recovery.Attempted + $recovery.Failed) -gt 0) {
            $recoveryStatus = if ($recovery.Failed -gt 0) { "Failed" } else { "Succeeded" }
            $recoveryMessage = if ($recovery.Messages.Count -gt 0) {
                $recovery.Messages -join "; "
            } else {
                "No unfinished privileged mutations required recovery"
            }
            [void](Add-StageResult (New-StageResult -Name "MutationRecovery" `
                -Provider "SystemUpdatePro journal" -Status $recoveryStatus `
                -Attempted ([math]::Max($recovery.Attempted, $recovery.Failed)) -Installed $recovery.Recovered `
                -Failed $recovery.Failed -Message $recoveryMessage -StartedAt $recoveryStart))
            if ($recovery.Failed -gt 0) {
                $message = "Unfinished privileged mutation recovery failed; refusing new system changes"
                [void]$script:Errors.Add($message)
                $script:ExitCode = 3
                break run
            }
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

        [void](Invoke-LogRotation -RetentionDays $LogRetentionDays)
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
        Write-StageProgress -Stage "Preflight" -PercentComplete 5 -Status "Checking platform and policy"
        $preflightStart = Get-Date
        $preflightItems = [System.Collections.ArrayList]::new()

        $sysInfo = Get-SystemInfo
        $script:CurrentSystemInfo = $sysInfo
        Add-SensitiveEvidenceValue -Value ([string]$sysInfo.SerialNumber)
        Write-Log "System: $($sysInfo.Manufacturer) $($sysInfo.Model)" "INFO"
        Write-Log "OS: $($sysInfo.OSName) (Build $($sysInfo.OSBuild))" "DEBUG"
        $script:CapabilityAssessment = Get-CapabilityAssessment -SystemInfo $sysInfo
        $platformCapability = $script:CapabilityAssessment.Platform
        $windowsCapability = Get-AssessedProviderCapability -Provider "WindowsUpdate"
        $servicingCapability = Get-AssessedProviderCapability -Provider "WindowsServicing"
        $wingetCapability = Get-AssessedProviderCapability -Provider "Winget"
        $oemCapabilityName = if ([string]$sysInfo.Manufacturer -match "DELL|ALIENWARE") {
            "Dell"
        } elseif ([string]$sysInfo.Manufacturer -match "LENOVO") {
            "Lenovo"
        } elseif ([string]$sysInfo.Manufacturer -match "HP|HEWLETT") {
            "HP"
        } else {
            ""
        }
        $oemCapability = if ($oemCapabilityName) {
            Get-AssessedProviderCapability -Provider $oemCapabilityName
        } else {
            $null
        }

        $platformItemStatus = switch ([string]$platformCapability.Status) {
            "Ready" { "Succeeded" }
            "Unknown" { "Unknown" }
            default { "Blocked" }
        }
        [void]$preflightItems.Add((New-UpdateItemResult -Name "Platform capability" `
            -Status $platformItemStatus -Message ([string]$platformCapability.Reason) `
            -Evidence @(
                "capability-schema:$($script:CapabilityAssessment.SchemaVersion)",
                "build:$($platformCapability.OSBuild)",
                "installation:$($platformCapability.InstallationType)",
                "architecture:$($platformCapability.Architecture)",
                "context:$($platformCapability.ExecutionContext)"
            )))
        if (-not [bool]$platformCapability.Supported) {
            $message = "Platform capability is $($platformCapability.Status): $($platformCapability.Reason)"
            Write-Log $message "ERROR"
            [void](Add-StageResult (New-StageResult -Name "Preflight" -Status "Failed" `
                -Attempted $preflightItems.Count -Failed 1 -ProviderCode 3 -Message $message `
                -Items @($preflightItems) -StartedAt $preflightStart))
            $script:ExitCode = 3
            break run
        }

        foreach ($capability in @(
            $(if (-not $SkipWindows) { $windowsCapability }),
            $(if (-not $SkipWinget) { $wingetCapability }),
            $(if (-not $SkipOEM -and $null -ne $oemCapability) { $oemCapability })
        )) {
            if ($null -eq $capability) { continue }
            $capabilityItemStatus = switch ([string]$capability.Status) {
                "Ready" { "Succeeded" }
                "RequiresAcquisition" { "Warning" }
                default { "Skipped" }
            }
            [void]$preflightItems.Add((New-UpdateItemResult `
                -Name "$($capability.DisplayName) capability" `
                -Status $capabilityItemStatus `
                -Message ([string]$capability.Reason) `
                -Evidence @(
                    "version-status:$($capability.VersionStatus)",
                    "detected-version:$($capability.DetectedVersion)",
                    "minimum-version:$($capability.MinimumVersion)",
                    "acquisition-version:$($capability.AcquisitionVersion)"
                )))
        }

        $dependencyNames = [System.Collections.ArrayList]::new()
        if (-not $SkipWindows) { [void]$dependencyNames.Add("PSWindowsUpdate") }
        if (-not $SkipWinget) { [void]$dependencyNames.Add("WinGet") }
        if (-not $SkipOEM -and $oemCapabilityName) {
            [void]$dependencyNames.Add($(switch ($oemCapabilityName) {
                "Dell" { "DellCommandUpdate" }
                "Lenovo" { "LSUClient" }
                "HP" { "HPIA" }
            }))
        }
        $dependencyNameArray = @($dependencyNames | ForEach-Object { [string]$_ })
        $script:DependencyReadiness = Get-DependencyReadiness -Names $dependencyNameArray `
            -CachePath $DependencyCachePath -OfflineMode $Offline -TimeoutSeconds $SourceTimeoutSeconds
        foreach ($source in @($script:DependencyReadiness.Sources)) {
            $sourceStatus = if ($source.Ready) { "Succeeded" } else { "Warning" }
            [void]$preflightItems.Add((New-UpdateItemResult -Name "Dependency source: $($source.Name)" `
                -Status $sourceStatus -Message ([string]$source.Reason) `
                -Evidence @("status:$($source.Status)", "host:$($source.Host)", "timeout:$($source.TimeoutSeconds)")))
        }
        if ($Offline) {
            [void]$preflightItems.Add((New-UpdateItemResult -Name "Network connectivity" -Status "Skipped" `
                -Message "Offline mode is active; only verified content-addressed cache artifacts may be consumed"))
        } else {
            $internetEndpoints = @($script:DependencyReadiness.Sources | ForEach-Object { [string]$_.Uri })
            if (Test-InternetConnection -Endpoints $internetEndpoints) {
                [void]$preflightItems.Add((New-UpdateItemResult -Name "Internet connectivity" -Status "Succeeded"))
            } else {
                Write-Log "No configured dependency origin is reachable" "WARNING"
                [void]$preflightItems.Add((New-UpdateItemResult -Name "Internet connectivity" -Status "Warning" `
                    -Message "No configured dependency origin responded to the readiness probe"))
            }
        }

        $networkCost = Get-NetworkCostState
        $script:DownloadPolicy = Get-DownloadPolicy -NetworkCost $networkCost `
            -AllowOverride $AllowMeteredNetwork -DryRunMode $DryRun
        $downloadItemStatus = switch ([string]$script:DownloadPolicy.Status) {
            "Allowed" { "Succeeded" }
            "AllowedWithWarning" { "Warning" }
            default { "Blocked" }
        }
        [void]$preflightItems.Add((New-UpdateItemResult -Name "Provider download policy" `
            -Status $downloadItemStatus -Message ([string]$script:DownloadPolicy.Reason) `
            -Evidence @("network-cost:$($networkCost.Status)", "audited-override:$($script:DownloadPolicy.AuditedOverride)", "dry-run:$DryRun")))

        $script:PackagePolicy = Get-WingetPackagePolicy -Path ([string]$PolicyPath)
        $script:WindowsUpdatePolicy = Get-WindowsUpdatePolicy -Path ([string]$PolicyPath)
        [void]$preflightItems.Add((New-UpdateItemResult -Name "Windows Update policy" -Status "Succeeded" `
            -Message "Feature deferral: $($script:WindowsUpdatePolicy.FeatureDeferralDays) day(s); security-only: $($script:WindowsUpdatePolicy.SecurityOnly); pre-stage: $($script:WindowsUpdatePolicy.PreStage)" `
            -Evidence @(
                "driver-allow:$(@($script:WindowsUpdatePolicy.DriverAllow) -join ',')",
                "driver-deny:$(@($script:WindowsUpdatePolicy.DriverDeny) -join ',')",
                "catalog-fallback:$($script:WindowsUpdatePolicy.CatalogFallback)",
                "admx-snapshot:$($script:WindowsUpdatePolicy.ADMXSnapshot)"
            )))
        $script:MaintenanceDecision = Test-MaintenanceWindow -Policy (Get-MaintenanceWindowPolicy -Path ([string]$PolicyPath))
        [void]$preflightItems.Add((New-UpdateItemResult -Name "Maintenance window" `
            -Status $(switch ([string]$script:MaintenanceDecision.Status) {
                "Allowed" { "Succeeded" }
                "NotConfigured" { "Skipped" }
                default { "Blocked" }
            }) -Message ([string]$script:MaintenanceDecision.Reason) `
            -Evidence @("source:$($script:MaintenanceDecision.Source)", "next:$($script:MaintenanceDecision.NextWindow)")))
        if (-not [bool]$script:MaintenanceDecision.Allowed) {
            $maintenanceMessage = "Update execution is blocked by the maintenance window policy: $($script:MaintenanceDecision.Reason)"
            [void](Add-StageResult (New-StageResult -Name "Preflight" -Status "Failed" -Attempted $preflightItems.Count `
                -Failed 1 -ProviderCode 2 -Message $maintenanceMessage -Items @($preflightItems) -StartedAt $preflightStart))
            $script:ExitCode = 2
            break run
        }
        $script:RolloutPolicy = Get-RolloutPolicy -Path ([string]$RolloutPolicyPath)
        $deviceIdentity = if (-not [string]::IsNullOrWhiteSpace([string]$sysInfo.SerialNumber)) {
            [string]$sysInfo.SerialNumber
        } else { [string]$env:COMPUTERNAME }
        $cohortValue = Get-EndpointCohort -DeviceIdentity $deviceIdentity -Cohort ([string]$script:RolloutPolicy.Cohort)
        $historyForRollout = Read-UpdateHistory
        $rolloutEvidence = Get-RolloutEvidence -HistoryEntries $(if ($null -ne $historyForRollout) {
            @($historyForRollout.entries)
        } else { @() })
        $script:RolloutDecision = Evaluate-RolloutPromotion -Policy $script:RolloutPolicy `
            -Evidence $rolloutEvidence -CohortValue $cohortValue
        [void]$preflightItems.Add((New-UpdateItemResult -Name "Progressive rollout policy" `
            -Status $(switch ([string]$script:RolloutDecision.Decision) {
                "Promote" { "Succeeded" }
                "NotConfigured" { "Skipped" }
                default { "Warning" }
            }) -Message ((@($script:RolloutDecision.Reasons) -join "; ")) `
            -Evidence @("decision:$($script:RolloutDecision.Decision)", "cohort:$cohortValue")))

        $disk = Test-DiskSpace -MinGB $MinDiskSpaceGB
        Write-Log $disk.Message "DEBUG"
        if ($disk.Status -eq "Unknown") {
            $message = "$($disk.Message); -Force cannot override an unknown disk-safety state"
            Write-Log $message "ERROR"
            [void]$preflightItems.Add((New-UpdateItemResult -Name "Disk space" -Status "Unknown" -ProviderCode 4 -Message $message))
            [void](Add-StageResult (New-StageResult -Name "Preflight" -Status "Failed" `
                -Attempted $preflightItems.Count -Failed 1 -ProviderCode 4 -Message $message `
                -Items @($preflightItems) -StartedAt $preflightStart))
            $script:ExitCode = 4
            break run
        } elseif ($disk.Status -eq "Blocked") {
            $message = $disk.Message
            if (-not $Force) {
                Write-Log $message "ERROR"
                [void]$preflightItems.Add((New-UpdateItemResult -Name "Disk space" -Status "Blocked" -ProviderCode 4 -Message $message))
                [void](Add-StageResult (New-StageResult -Name "Preflight" -Status "Failed" `
                    -Attempted $preflightItems.Count -Failed 1 -ProviderCode 4 -Message $message `
                    -Items @($preflightItems) -StartedAt $preflightStart))
                $script:ExitCode = 4
                break run
            }
            Write-Log "$message; overridden for non-firmware work by -Force" "WARNING"
            [void]$preflightItems.Add((New-UpdateItemResult -Name "Disk space" -Status "Warning" `
                -Message "$message; -Force permits non-firmware work only"))
        } else {
            [void]$preflightItems.Add((New-UpdateItemResult -Name "Disk space" -Status "Succeeded" -Message $disk.Message))
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

        $battery = $null
        if ($IncludeBIOS) {
            $battery = Test-BatteryPower -MinimumChargePercent $MinFirmwareChargePercent
            if ($battery.Status -eq "Ready") {
                Write-Log $battery.Message "INFO"
                [void]$preflightItems.Add((New-UpdateItemResult -Name "Firmware power" -Status "Succeeded" -Message $battery.Message))
            } else {
                $powerItemStatus = if ($battery.Status -eq "Blocked") { "Blocked" } else { "Unknown" }
                Write-Log "$($battery.Message); firmware will be excluded and -Force cannot override this decision" "WARNING"
                [void]$preflightItems.Add((New-UpdateItemResult -Name "Firmware power" -Status $powerItemStatus `
                    -ProviderCode 7 -Message "$($battery.Message); firmware will not run"))
            }
        } else {
            [void]$preflightItems.Add((New-UpdateItemResult -Name "Firmware power" -Status "Skipped" `
                -Message "Firmware and BIOS updates not requested; use -IncludeBIOS after validating power safeguards"))
        }

        if (Test-MeteredConnection) {
            Write-Log "Metered connection detected - large downloads may incur charges" "WARNING"
            [void]$preflightItems.Add((New-UpdateItemResult -Name "Network cost" -Status "Warning" -Message "Metered connection detected"))
        } else {
            [void]$preflightItems.Add((New-UpdateItemResult -Name "Network cost" -Status "Succeeded"))
        }

        $bitlocker = Test-BitLockerEnabled
        $bitlockerItemStatus = if ($bitlocker.Status -eq "Unknown") {
            $(if ($IncludeBIOS) { "Unknown" } else { "Warning" })
        } else {
            "Succeeded"
        }
        [void]$preflightItems.Add((New-UpdateItemResult -Name "BitLocker state" -Status $bitlockerItemStatus -Message $bitlocker.Message))
        if ($bitlocker.ProtectionOn) { Write-Log "BitLocker protection: Active" "INFO" }
        if ($IncludeBIOS -and $bitlocker.Status -eq "Unknown") {
            Write-Log "$($bitlocker.Message); firmware will be excluded and -Force cannot override this decision" "WARNING"
        }

        $modelKnown = -not [string]::IsNullOrWhiteSpace([string]$sysInfo.Manufacturer) -and
            -not [string]::IsNullOrWhiteSpace([string]$sysInfo.Model)
        if ($IncludeBIOS) {
            [void]$preflightItems.Add((New-UpdateItemResult -Name "Firmware model inventory" `
                -Status $(if ($modelKnown) { "Succeeded" } else { "Unknown" }) `
                -Message $(if ($modelKnown) {
                    "Detected $($sysInfo.Manufacturer) $($sysInfo.Model); the OEM applicability scan must still verify support"
                } else {
                    "Manufacturer/model is unknown. Repair Win32_ComputerSystem CIM inventory before firmware installation"
                })))
        }

        $script:FirmwarePrerequisites = [PSCustomObject][ordered]@{
            Disk = $disk
            Power = $battery
            BitLocker = $bitlocker
        }
        $firmwarePreflightBlocked = $IncludeBIOS -and (
            $disk.Status -ne "Ready" -or
            $null -eq $battery -or $battery.Status -ne "Ready" -or
            $bitlocker.Status -ne "Ready" -or
            -not $modelKnown
        )
        $preflightStatus = if ($firmwarePreflightBlocked) { "Partial" } else { "Succeeded" }
        $preflightMessage = if ($firmwarePreflightBlocked) {
            "General preflight completed; firmware is blocked until every prerequisite is known-ready"
        } else {
            "Preflight checks completed"
        }
        [void](Add-StageResult (New-StageResult -Name "Preflight" -Status $preflightStatus `
            -Attempted $preflightItems.Count -Message $preflightMessage -Items @($preflightItems) `
            -StartedAt $preflightStart -DurationSeconds ([int]((Get-Date) - $preflightStart).TotalSeconds)))
        Write-StageProgress -Stage "Preflight" -PercentComplete 25 -Status $preflightMessage
        Write-Host ""

        $parallelSafety = Get-ParallelUpdateSafety `
            -DryRunMode ([bool]$DryRun) -SkipOEMMode ([bool]$SkipOEM) `
            -SkipWindowsMode ([bool]$SkipWindows) -IncludeFirmware ([bool]$IncludeBIOS) `
            -Repair ([bool]$RepairWindowsUpdate) -Bypass ([bool]$BypassWSUS) `
            -Cleanup ([bool]$CleanupAfter) -ResetBase ([bool]$ResetComponentBase) `
            -Backup ([bool]$BackupDrivers) -Rollback ([bool]$RollbackDrivers) `
            -PreStageMode ([bool]$PreStage) -ContinueMode ([bool]$ContinueAfterReboot)
        if ($parallelSafety.Allowed) {
            Write-Log "Read-only OEM and Windows Update planning will run concurrently" "INFO"
        }

        if (-not $script:ContinuationActive) {
            $preHealthStart = Get-Date
            $script:PreHealthCheck = Invoke-SystemHealthCheck -Phase "PreRun" -TimeoutSeconds 120
            $preHealthStatus = switch ([string]$script:PreHealthCheck.Status) {
                "Healthy" { "Succeeded" }
                "Unknown" { "Skipped" }
                default { "Succeeded" }
            }
            if ($script:PreHealthCheck.Status -eq "Degraded") {
                Write-Log "Pre-run health baseline is degraded; post-run comparison will only fail on a new regression" "WARNING"
            }
            [void](Add-StageResult (New-StageResult -Name "PreHealth" -Provider "DISM/SFC/CBS" `
                -Status $preHealthStatus -Attempted ([int]$script:PreHealthCheck.Attempted) `
                -Skipped ([int]$script:PreHealthCheck.Skipped) -Message ([string]$script:PreHealthCheck.Reason) `
                -Items @($script:PreHealthCheck.Commands) -Evidence @($script:PreHealthCheck.Evidence) `
                -StartedAt $preHealthStart))

            $restorePointStart = Get-Date
            $restorePointResult = Invoke-RestorePointIfNeeded -RunId $script:RunId `
                -DryRunMode ([bool]$DryRun)
            $restorePointStage = ConvertTo-StageResult -Name "RestorePoint" `
                -Provider "System Restore" -Result $restorePointResult -StartedAt $restorePointStart
            [void](Add-StageResult $restorePointStage)
            if ($restorePointStage.Status -eq "Failed") {
                $script:ExitCode = 2
                Write-Log $restorePointStage.Message "WARNING"
            }

            $powerStart = Get-Date
            $script:PowerPlanState = Set-HighPerformancePowerPlan -DryRunMode ([bool]$DryRun)
            $powerStage = ConvertTo-StageResult -Name "PowerManagement" -Provider "powercfg" `
                -Result @{ Success = $script:PowerPlanState.Success; Message = $script:PowerPlanState.Reason } -StartedAt $powerStart
            [void](Add-StageResult $powerStage)
            if ($powerStage.Status -notin @("Succeeded", "Skipped")) { $script:ExitCode = 2 }
            $policySnapshotReady = $true
            $policySnapshotStart = Get-Date
            if ($script:WindowsUpdatePolicy.ADMXSnapshot -and ($RepairWindowsUpdate -or $BypassWSUS)) {
                $policySnapshot = Invoke-WindowsPolicySnapshot -DryRunMode ([bool]$DryRun)
                $policySnapshotReady = [bool]$policySnapshot.Success
                [void](Add-StageResult (ConvertTo-StageResult -Name "WindowsPolicySnapshot" `
                    -Provider "Windows Update ADMX" -Result @{ Success = $policySnapshot.Success; Message = $policySnapshot.Reason; Evidence = @($policySnapshot.Path) } `
                    -StartedAt $policySnapshotStart))
                if (-not $policySnapshotReady) {
                    $script:ExitCode = 2
                    Write-Log $policySnapshot.Reason "ERROR"
                }
            }
            $repairStart = Get-Date
            if (($RepairWindowsUpdate -or $BypassWSUS) -and -not $policySnapshotReady) {
                if ($RepairWindowsUpdate) {
                    [void](Add-StageResult (New-StageResult -Name "WindowsUpdateRepair" -Provider "Windows servicing" `
                        -Status "Skipped" -Skipped 1 -Message "Windows Update component repair was blocked because the ADMX snapshot failed" -StartedAt $repairStart))
                }
            } elseif ($RepairWindowsUpdate -and -not $servicingCapability.Supported) {
                [void](Add-StageResult (New-StageResult -Name "WindowsUpdateRepair" `
                    -Provider "Windows servicing" -Status "Skipped" -Skipped 1 `
                    -Message $servicingCapability.Reason -StartedAt $repairStart))
            } elseif ($RepairWindowsUpdate) {
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
            if ($BypassWSUS -and -not $policySnapshotReady) {
                [void](Add-StageResult (New-StageResult -Name "WSUSBypass" -Provider "Windows Update policy" `
                    -Status "Skipped" -Skipped 1 -Message "WSUS bypass was blocked because the ADMX snapshot failed" -StartedAt $wsusStart))
            } elseif ($BypassWSUS -and -not $windowsCapability.Supported) {
                [void](Add-StageResult (New-StageResult -Name "WSUSBypass" `
                    -Provider "Windows Update policy" -Status "Skipped" -Skipped 1 `
                    -Message $windowsCapability.Reason -StartedAt $wsusStart))
            } elseif ($BypassWSUS) {
                Set-WSUSBypass -Enable
                $wsusBypassApplied = $true
                [void](Add-StageResult (New-StageResult -Name "WSUSBypass" -Provider "Windows Update policy" `
                    -Status "Succeeded" -Attempted 1 -Message "Temporary WSUS bypass applied" -StartedAt $wsusStart))
            } else {
                [void](Add-StageResult (New-StageResult -Name "WSUSBypass" -Provider "Windows Update policy" `
                    -Status "Skipped" -Skipped 1 -Message "WSUS bypass not requested" -StartedAt $wsusStart))
            }

            $rollbackStart = Get-Date
            if ($RollbackDrivers) {
                $rollbackResult = Invoke-DriverRollback -DryRunMode ([bool]$DryRun)
                $rollbackStage = ConvertTo-StageResult -Name "DriverRollback" `
                    -Provider "DISM" -Result $rollbackResult -StartedAt $rollbackStart
                [void](Add-StageResult $rollbackStage)
                if ($rollbackStage.Status -eq "Failed") {
                    $script:ExitCode = 3
                    Write-Log $rollbackStage.Message "ERROR"
                }
            }

            $backupStart = Get-Date
            if ($RollbackDrivers) {
                [void](Add-StageResult (New-StageResult -Name "DriverBackup" `
                    -Provider "Export-WindowsDriver" -Status "Skipped" -Skipped 1 `
                    -Message "Driver backup skipped while -RollbackDrivers is active" -StartedAt $backupStart))
            } elseif ($BackupDrivers -and -not $SkipOEM -and -not $servicingCapability.Supported) {
                [void](Add-StageResult (New-StageResult -Name "DriverBackup" `
                    -Provider "Export-WindowsDriver" -Status "Skipped" -Skipped 1 `
                    -Message $servicingCapability.Reason -StartedAt $backupStart))
            } elseif ($BackupDrivers -and -not $SkipOEM) {
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
            if ($RollbackDrivers) {
                [void](Add-StageResult (New-StageResult -Name "OEM" -Provider "OEM" -Status "Skipped" `
                    -Skipped 1 -Message "OEM updates skipped while -RollbackDrivers is active" -StartedAt $oemStart))
            } elseif ($parallelSafety.Allowed -and [bool]$windowsCapability.Supported -and
                ($null -eq $oemCapability -or [bool]$oemCapability.Supported)) {
                $parallelPlan = Invoke-ParallelOEMWindowsUpdatePlan -SystemInfo $sysInfo -MaxPasses $MaxUpdatePasses
                $parallelPlanUsed = $true
                $parallelOEMValue = if ($null -ne $parallelPlan.OEM) { $parallelPlan.OEM.Value } else { $null }
                $oemResult = if ($parallelPlan.Success -and $null -ne $parallelOEMValue) {
                    $parallelOEMValue.Result
                } else {
                    @{ Success = $false; Failed = 1; Message = $parallelPlan.Reason; Evidence = @($parallelPlan.Errors) }
                }
                $parallelOEMItemNames = if ($null -ne $parallelOEMValue) {
                    @($parallelOEMValue.ItemNames)
                } else { @() }
                $oemStage = ConvertTo-StageResult -Name "OEM" -Provider "Parallel OEM plan" `
                    -Result $oemResult -ItemNames $parallelOEMItemNames -StartedAt $oemStart
                [void](Add-StageResult $oemStage)
                if ($oemStage.Status -in @("Failed", "Partial")) { $script:ExitCode = 2 }
            } elseif ($SkipOEM) {
                [void](Add-StageResult (New-StageResult -Name "OEM" -Provider "OEM" -Status "Skipped" `
                    -Skipped 1 -Message "OEM updates skipped by -SkipOEM; remove -SkipOEM to request firmware or OEM drivers" `
                    -StartedAt $oemStart))
            } elseif ($null -ne $oemCapability -and -not $oemCapability.Supported) {
                $capabilityItem = New-UpdateItemResult -Name "$($oemCapability.DisplayName) capability" `
                    -Status $(if ($IncludeBIOS) { "Blocked" } else { "Skipped" }) `
                    -Message $oemCapability.Reason `
                    -Evidence @("minimum-version:$($oemCapability.MinimumVersion)")
                [void](Add-StageResult (New-StageResult -Name "OEM" `
                    -Provider $oemCapability.DisplayName `
                    -Status $(if ($IncludeBIOS) { "Partial" } else { "Skipped" }) `
                    -Skipped 1 -Message $oemCapability.Reason -Items @($capabilityItem) `
                    -StartedAt $oemStart))
                Write-Log $oemCapability.Reason "WARNING"
                if ($IncludeBIOS) { $script:ExitCode = 7 }
            } else {
                $manufacturer = [string]$sysInfo.Manufacturer
                $provider = $manufacturer
                $oemResult = $null
                $additionalOemResult = Invoke-AdditionalOEMUpdate -SystemInfo $sysInfo -IncludeGPU

                if ($manufacturer -match "DELL|ALIENWARE") {
                    $provider = "Dell Command Update"
                    $oemResult = Invoke-DellUpdate -IncludeBIOS:$IncludeBIOS -SystemInfo $sysInfo `
                        -FirmwarePrerequisites $script:FirmwarePrerequisites
                } elseif ($manufacturer -match "LENOVO") {
                    $provider = "LSUClient"
                    $oemResult = Invoke-LenovoUpdate -IncludeBIOS:$IncludeBIOS -SystemInfo $sysInfo `
                        -FirmwarePrerequisites $script:FirmwarePrerequisites
                } elseif ($manufacturer -match "HP|HEWLETT") {
                    $provider = "HP Image Assistant"
                    $oemResult = Invoke-HPUpdate -IncludeBIOS:$IncludeBIOS -SystemInfo $sysInfo `
                        -FirmwarePrerequisites $script:FirmwarePrerequisites
                } else {
                    if (@($additionalOemResult.Plans).Count -gt 0) {
                        $provider = "Additional OEM/GPU providers"
                        $oemResult = $additionalOemResult
                    } else {
                        Write-Log "========== OEM UPDATES ==========" "HEADER"
                        $unsupportedMessage = "Manufacturer '$manufacturer' has no supported OEM adapter. Use -SkipOEM for a non-OEM run or add a verified adapter before requesting firmware"
                        Write-Log $unsupportedMessage $(if ($IncludeBIOS) { "WARNING" } else { "INFO" })
                        $unsupportedItem = New-UpdateItemResult -Name "OEM model support" `
                            -Status $(if ($IncludeBIOS) { "Blocked" } else { "Skipped" }) -Message $unsupportedMessage
                        [void](Add-StageResult (New-StageResult -Name "OEM" -Provider $manufacturer `
                            -Status $(if ($IncludeBIOS) { "Partial" } else { "Skipped" }) `
                            -Skipped 1 -Message $unsupportedMessage -Items @($unsupportedItem) -StartedAt $oemStart))
                        if ($IncludeBIOS) { $script:ExitCode = 7 }
                    }
                }

                if ($oemResult) {
                    if ($oemResult -ne $additionalOemResult -and @($additionalOemResult.Items).Count -gt 0) {
                        $oemResult.Items = @($oemResult.Items) + @($additionalOemResult.Items)
                        $oemResult.Available = [int]$oemResult.Available + [int]$additionalOemResult.Available
                        $oemResult.Attempted = [int]$oemResult.Attempted + [int]$additionalOemResult.Attempted
                        $oemResult.Installed = [int]$oemResult.Installed + [int]$additionalOemResult.Installed
                        $oemResult.UpdateCount = [int]$oemResult.UpdateCount + [int]$additionalOemResult.UpdateCount
                        $oemResult.Failed = [int]$oemResult.Failed + [int]$additionalOemResult.Failed
                        $oemResult.Skipped = [int]$oemResult.Skipped + [int]$additionalOemResult.Skipped
                        $oemResult.RebootRequired = [bool]$oemResult.RebootRequired -or [bool]$additionalOemResult.RebootRequired
                        if ([int]$additionalOemResult.Failed -gt 0 -and [int]$oemResult.Installed -gt 0) {
                            $oemResult.Status = "Partial"
                            $oemResult.Success = $false
                        }
                        $oemResult.Message = "$($oemResult.Message); additional providers: $($additionalOemResult.Message)"
                    }
                    $oemStage = ConvertTo-StageResult -Name "OEM" -Provider $provider -Result $oemResult `
                        -ItemNames @($script:OEMUpdates) -StartedAt $oemStart
                    [void](Add-StageResult $oemStage)
                    if ($oemStage.Status -in @("Failed", "Partial")) { $script:ExitCode = 2 }
                    $firmwareReadiness = Get-ResultValue -Result $oemResult -Names @("FirmwareReadiness") -Default $null
                    if ($IncludeBIOS -and $null -ne $firmwareReadiness -and $firmwareReadiness.Status -ne "Ready") {
                        $script:ExitCode = 7
                        Write-Log $firmwareReadiness.Message "WARNING"
                    }
                }
                Write-Host ""
            }
        } else {
            Write-Log "Stages before '$($script:ResumeStageCursor)' were restored from continuation state" "INFO"
        }

        if (Test-ShouldRunContinuationStage -Stage "WindowsUpdate") {
            Write-StageProgress -Stage "Windows Update" -PercentComplete 45 -Status "Processing policy-approved updates"
            $windowsStart = Get-Date
            if ($RollbackDrivers) {
                [void](Add-StageResult (New-StageResult -Name "WindowsUpdate" -Provider "Windows Update" `
                    -Status "Skipped" -Skipped 1 -Message "Windows Update skipped while -RollbackDrivers is active" -StartedAt $windowsStart))
            } elseif ($parallelPlanUsed) {
                $parallelWindowsValue = if ($null -ne $parallelPlan.Windows) { $parallelPlan.Windows.Value } else { $null }
                $wuResult = if ($parallelPlan.Success -and $null -ne $parallelWindowsValue) {
                    $parallelWindowsValue.Result
                } else {
                    @{ Success = $false; Failed = 1; Message = $parallelPlan.Reason; Evidence = @($parallelPlan.Errors) }
                }
                $parallelWindowsItemNames = if ($null -ne $parallelWindowsValue) {
                    @($parallelWindowsValue.ItemNames)
                } else { @() }
                $wuStage = ConvertTo-StageResult -Name "WindowsUpdate" -Provider "Parallel Windows Update plan" `
                    -Result $wuResult -ItemNames $parallelWindowsItemNames -StartedAt $windowsStart
                [void](Add-StageResult $wuStage)
                if ($wuStage.Status -in @("Failed", "Partial")) { $script:ExitCode = 2 }
            } elseif ($SkipWindows) {
                [void](Add-StageResult (New-StageResult -Name "WindowsUpdate" -Provider "Windows Update" `
                    -Status "Skipped" -Skipped 1 -Message "Windows Update skipped by run configuration" -StartedAt $windowsStart))
            } elseif (-not $windowsCapability.Supported) {
                [void](Add-StageResult (New-StageResult -Name "WindowsUpdate" -Provider "Windows Update" `
                    -Status "Skipped" -Skipped 1 -Message $windowsCapability.Reason `
                    -Items @((New-UpdateItemResult -Name "Windows Update capability" -Status "Skipped" `
                        -Message $windowsCapability.Reason)) -StartedAt $windowsStart))
            } else {
                $wuResult = Invoke-WindowsUpdate -MaxPasses $MaxUpdatePasses
                $wuStage = ConvertTo-StageResult -Name "WindowsUpdate" -Provider "Windows Update" `
                    -Result $wuResult -ItemNames @($script:WindowsUpdates) -StartedAt $windowsStart
                [void](Add-StageResult $wuStage)
                if ($wuStage.Status -in @("Failed", "Partial")) { $script:ExitCode = 2 }
                Write-Host ""
            }
            Write-StageProgress -Stage "Windows Update" -PercentComplete 60 -Status "Completed"
            if (-not (Set-ContinuationCursor -StageCursor "Winget")) {
                throw "Failed to persist continuation cursor after Windows Update"
            }
        } else {
            Write-Log "Windows Update stage already completed before continuation resumed" "INFO"
        }

        if (Test-ShouldRunContinuationStage -Stage "Winget") {
            Write-StageProgress -Stage "WinGet" -PercentComplete 65 -Status "Processing application updates"
            $wingetStart = Get-Date
            if ($RollbackDrivers) {
                [void](Add-StageResult (New-StageResult -Name "Winget" -Provider "WinGet" `
                    -Status "Skipped" -Skipped 1 -Message "WinGet skipped while -RollbackDrivers is active" -StartedAt $wingetStart))
            } elseif ($SkipWinget) {
                [void](Add-StageResult (New-StageResult -Name "Winget" -Provider "WinGet" -Status "Skipped" `
                    -Skipped 1 -Message "WinGet skipped by run configuration" -StartedAt $wingetStart))
            } elseif (-not $wingetCapability.Supported) {
                [void](Add-StageResult (New-StageResult -Name "Winget" -Provider "WinGet" `
                    -Status "Skipped" -Skipped 1 -Message $wingetCapability.Reason `
                    -Items @((New-UpdateItemResult -Name "WinGet capability" -Status "Skipped" `
                        -Message $wingetCapability.Reason `
                        -Evidence @("minimum-version:$($wingetCapability.MinimumVersion)"))) `
                    -StartedAt $wingetStart))
                Write-Log $wingetCapability.Reason "WARNING"
            } else {
                $wingetResult = Invoke-WingetUpgradeAll
                $wingetStage = ConvertTo-StageResult -Name "Winget" -Provider "WinGet" `
                    -Result $wingetResult -ItemNames @($script:WingetUpdates) -StartedAt $wingetStart
                [void](Add-StageResult $wingetStage)
                if ($wingetStage.Status -in @("Failed", "Partial")) { $script:ExitCode = 2 }
                Write-Host ""
            }
            Write-StageProgress -Stage "WinGet" -PercentComplete 75 -Status "Completed"
            if (-not (Set-ContinuationCursor -StageCursor "PackageManagers")) {
                throw "Failed to persist continuation cursor after WinGet"
            }
        } else {
            Write-Log "WinGet stage already completed before continuation resumed" "INFO"
        }

        if (Test-ShouldRunContinuationStage -Stage "PackageManagers") {
            Write-StageProgress -Stage "Package managers" -PercentComplete 78 -Status "Processing optional sources"
            $packageStart = Get-Date
            if ($RollbackDrivers) {
                [void](Add-StageResult (New-StageResult -Name "PackageManagers" `
                    -Provider "Chocolatey/Scoop/StoreEdgeFD/WSL" -Status "Skipped" -Skipped 1 `
                    -Message "Package managers skipped while -RollbackDrivers is active" -StartedAt $packageStart))
            } elseif ($SkipWinget) {
                [void](Add-StageResult (New-StageResult -Name "PackageManagers" -Provider "Chocolatey/Scoop/StoreEdgeFD/WSL" `
                    -Status "Skipped" -Skipped 1 -Message "Additional package managers skipped with -SkipWinget" -StartedAt $packageStart))
            } else {
                $packageResult = Invoke-PackageManagerUpgrades
                $packageStage = ConvertTo-StageResult -Name "PackageManagers" `
                    -Provider "Chocolatey/Scoop/StoreEdgeFD/WSL" -Result $packageResult `
                    -ItemNames @($packageResult.Items | ForEach-Object { [string]$_.Name }) -StartedAt $packageStart
                [void](Add-StageResult $packageStage)
                if ($packageStage.Status -in @("Failed", "Partial")) { $script:ExitCode = 2 }
                Write-Host ""
            }
            Write-StageProgress -Stage "Package managers" -PercentComplete 84 -Status "Completed"
            if (-not (Set-ContinuationCursor -StageCursor "Cleanup")) {
                throw "Failed to persist continuation cursor after package managers"
            }
        } else {
            Write-Log "Package manager stage already completed before continuation resumed" "INFO"
        }

        if (Test-ShouldRunContinuationStage -Stage "Cleanup") {
            Write-StageProgress -Stage "Cleanup" -PercentComplete 88 -Status "Finalizing maintenance"
            $cleanupStart = Get-Date
            if ($RollbackDrivers) {
                [void](Add-StageResult (New-StageResult -Name "Cleanup" -Provider "DISM and cleanmgr" `
                    -Status "Skipped" -Skipped 1 -Message "Cleanup skipped while -RollbackDrivers is active" -StartedAt $cleanupStart))
            } elseif (($CleanupAfter -or $ResetComponentBase) -and -not $servicingCapability.Supported) {
                [void](Add-StageResult (New-StageResult -Name "Cleanup" `
                    -Provider "DISM and cleanmgr" -Status "Skipped" -Skipped 1 `
                    -Message $servicingCapability.Reason -StartedAt $cleanupStart))
            } elseif ($CleanupAfter -or $ResetComponentBase) {
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
            Write-StageProgress -Stage "Cleanup" -PercentComplete 95 -Status "Completed"
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
            $rebootDecision = Get-StaggeredRebootDecision
            $leaseResult = Save-StaggeredRebootLease -Decision $rebootDecision
            [void](Add-StageResult (ConvertTo-StageResult -Name "RebootCoordination" -Provider "Cluster policy" `
                -Result @{ Success = ($rebootDecision.Allowed -and $leaseResult.Success); Message = "$($rebootDecision.Reason); $($leaseResult.Reason)"; Evidence = @($leaseResult.Path) } `
                -StartedAt (Get-Date)))
            if (-not $rebootDecision.Allowed -or -not $leaseResult.Success) {
                Write-Log "Automatic reboot deferred by cluster coordination: $($rebootDecision.Reason)" "WARNING"
            } elseif ($Interactive) {
                $answer = Read-Host "Updates require a reboot. Reboot now? [y/N]"
                if ($answer -match "^(?i)y(?:es)?$") {
                    if (-not $ContinueAfterReboot -or $script:ContinuationRegistered) { $shutdownRequested = $true }
                    else { Write-Log "Automatic reboot cancelled because continuation was not registered" "ERROR" }
                } else {
                    Write-Log "Interactive mode deferred the requested reboot" "WARNING"
                }
            } elseif (-not $ContinueAfterReboot -or $script:ContinuationRegistered) {
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
        $hasContinuationJournal = $false
        if ($null -eq $script:MutationJournal) {
            $script:MutationJournal = Get-MutationJournal -RunId $script:RunId
        }
        if ($null -ne $script:MutationJournal) {
            $hasContinuationJournal = @($script:MutationJournal.Entries | Where-Object {
                [string]$_.Scope -eq "Continuation" -and
                [string]$_.State -notin @("Restored", "Committed")
            }).Count -gt 0
        }
        $continuationArtifactsPresent = $script:ContinuationActive -or
            (Test-Path -LiteralPath $script:StateFile) -or $hasContinuationJournal
        $continuationCleanupStart = Get-Date
        $continuationCleanupSucceeded = if ($continuationArtifactsPresent) {
            [bool](Unregister-ContinuationTask)
        } else {
            $true
        }
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

    if (-not $DryRun) {
        $mutationFinalizeStart = Get-Date
        $mutationCountBeforeFinalize = @($script:MutationEvidence).Count
        $preserveMutationScopes = @()
        if ($script:ContinuationRegistered) { $preserveMutationScopes = @("Continuation") }
        $mutationFinalizeSucceeded = Complete-MutationJournal -PreserveScopes $preserveMutationScopes
        $mutationActions = [math]::Max(
            0,
            @($script:MutationEvidence).Count - $mutationCountBeforeFinalize
        )
        if ($mutationCountBeforeFinalize -gt 0 -or $mutationActions -gt 0 -or
            $null -ne $script:MutationJournal) {
            [void](Add-StageResult (New-StageResult -Name "MutationFinalization" `
                -Provider "SystemUpdatePro journal" `
                -Status $(if ($mutationFinalizeSucceeded) { "Succeeded" } else { "Failed" }) `
                -Attempted ([math]::Max(1, $mutationActions)) `
                -Failed $(if ($mutationFinalizeSucceeded) { 0 } else { 1 }) `
                -Message $(if ($mutationFinalizeSucceeded) {
                    $(if ($script:ContinuationRegistered) {
                        "Temporary mutations restored or committed; continuation-task recovery remains durable across reboot"
                    } else {
                        "Temporary privileged mutations restored or committed and verified"
                    })
                } else {
                    "One or more privileged mutations could not be restored or committed"
                }) -StartedAt $mutationFinalizeStart))
        }
    if (-not $mutationFinalizeSucceeded) {
            $message = "Privileged mutation journal finalization failed"
            if (-not ($script:Errors -contains $message)) { [void]$script:Errors.Add($message) }
            $script:ExitCode = 3
        }
    }

    if ($null -ne $script:PowerPlanState) {
        $powerRestoreStart = Get-Date
        $powerRestore = Restore-PowerPlan -State $script:PowerPlanState -DryRunMode ([bool]$DryRun)
        [void](Add-StageResult (ConvertTo-StageResult -Name "PowerManagementRestore" -Provider "powercfg" `
            -Result @{ Success = $powerRestore.Success; Message = $powerRestore.Reason } -StartedAt $powerRestoreStart))
        if (-not $powerRestore.Success) {
            $message = "Power plan restoration failed: $($powerRestore.Reason)"
            if (-not ($script:Errors -contains $message)) { [void]$script:Errors.Add($message) }
            $script:ExitCode = 3
        }
    }

    if ($lockAcquired) {
        Remove-LockFile
    }

    if ($DryRun -and $null -ne $script:DryRunMutationBaseline) {
        try {
            $dryRunComparison = Compare-DryRunMutationSnapshot -Before $script:DryRunMutationBaseline -After (Get-DryRunMutationSnapshot)
            [void](Add-StageResult (New-StageResult -Name "DryRunContract" -Provider "SystemUpdatePro mutation guard" `
                -Status $(if ($dryRunComparison.Changed) { "Failed" } else { "Succeeded" }) `
                -Attempted 1 -Failed $(if ($dryRunComparison.Changed) { 1 } else { 0 }) `
                -Message $dryRunComparison.Reason -Items @(
                    New-UpdateItemResult -Name "Tracked persistent state" `
                        -Status $(if ($dryRunComparison.Changed) { "Failed" } else { "Succeeded" }) `
                        -Message $dryRunComparison.Reason -Evidence @($dryRunComparison.Changes)
                ) -StartedAt (Get-Date)))
            if ($dryRunComparison.Changed) { $script:ExitCode = 3 }
        } catch {
            $message = "Dry-run mutation contract could not be verified: $($_.Exception.Message)"
            [void]$script:Errors.Add($message)
            $script:ExitCode = 3
        }
    }

    if ($null -ne $script:PreHealthCheck) {
        $postHealthStart = Get-Date
        try {
            $script:PostHealthCheck = Invoke-SystemHealthCheck -Phase "PostRun" -TimeoutSeconds 120
            $script:HealthRegression = Compare-SystemHealthRegression `
                -Before $script:PreHealthCheck -After $script:PostHealthCheck
            $postHealthStatus = if ($script:HealthRegression.Regressed) {
                "Failed"
            } elseif ($script:PostHealthCheck.Status -eq "Unknown") {
                "Skipped"
            } else { "Succeeded" }
            $postHealthFailed = if ($script:HealthRegression.Regressed) { 1 } else { 0 }
            [void](Add-StageResult (New-StageResult -Name "PostHealth" -Provider "DISM/SFC/CBS" `
                -Status $postHealthStatus -Attempted ([int]$script:PostHealthCheck.Attempted) `
                -Failed $postHealthFailed -Skipped ([int]$script:PostHealthCheck.Skipped) `
                -Message ([string]$script:HealthRegression.Reason) `
                -Items @($script:PostHealthCheck.Commands) -Evidence @(
                    $script:PostHealthCheck.Evidence + "before:$($script:HealthRegression.Before)" +
                    "after:$($script:HealthRegression.After)"
                ) -StartedAt $postHealthStart))
            if ($script:HealthRegression.Regressed) {
                $message = "Post-run health regression detected: $($script:PostHealthCheck.Reason)"
                if (-not ($script:Errors -contains $message)) { [void]$script:Errors.Add($message) }
                $script:ExitCode = 3
            } elseif ($script:PostHealthCheck.Status -eq "Degraded") {
                Write-Log "Post-run health remains degraded but no regression from the pre-run baseline was established" "WARNING"
            }
        } catch {
            $message = "Post-run health check failed to complete: $($_.Exception.Message)"
            [void](Add-StageResult (New-StageResult -Name "PostHealth" -Provider "DISM/SFC/CBS" `
                -Status "Skipped" -Skipped 1 -Message $message -StartedAt $postHealthStart))
            Write-Log $message "WARNING"
        }
    }

    $completedAt = Get-Date
    $runData = New-RunData -StartedAt $scriptStart -CompletedAt $completedAt -RequestedExitCode $script:ExitCode
    $runData.Metrics = Write-PrometheusMetrics -RunData $runData -DryRunMode ([bool]$DryRun)
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
        [void](Protect-EvidenceFile -Path $script:TranscriptFile `
            -SensitiveValues @([string]$sysInfo.SerialNumber))
    }
}

if ($shutdownRequested) {
    Write-Log "Rebooting in 30 seconds..." "WARNING"
    shutdown.exe /r /t 30 /c "SystemUpdatePro - Reboot Required"
}

exit $script:ExitCode
