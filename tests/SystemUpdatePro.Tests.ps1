BeforeAll {
    $scriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) "SystemUpdatePro.ps1"
    $script:SourceScriptPath = $scriptPath
    $tokens = $null
    $parseErrors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile(
        $scriptPath,
        [ref]$tokens,
        [ref]$parseErrors
    )
    if ($parseErrors.Count -gt 0) {
        throw "SystemUpdatePro.ps1 did not parse: $($parseErrors -join '; ')"
    }

    $functionNames = @(
        "New-UpdateItemResult",
        "New-StageResult",
        "Get-ResultValue",
        "ConvertTo-StageResult",
        "Add-StageResult",
        "Get-RunExitCode",
        "New-RunData",
        "Test-CurrentIdentityIsAdministrator",
        "Set-EvidencePathAccess",
        "Test-EvidencePathAccess",
        "Test-EvidencePathWriteSafety",
        "New-ProtectedDirectory",
        "Add-SensitiveEvidenceValue",
        "Protect-EvidenceText",
        "Protect-EvidenceObject",
        "Write-ProtectedAtomicFile",
        "Write-ProtectedAtomicJson",
        "Add-ProtectedEvidenceLine",
        "Move-EvidenceToQuarantine",
        "Read-ProtectedJsonFile",
        "Protect-EvidenceFile",
        "Protect-EvidenceTree",
        "Test-HttpsWebhookEndpoint",
        "Resolve-WebhookSecretReference",
        "Test-LockDocument",
        "Get-AcquisitionManifest",
        "Get-AcquisitionManifestEntry",
        "Get-DependencyCacheArtifactPath",
        "Test-DependencyCacheArtifact",
        "Get-SourceReadiness",
        "Get-DependencyReadiness",
        "Get-SystemArchitecture",
        "Get-SystemInfo",
        "Get-CapabilityMatrix",
        "Get-ProviderVersionInventory",
        "Get-PlatformCapability",
        "Get-ProviderCapability",
        "Get-CapabilityAssessment",
        "Get-AssessedProviderCapability",
        "Test-AcquisitionUri",
        "ConvertTo-SafeVersion",
        "Test-VersionAtLeast",
        "Get-AuthenticodeEvidence",
        "Test-AcquiredFile",
        "Invoke-VerifiedDownload",
        "Add-AcquisitionProvenance",
        "New-SystemUpdateProTemporaryDirectory",
        "Remove-SystemUpdateProTemporaryDirectory",
        "ConvertTo-Hashtable",
        "Get-EffectiveRunParameter",
        "Get-ContinuationParameterName",
        "Convert-ContinuationStateSchema",
        "Test-ContinuationState",
        "Set-ContinuationStateAccess",
        "Test-ContinuationStateAccess",
        "Save-State",
        "Move-StateToQuarantine",
        "Get-State",
        "Clear-State",
        "New-ContinuationState",
        "Import-ContinuationState",
        "Set-ContinuationCursor",
        "Test-ShouldRunContinuationStage",
        "Get-MutationJournalPath",
        "New-MutationJournal",
        "Test-MutationJournal",
        "Save-MutationJournal",
        "Get-MutationJournal",
        "Get-OrCreateMutationJournal",
        "Add-MutationJournalEntry",
        "Set-MutationJournalEntryState",
        "Test-MutationDirectoryContract",
        "Restore-DirectoryRenameSnapshot",
        "Invoke-JournaledDirectoryReset",
        "Complete-DirectoryRenameSnapshot",
        "Restore-MutationJournalEntry",
        "Complete-MutationJournalEntry",
        "Restore-MutationJournalScope",
        "Complete-MutationJournal",
        "Invoke-UnfinishedMutationRecovery",
        "Get-EvidenceArtifactSize",
        "Get-EvidenceRetentionCandidate",
        "Remove-EvidenceRetentionCandidate",
        "Invoke-EvidenceRetention",
        "Invoke-LogRotation",
        "Convert-HistorySchema",
        "Test-HistoryDocument",
        "Read-UpdateHistory",
        "Get-ScheduledTaskSnapshot",
        "Restore-ScheduledTaskSnapshot",
        "Register-ContinuationTask",
        "Test-ContinuationTask",
        "Unregister-ContinuationTask",
        "Get-SystemPowerStatus",
        "Test-DiskSpace",
        "Test-BatteryPower",
        "Test-BitLockerEnabled",
        "Test-OEMUpdateIsFirmware",
        "Test-FirmwareReadiness",
        "Get-FirmwareUpdatePolicy",
        "New-FirmwareReadinessItem",
        "Get-NetworkCostState",
        "Get-DownloadPolicy",
        "Test-DownloadAllowed",
        "Get-PolicyDocument",
        "Get-WingetScopePlan",
        "ConvertFrom-WingetPackageOutput",
        "Get-WingetScopeInventory",
        "Get-WingetPackagePolicy",
        "Get-WingetPackagePlan",
        "Get-EndpointCohort",
        "Get-RolloutPolicy",
        "Get-RolloutEvidence",
        "Evaluate-RolloutPromotion",
        "Get-DirectoryPayloadHash",
        "Get-PSModuleInstallPath",
        "Test-InstalledPSModule",
        "Get-VerifiedPSModulePath",
        "Install-VerifiedPSModule",
        "Get-InstalledExecutableEvidence",
        "Test-WingetPackageContractText",
        "Test-DellWinGetPackageContract",
        "Get-WingetTrustEvidence",
        "Test-WingetInstalled",
        "Install-Winget",
        "Get-DCUPath",
        "Get-DellInventoryCollectorEvidence",
        "Test-DellInventoryCollector",
        "Update-DellInventoryCollector",
        "Install-DellCommandUpdate",
        "Get-ServiceStartupType",
        "Get-ServiceDelayedAutoStartSnapshot",
        "Set-ServiceStartupType",
        "Get-ServiceSnapshot",
        "Restore-ServiceSnapshot",
        "Set-JournaledServiceState",
        "Test-WindowsUpdateRuntime",
        "Repair-WindowsUpdateServices",
        "Set-WSUSBypass",
        "Repair-DellServices",
        "Invoke-DellUpdate",
        "Invoke-LenovoUpdate",
        "Get-HPIAPath",
        "Install-HPIA",
        "Invoke-HPUpdate",
        "Get-GPUProviderPlan",
        "Get-AdditionalOEMProviderPlan",
        "Invoke-AdditionalOEMUpdate",
        "Get-RegistryValueSnapshot",
        "Test-MutationRegistryTarget",
        "ConvertTo-RegistryValueKind",
        "Test-RegistrySnapshotEqual",
        "Restore-RegistryValueSnapshot",
        "Invoke-ComponentCleanup",
        "Invoke-WindowsUpdate",
        "New-HTMLReport",
        "Initialize-EventLog",
        "Write-EventLogEntry",
        "Get-WebhookIdempotencyKey",
        "Get-WebhookEvidenceUri",
        "New-WebhookPayload",
        "Get-WebhookChannel",
        "ConvertTo-WebhookRequest",
        "Get-WebhookRetryDelay",
        "Invoke-WebhookRequest",
        "Get-WebhookDeliveryPath",
        "Test-WebhookDeliveryResult",
        "Save-WebhookDeliveryResult",
        "Send-WebhookNotification",
        "Save-UpdateHistory",
        "Invoke-TerminalEvidence",
        "ConvertTo-ProcessArgument",
        "Invoke-CapturedCommand",
        "Read-BoundedEvidenceText",
        "Add-DiagnosticBundleEntry",
        "Add-DiagnosticBundleFile",
        "Read-DiagnosticJsonSnapshot",
        "Get-DiagnosticEvidenceFile",
        "Get-DiagnosticWindowsEvidence",
        "Get-DiagnosticRuntimeSnapshot",
        "Get-DiagnosticRecoverySnapshot",
        "New-DiagnosticZipArchive",
        "Install-ProtectedAtomicArtifact",
        "New-DiagnosticBundle",
        "Install-PSModuleWithRetry",
        "Invoke-WindowsUpdatePSWU",
        "Invoke-WindowsUpdateWUA",
        "Write-Log"
    )

    $functionAsts = $ast.FindAll({
        param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
        $functionNames -contains $node.Name
    }, $true)

    foreach ($functionAst in $functionAsts) {
        . ([scriptblock]::Create($functionAst.Extent.Text))
    }

    function Get-LSUpdate {}
    function Install-LSUpdate {
        param([Parameter(ValueFromPipeline = $true)][object]$InputObject)
        process { $InputObject }
    }

    function Initialize-SystemUpdateProTestState {
        $script:DryRun = [switch]$false
        $script:WebhookUrl = ""
        $script:WebhookSecretReference = ""
        $script:ResultSchemaVersion = 1
        $script:CapabilitySchemaVersion = 1
        $script:CapabilityAssessment = $null
        $script:HistorySchemaVersion = 2
        $script:LockSchemaVersion = 1
        $script:DiagnosticBundleSchemaVersion = 1
        $script:WebhookPayloadSchemaVersion = 2
        $script:WebhookDeliverySchemaVersion = 1
        $script:RunId = "11111111-1111-1111-1111-111111111111"
        $script:StageResults = [System.Collections.ArrayList]::new()
        $script:Errors = [System.Collections.ArrayList]::new()
        $script:Warnings = [System.Collections.ArrayList]::new()
        $script:RebootRequired = $false
        $script:RunFinalized = $false
        $script:LastEvidenceDelivery = $null
        $script:OEMUpdates = [System.Collections.ArrayList]::new()
        $script:WindowsUpdates = [System.Collections.ArrayList]::new()
        $script:WingetUpdates = [System.Collections.ArrayList]::new()
        $script:StateSchemaVersion = 5
        $script:MaxContinuationAttempts = 3
        $script:RunStartedAt = Get-Date
        $stateTestDirectory = Join-Path $TestDrive ([guid]::NewGuid().ToString("N"))
        New-Item -ItemType Directory -Path $stateTestDirectory -Force | Out-Null
        $script:StateFile = Join-Path $stateTestDirectory "state.json"
        $script:HistoryFile = Join-Path $stateTestDirectory "update_history.json"
        $script:LockFile = Join-Path $stateTestDirectory "update.lock"
        $script:DataPath = $stateTestDirectory
        $script:TaskName = "SystemUpdatePro_TestContinue"
        $script:ProductName = "SystemUpdatePro"
        $script:ContinuationAttempt = 0
        $script:ContinuationActive = $false
        $script:ContinuationRegistered = $false
        $script:ContinuationState = $null
        $script:ResumeStageCursor = ""
        $script:FirmwarePrerequisites = $null
        $script:Version = "4.1.0"
        $script:Offline = $false
        $script:DependencyCachePath = Join-Path $stateTestDirectory "Cache"
        $script:SourceTimeoutSeconds = 30
        $script:AllowMeteredNetwork = $false
        $script:PolicyPath = ""
        $script:RolloutPolicyPath = ""
        $script:DependencyReadiness = $null
        $script:DownloadPolicy = $null
        $script:CurrentSystemInfo = $null
        $script:WingetScopeResults = @()
        $script:RolloutDecision = $null
        $script:PackagePolicy = $null
        $script:AcquisitionManifestVersion = 1
        $script:AcquisitionManifest = Get-AcquisitionManifest
        $script:AcquisitionProvenance = [System.Collections.ArrayList]::new()
        $script:MutationJournalSchemaVersion = 1
        $script:MutationJournalDirectory = Join-Path $stateTestDirectory "Journals"
        New-Item -ItemType Directory -Path $script:MutationJournalDirectory -Force | Out-Null
        $script:MutationJournal = $null
        $script:MutationEvidence = [System.Collections.ArrayList]::new()
        $script:WebhookDeliveryDirectory = Join-Path $stateTestDirectory "WebhookDeliveries"
        $script:RetentionResult = $null
        $script:SensitiveEvidenceValues = [System.Collections.ArrayList]::new()
        $script:ProtectedEvidenceDirectories = @{}
        $script:WindowsRoot = Join-Path $stateTestDirectory "Windows"
        New-Item -ItemType Directory -Path (Join-Path $script:WindowsRoot "System32") -Force | Out-Null
        $script:PSModuleInstallRoot = Join-Path $stateTestDirectory "Modules"
        $script:HPIAInstallRoot = Join-Path $stateTestDirectory "HPIA"

        $script:SkipOEM = [switch]$false
        $script:SkipWindows = [switch]$false
        $script:SkipWinget = [switch]$false
        $script:IncludeBIOS = [switch]$false
        $script:BypassWSUS = [switch]$false
        $script:RepairWindowsUpdate = [switch]$false
        $script:CleanupAfter = [switch]$false
        $script:ResetComponentBase = [switch]$false
        $script:ContinueAfterReboot = [switch]$true
        $script:BackupDrivers = [switch]$false
        $script:ShowHistory = [switch]$false
        $script:Reboot = [switch]$true
        $script:Force = [switch]$false
        $script:HistoryCount = 10
        $script:MaxRetries = 3
        $script:MaxUpdatePasses = 3
        $script:MinDiskSpaceGB = 10
        $script:MinFirmwareChargePercent = 50
        $script:LogPath = $stateTestDirectory
        $script:LogRetentionDays = 30
        $script:EvidenceMaxSizeMB = 512
        $script:RedactionMode = "SecretsAndSerials"
        $script:DiagnosticBundleMaxSizeMB = 50
        $script:LogFile = Join-Path $stateTestDirectory "SystemUpdatePro_test.log"
        $script:TranscriptFile = Join-Path $stateTestDirectory "SystemUpdatePro_Transcript_test.log"
        $script:EntryScriptPath = $script:SourceScriptPath
    }

    function New-CapabilitySystemInfo {
        param(
            [string]$Manufacturer = "Contoso",
            [string]$OSBuild = "22631",
            [int]$ProductType = 1,
            [string]$InstallationType = "Client",
            [string]$EditionID = "Professional",
            [int]$OperatingSystemSKU = 48,
            [string]$Architecture = "x64",
            [string]$PowerShellVersion = "5.1.22621.2506",
            [string]$PowerShellEdition = "Desktop",
            [string]$RunContext = "AdministratorUser"
        )

        return @{
            Manufacturer = $Manufacturer
            OSName = "Microsoft Windows"
            OSVersion = "10.0.$OSBuild"
            OSBuild = $OSBuild
            ProductType = $ProductType
            InstallationType = $InstallationType
            EditionID = $EditionID
            OperatingSystemSKU = $OperatingSystemSKU
            Architecture = $Architecture
            PowerShellVersion = $PowerShellVersion
            PowerShellEdition = $PowerShellEdition
            ExecutionContext = $RunContext
        }
    }

    function New-CapabilityVersionInventory {
        param(
            [string]$WingetStatus = "Ready",
            [string]$DellStatus = "NotApplicable",
            [string]$LenovoStatus = "NotApplicable",
            [string]$HPStatus = "NotApplicable"
        )

        return [ordered]@{
            WindowsUpdate = [ordered]@{
                Status = "BuiltInFallback"; Version = "10.0.22631"
                Source = "Windows Update Agent"; Reason = "Inbox fallback"
            }
            WindowsServicing = [ordered]@{
                Status = "Ready"; Version = "10.0.22631"
                Source = "Windows inbox"; Reason = "DISM and SFC present"
            }
            Winget = [ordered]@{
                Status = $WingetStatus; Version = $(if ($WingetStatus -eq "Ready") { "1.29.280" } else { "" })
                Source = "Microsoft Desktop App Installer"; Reason = "WinGet state"
            }
            Dell = [ordered]@{
                Status = $DellStatus; Version = $(if ($DellStatus -eq "Ready") { "5.7.0" } else { "" })
                Source = "Dell Command Update"; Reason = "Dell state"
            }
            Lenovo = [ordered]@{
                Status = $LenovoStatus; Version = $(if ($LenovoStatus -eq "Ready") { "1.8.1" } else { "" })
                Source = "LSUClient"; Reason = "Lenovo state"
            }
            HP = [ordered]@{
                Status = $HPStatus; Version = $(if ($HPStatus -eq "Ready") { "5.3.6" } else { "" })
                Source = "HP Image Assistant"; Reason = "HP state"
            }
        }
    }
}

Describe "Dependency readiness, scoped packages, and rollout policy" {
    BeforeEach {
        Initialize-SystemUpdateProTestState
        Mock Write-Log {}
    }

    It "accepts only a matching content-addressed dependency cache artifact" {
        $payload = [Text.Encoding]::UTF8.GetBytes("verified dependency")
        $hash = ([BitConverter]::ToString(([Security.Cryptography.SHA256]::Create()).ComputeHash($payload)) -replace "-", "")
        $artifactDirectory = Join-Path $script:DependencyCachePath "sha256"
        New-Item -ItemType Directory -Path $artifactDirectory -Force | Out-Null
        $artifactPath = Join-Path $artifactDirectory $hash
        [IO.File]::WriteAllBytes($artifactPath, $payload)

        $evidence = Test-DependencyCacheArtifact -Name "Test dependency" -ExpectedSha256 $hash `
            -CachePath $script:DependencyCachePath

        $evidence.Valid | Should -BeTrue -Because $evidence.Reason
        $evidence.Status | Should -Be "Ready"
        $evidence.Path | Should -Be $artifactPath
    }

    It "never probes the network for an offline cache miss" {
        Mock Invoke-WebRequest { throw "network must not be contacted" }

        $readiness = Get-SourceReadiness -Name "Offline source" -Uri "https://example.test/package" `
            -AllowedHosts @("example.test") -ExpectedSha256 ("A" * 64) -OfflineMode $true

        $readiness.Status | Should -Be "Unavailable"
        $readiness.Ready | Should -BeFalse
        Should -Invoke Invoke-WebRequest -Times 0 -Exactly
    }

    It "records an approved source readiness response with its timeout" {
        Mock Invoke-WebRequest { [PSCustomObject]@{ StatusCode = 204 } }

        $readiness = Get-SourceReadiness -Name "Online source" -Uri "https://example.test/package" `
            -AllowedHosts @("example.test") -TimeoutSeconds 11

        $readiness.Status | Should -Be "Ready"
        $readiness.TimeoutSeconds | Should -Be 11
        $readiness.Host | Should -Be "example.test"
    }

    It "models SYSTEM scope without claiming unseen user-package success" {
        $plan = Get-WingetScopePlan -SystemInfo @{ ExecutionContext = "System" }

        @($plan.Scopes | Where-Object Scope -eq "machine").Status | Should -Be "Available"
        @($plan.Scopes | Where-Object Scope -eq "current-user").Status | Should -Be "Skipped"
        @($plan.Scopes | Where-Object Scope -eq "other-user").CanUpgrade | Should -BeFalse
        $plan.SystemClaimsPerUserSuccess | Should -BeFalse
    }

    It "applies wildcard exclusions, pins, and process conflicts without force-closing" {
        $policy = [PSCustomObject]@{
            Exclude = @("Contoso.*")
            Pins = [ordered]@{ "Fabrikam.App" = "2.0.0" }
            Conflicts = [ordered]@{
                "Tailspin.App" = [ordered]@{
                    processes = @("tailspin")
                    action = "close-with-deadline"
                    deadline_minutes = 10
                }
            }
        }
        $packages = @(
            @{ Id = "Contoso.Tool"; Name = "Contoso Tool"; Version = "1.0.0"; AvailableVersion = "2.0.0" },
            @{ Id = "Fabrikam.App"; Name = "Fabrikam"; Version = "1.5.0"; AvailableVersion = "2.1.0" },
            @{ Id = "Tailspin.App"; Name = "Tailspin"; Version = "1.0.0"; AvailableVersion = "1.1.0" },
            @{ Id = "Northwind.App"; Name = "Northwind"; Version = "1.0.0"; AvailableVersion = "1.1.0" }
        )

        $result = Get-WingetPackagePlan -Packages $packages -Policy $policy -RunningProcessNames @("tailspin") `
            -Scope "current-user"

        @($result.Items | Where-Object Id -eq "Contoso.Tool").Status | Should -Be "Excluded"
        @($result.Items | Where-Object Id -eq "Fabrikam.App").Status | Should -Be "Pinned"
        @($result.Items | Where-Object Id -eq "Tailspin.App").Status | Should -Be "Deferred"
        @($result.Items | Where-Object Id -eq "Northwind.App").Status | Should -Be "Eligible"
        @($result.Conflicts).Count | Should -Be 1
    }

    It "blocks known metered downloads unless an explicit override is recorded" {
        $metered = [PSCustomObject]@{ Status = "Metered"; CostType = "Fixed"; Known = $true; Reason = "fixed" }

        $blocked = Get-DownloadPolicy -NetworkCost $metered -DryRunMode $true
        $allowed = Get-DownloadPolicy -NetworkCost $metered -AllowOverride $true

        $blocked.Status | Should -Be "Blocked"
        $blocked.Allowed | Should -BeFalse
        $allowed.Allowed | Should -BeTrue
        $allowed.AuditedOverride | Should -BeTrue
    }

    It "returns deterministic rollout hold, halt, and promote decisions" {
        $policy = [PSCustomObject]@{
            Enabled = $true; PercentageStart = 0; PercentageEnd = 100
            MinimumSuccessCount = 2; MinimumSuccessRate = 0.9; BakeTimeHours = 0
            EmergencyOverride = $false; Cohort = "pilot"; Reason = "test"
        }
        $hold = Evaluate-RolloutPromotion -Policy $policy -Evidence @{ SuccessCount = 1; FailureCount = 0; LastSuccessAt = (Get-Date).ToUniversalTime().ToString("o") } -CohortValue 10
        $halt = Evaluate-RolloutPromotion -Policy $policy -Evidence @{ SuccessCount = 1; FailureCount = 2; LastSuccessAt = (Get-Date).ToUniversalTime().ToString("o") } -CohortValue 10
        $promote = Evaluate-RolloutPromotion -Policy $policy -Evidence @{ SuccessCount = 2; FailureCount = 0; LastSuccessAt = (Get-Date).ToUniversalTime().ToString("o") } -CohortValue 10

        $hold.Decision | Should -Be "Hold"
        $halt.Decision | Should -Be "Halt"
        $promote.Decision | Should -Be "Promote"
    }
}

Describe "Additional OEM and GPU provider plans" {
    BeforeEach {
        Initialize-SystemUpdateProTestState
        Mock Write-Log {}
    }

    It "selects the ASUS adapter for an ASUS system without claiming an uninstalled tool" {
        $plans = @(Get-AdditionalOEMProviderPlan -SystemInfo @{
            Manufacturer = "ASUSTeK COMPUTER INC."; Model = "ROG Zephyrus"
        })

        $plans.Provider | Should -Contain "ASUS"
        $asus = @($plans | Where-Object Provider -eq "ASUS")[0]
        $asus.Category | Should -Be "OEM"
        $asus.SourceUri | Should -Match "asus.com"
        $asus.Status | Should -BeIn @("Ready", "AcquisitionRequired")
    }

    It "maps GPU inventory to one plan per supported vendor" {
        $plans = @(Get-GPUProviderPlan -VideoControllers @(
            @{ Name = "NVIDIA GeForce RTX 4060" },
            @{ Name = "NVIDIA GeForce RTX 4060" },
            @{ Name = "Intel(R) UHD Graphics" }
        ))

        @($plans).Count | Should -Be 2
        @($plans | Where-Object Provider -eq "NVIDIA").Count | Should -Be 1
        @($plans | Where-Object Provider -eq "Intel").Count | Should -Be 1
        $plans.SourceUri | Should -Not -Contain "http://"
    }

    It "reports missing vendor tools as skipped and never invokes a network shell" {
        $script:DryRun = [switch]$true
        Mock Get-AdditionalOEMProviderPlan {
            @([PSCustomObject][ordered]@{
                Provider = "Framework"; Category = "OEM"; DisplayName = "Framework firmware"
                SourceUri = "https://knowledgebase.frame.work/"
                CandidatePaths = @(); ExecutablePath = ""; Arguments = @()
                Status = "AcquisitionRequired"; Applicable = $true; Reason = "No approved updater"
            })
        }

        $result = Invoke-AdditionalOEMUpdate -SystemInfo @{ Manufacturer = "Framework"; Model = "13" }

        $result.Success | Should -BeTrue
        $result.Skipped | Should -Be 1
        $result.Items[0].Status | Should -Be "Skipped"
        Should -Invoke Get-AdditionalOEMProviderPlan -Times 1 -Exactly
    }
}

Describe "Schema-versioned result contract" {
    BeforeEach {
        Initialize-SystemUpdateProTestState
    }

    It "emits the complete per-item evidence shape" {
        $item = New-UpdateItemResult -Name "KB5000001" -Id "update-id" -Status "Failed" `
            -ProviderCode 4 -HResult -2145124300 -RebootRequired $true -Evidence @("provider.log")

        foreach ($property in @(
            "Id", "Name", "Status", "Attempted", "Available", "Installed", "Failed", "Skipped",
            "ProviderExitCode", "HResult", "RebootRequired", "Message", "Evidence",
            "StartedAt", "DurationSeconds"
        )) {
            $item.PSObject.Properties.Name | Should -Contain $property
        }
        $item.Attempted | Should -BeTrue
        $item.Failed | Should -BeTrue
        $item.ProviderExitCode | Should -Be 4
        $item.HResult | Should -Be -2145124300
    }

    It "does not promote a zero-item provider failure to success" {
        $stage = ConvertTo-StageResult -Name "WindowsUpdate" -Provider "PSWindowsUpdate" -Result @{
            Success = $false
            ExitCode = 9
            HResult = -2145124300
            Message = "Provider invocation failed"
        }

        $stage.Status | Should -Be "Failed"
        $stage.Attempted | Should -Be 1
        $stage.Failed | Should -Be 1
        $stage.ProviderExitCode | Should -Be 9
        $stage.HResult | Should -Be -2145124300
        $stage.Items.Count | Should -Be 1
        $stage.Items[0].Status | Should -Be "Failed"
    }

    It "reports dry-run discoveries as available rather than installed" {
        $script:DryRun = [switch]$true

        $stage = ConvertTo-StageResult -Name "Winget" -Provider "WinGet" -Result @{
            Success = $true
            UpdateCount = 3
            Message = "3 upgrades available"
        } -ItemNames @("App A", "App B", "App C")

        $stage.Status | Should -Be "Succeeded"
        $stage.Attempted | Should -Be 0
        $stage.Available | Should -Be 3
        $stage.Installed | Should -Be 0
        @($stage.Items | Where-Object Status -eq "Available").Count | Should -Be 3
    }

    It "keeps totals and process exit semantics aligned" {
        [void]$script:StageResults.Add((New-StageResult -Name "Initialization" -Status "Succeeded" -Attempted 1))
        [void]$script:StageResults.Add((New-StageResult -Name "OEM" -Provider "Dell" -Status "Succeeded" -Attempted 1 -Installed 1))
        [void]$script:StageResults.Add((New-StageResult -Name "WindowsUpdate" -Provider "WUA" -Status "Failed" -Attempted 1 -Failed 1))

        $run = New-RunData -StartedAt (Get-Date).AddMinutes(-1) -CompletedAt (Get-Date)

        $run.SchemaVersion | Should -Be 1
        $run.Status | Should -Be "Partial"
        $run.ExitCode | Should -Be 2
        $run.TotalInstalled | Should -Be 1
        $run.TotalFailed | Should -Be 1
        $run.Stages.Count | Should -Be 3
    }

    It "treats a provider-only failure as critical despite successful initialization" {
        $stages = @(
            (New-StageResult -Name "Initialization" -Status "Succeeded" -Attempted 1),
            (New-StageResult -Name "WindowsUpdate" -Provider "WUA" -Status "Failed" -Attempted 1 -Failed 1)
        )

        Get-RunExitCode -Stages $stages | Should -Be 3
    }

    It "returns reboot-required only when no stage failed" {
        $stages = @(
            (New-StageResult -Name "WindowsUpdate" -Status "Succeeded" -Attempted 1 -Installed 1 -RebootRequired $true)
        )

        Get-RunExitCode -Stages $stages | Should -Be 1
    }
}

Describe "Platform and provider capability matrix" {
    BeforeEach {
        Initialize-SystemUpdateProTestState
    }

    It "derives provider version and architecture gates from the pinned acquisition manifest" {
        $manifest = Get-AcquisitionManifest
        $matrix = Get-CapabilityMatrix

        $matrix.SchemaVersion | Should -Be 1
        $matrix.Providers.Winget.MinimumVersion | Should -Be $manifest.WinGet.MinimumVersion
        $matrix.Providers.Dell.MinimumVersion | Should -Be $manifest.DellCommandUpdate.MinimumVersion
        $matrix.Providers.Dell.AdditionalMinimumVersions.InventoryCollector |
            Should -Be $manifest.DellCommandUpdate.InventoryCollectorMinimum
        $matrix.Providers.Lenovo.MinimumVersion | Should -Be $manifest.LSUClient.MinimumVersion
        $matrix.Providers.HP.AcquisitionVersion | Should -Be $manifest.HPIA.ExactVersion
        @($matrix.Providers.HP.Architectures) | Should -Contain "arm64"
    }

    It "approves a supported Windows 11 Dell client and records detected provider versions" {
        $systemInfo = New-CapabilitySystemInfo -Manufacturer "Dell Inc."
        $versions = New-CapabilityVersionInventory -DellStatus "Ready"

        $assessment = Get-CapabilityAssessment -SystemInfo $systemInfo -VersionInventory $versions

        $assessment.Platform.Status | Should -Be "Ready"
        $assessment.Providers.WindowsUpdate.Status | Should -Be "Ready"
        $assessment.Providers.Winget.Status | Should -Be "Ready"
        $assessment.Providers.Dell.Status | Should -Be "Ready"
        $assessment.Providers.Dell.DetectedVersion | Should -Be "5.7.0"
        $assessment.Providers.Dell.VersionPolicy | Should -Be "VerifiedAcquisition"
        $assessment.Providers.HP.Status | Should -Be "Unsupported"
        $script:CapabilityAssessment = $assessment
        (New-RunData -StartedAt (Get-Date).AddSeconds(-1)).Capabilities.Providers.Dell.DetectedVersion |
            Should -Be "5.7.0"
    }

    It "keeps inbox updating available but rejects WinGet below the Windows 10 build floor" {
        $systemInfo = New-CapabilitySystemInfo -OSBuild "17134"
        $versions = New-CapabilityVersionInventory

        $assessment = Get-CapabilityAssessment -SystemInfo $systemInfo -VersionInventory $versions

        $assessment.Platform.Status | Should -Be "Ready"
        $assessment.Providers.WindowsUpdate.Status | Should -Be "Ready"
        $assessment.Providers.WindowsServicing.Status | Should -Be "Ready"
        $assessment.Providers.Winget.Status | Should -Be "Unsupported"
        $assessment.Providers.Winget.Reason | Should -Match "17763"
    }

    It "allows inbox servicing on Server Core while explicitly blocking unsupported providers" {
        $systemInfo = New-CapabilitySystemInfo -OSBuild "20348" -ProductType 3 `
            -InstallationType "Server Core" -EditionID "ServerDatacenterCor" `
            -OperatingSystemSKU 13 -RunContext "System"
        $versions = New-CapabilityVersionInventory -WingetStatus "AcquisitionRequired"

        $assessment = Get-CapabilityAssessment -SystemInfo $systemInfo -VersionInventory $versions

        $assessment.Platform.Status | Should -Be "Ready"
        $assessment.Platform.IsServerCore | Should -BeTrue
        $assessment.Providers.WindowsUpdate.Status | Should -Be "Ready"
        $assessment.Providers.WindowsServicing.Status | Should -Be "Ready"
        $assessment.Providers.Winget.Status | Should -Be "Unsupported"
        $assessment.Providers.Winget.Reason | Should -Match "Windows Server build 26100|Server Core"
        $assessment.Providers.Dell.Status | Should -Be "Unsupported"
    }

    It "supports WinGet on Server 2025 Desktop Experience only in an administrator user context" {
        $versions = New-CapabilityVersionInventory
        $serverUser = New-CapabilitySystemInfo -OSBuild "26100" -ProductType 3 `
            -InstallationType "Server" -EditionID "ServerDatacenter" `
            -OperatingSystemSKU 8 -RunContext "AdministratorUser"
        $serverSystem = New-CapabilitySystemInfo -OSBuild "26100" -ProductType 3 `
            -InstallationType "Server" -EditionID "ServerDatacenter" `
            -OperatingSystemSKU 8 -RunContext "System"

        $userAssessment = Get-CapabilityAssessment -SystemInfo $serverUser -VersionInventory $versions
        $systemAssessment = Get-CapabilityAssessment -SystemInfo $serverSystem -VersionInventory $versions

        $userAssessment.Providers.Winget.Status | Should -Be "Ready"
        $systemAssessment.Providers.Winget.Status | Should -Be "Unsupported"
        $systemAssessment.Providers.Winget.Reason | Should -Match "System context"
    }

    It "permits verified HPIA acquisition on an ARM64 HP client" {
        $systemInfo = New-CapabilitySystemInfo -Manufacturer "HP Inc." -OSBuild "26100" `
            -Architecture "arm64"
        $versions = New-CapabilityVersionInventory -HPStatus "AcquisitionRequired"

        $assessment = Get-CapabilityAssessment -SystemInfo $systemInfo -VersionInventory $versions

        $assessment.Platform.Status | Should -Be "Ready"
        $assessment.Providers.HP.Status | Should -Be "RequiresAcquisition"
        $assessment.Providers.HP.Supported | Should -BeTrue
        $assessment.Providers.HP.AcquisitionVersion | Should -Be "5.3.6"
    }

    It "blocks Dell bootstrap under SYSTEM but allows an already verified Dell CLI" {
        $systemInfo = New-CapabilitySystemInfo -Manufacturer "Dell Inc." -RunContext "System"
        $missingVersions = New-CapabilityVersionInventory -DellStatus "AcquisitionRequired"
        $installedVersions = New-CapabilityVersionInventory -DellStatus "Ready"

        $missing = Get-CapabilityAssessment -SystemInfo $systemInfo -VersionInventory $missingVersions
        $installed = Get-CapabilityAssessment -SystemInfo $systemInfo -VersionInventory $installedVersions

        $missing.Providers.Dell.Status | Should -Be "Unsupported"
        $missing.Providers.Dell.Reason | Should -Match "cannot be bootstrapped"
        $installed.Providers.Dell.Status | Should -Be "Ready"
    }

    It "fails closed when edition, architecture, or execution context cannot be classified" {
        $systemInfo = New-CapabilitySystemInfo -EditionID "" -OperatingSystemSKU 0 `
            -Architecture "unknown" -RunContext "Unknown"
        $versions = New-CapabilityVersionInventory

        $assessment = Get-CapabilityAssessment -SystemInfo $systemInfo -VersionInventory $versions

        $assessment.Platform.Status | Should -Be "Unknown"
        $assessment.Platform.Supported | Should -BeFalse
        $assessment.Platform.Reason | Should -Match "edition"
        $assessment.Platform.Reason | Should -Match "architecture"
        $assessment.Platform.Reason | Should -Match "Execution context"
        $assessment.Providers.WindowsUpdate.Status | Should -Be "Unsupported"
    }

    It "builds a read-only installed-version inventory for the applicable provider" {
        Mock Test-InstalledPSModule {
            param($ModuleName)
            [PSCustomObject]@{
                Valid = $true
                Version = $(if ($ModuleName -eq "PSWindowsUpdate") { "2.2.1.5" } else { "1.8.1" })
                Reason = ""
            }
        }
        Mock Test-Path { $true }
        Mock Get-WingetTrustEvidence {
            [PSCustomObject]@{ Valid = $true; Version = "1.29.280"; Reason = "" }
        }
        Mock Get-InstalledExecutableEvidence {
            [PSCustomObject]@{ Valid = $true; Version = "5.7.0"; Reason = "" }
        }
        Mock Get-DellInventoryCollectorEvidence {
            [PSCustomObject]@{ Valid = $true; Version = "13.8.0"; Reason = "" }
        }

        $systemInfo = New-CapabilitySystemInfo -Manufacturer "Dell Inc."
        $inventory = Get-ProviderVersionInventory -SystemInfo $systemInfo

        $inventory.WindowsUpdate.Status | Should -Be "Ready"
        $inventory.WindowsServicing.Status | Should -Be "Ready"
        $inventory.Winget.Version | Should -Be "1.29.280"
        $inventory.Dell.Status | Should -Be "Ready"
        $inventory.Lenovo.Status | Should -Be "NotApplicable"
        Assert-MockCalled Get-InstalledExecutableEvidence -Times 1 -Exactly
    }
}

Describe "Verified elevated dependency acquisition" {
    BeforeEach {
        Initialize-SystemUpdateProTestState
    }

    It "defines a complete manifest contract for every elevated dependency" {
        $manifest = Get-AcquisitionManifest

        @($manifest.Keys).Count | Should -Be 5
        foreach ($name in @("WinGet", "PSWindowsUpdate", "LSUClient", "DellCommandUpdate", "HPIA")) {
            $entry = $manifest[$name]
            $entry.Uri | Should -Match "^https://"
            (Test-AcquisitionUri -Uri $entry.Uri -AllowedHosts $entry.AllowedHosts) | Should -BeTrue
            $entry.ExactVersion | Should -Not -BeNullOrEmpty
            $entry.MinimumVersion | Should -Not -BeNullOrEmpty
            @($entry.Architectures).Count | Should -BeGreaterThan 0
            $entry.Sha256 | Should -Match "^[A-F0-9]{64}$"
        }
    }

    It "pins remediated HP and Dell security floors" {
        $manifest = Get-AcquisitionManifest

        $manifest.HPIA.ExactVersion | Should -Be "5.3.6"
        (Test-VersionAtLeast -Version $manifest.HPIA.ExactVersion `
            -MinimumVersion $manifest.HPIA.MinimumVersion) | Should -BeTrue
        $manifest.HPIA.MinimumVersion | Should -Be "5.3.3"
        $manifest.DellCommandUpdate.ExactVersion | Should -Be "5.7.0"
        $manifest.DellCommandUpdate.InventoryCollectorMinimum | Should -Be "13.8.0"
    }

    It "requires the Dell WinGet source to match the pinned URI, digest, and publisher" {
        $spec = (Get-AcquisitionManifest).DellCommandUpdate
        $matching = @"
Publisher: Dell Inc.
Installer Url: $($spec.Uri)
Installer SHA256: $($spec.Sha256)
"@

        (Test-WingetPackageContractText -OutputText $matching -ExpectedUri $spec.Uri `
            -ExpectedSha256 $spec.Sha256 -ExpectedPublisher $spec.PublisherDisplayName) | Should -BeTrue
        (Test-WingetPackageContractText -OutputText ($matching -replace $spec.Sha256, ("0" * 64)) `
            -ExpectedUri $spec.Uri -ExpectedSha256 $spec.Sha256 `
            -ExpectedPublisher $spec.PublisherDisplayName) | Should -BeFalse
    }

    It "rejects non-HTTPS, credentialed, nonstandard-port, and unapproved origins" {
        $allowed = @("downloads.example.test")

        (Test-AcquisitionUri -Uri "https://downloads.example.test/file.exe" -AllowedHosts $allowed) | Should -BeTrue
        (Test-AcquisitionUri -Uri "http://downloads.example.test/file.exe" -AllowedHosts $allowed) | Should -BeFalse
        (Test-AcquisitionUri -Uri "https://user:pass@downloads.example.test/file.exe" -AllowedHosts $allowed) | Should -BeFalse
        (Test-AcquisitionUri -Uri "https://downloads.example.test:8443/file.exe" -AllowedHosts $allowed) | Should -BeFalse
        (Test-AcquisitionUri -Uri "https://redirect.example.test/file.exe" -AllowedHosts $allowed) | Should -BeFalse
    }

    It "rejects a tampered file before publisher verification" {
        $path = Join-Path $TestDrive "tampered.bin"
        Set-Content -LiteralPath $path -Value "tampered" -NoNewline
        Mock Get-AuthenticodeEvidence {
            [PSCustomObject]@{ Valid = $true; Subject = "CN=Approved"; Thumbprint = "ABC"; Reason = "" }
        }

        $result = Test-AcquiredFile -Path $path `
            -ExpectedSha256 ("0" * 64) -PublisherPattern "^CN=Approved"

        $result.Valid | Should -BeFalse
        $result.Reason | Should -Match "SHA-256 mismatch"
        Should -Invoke Get-AuthenticodeEvidence -Times 0 -Exactly
    }

    It "rejects a wrong publisher after a matching digest" {
        $path = Join-Path $TestDrive "wrong-publisher.exe"
        Set-Content -LiteralPath $path -Value "verified bytes" -NoNewline
        $hash = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
        Mock Get-AuthenticodeEvidence {
            [PSCustomObject]@{
                Valid = $false; Subject = "CN=Unexpected Corp"; Thumbprint = "BAD"
                Reason = "Signer does not match the approved publisher"
            }
        }

        $result = Test-AcquiredFile -Path $path -ExpectedSha256 $hash `
            -PublisherPattern "^CN=Approved Corp,"

        $result.Valid | Should -BeFalse
        $result.Reason | Should -Match "approved publisher"
    }

    It "rejects an executable below its approved minimum version" {
        Mock Test-Path { $true }
        Mock Get-Item {
            [PSCustomObject]@{
                VersionInfo = [PSCustomObject]@{ ProductVersion = "13.7.9"; FileVersion = "13.7.9" }
            }
        }

        $result = Get-InstalledExecutableEvidence -Path "C:\unsafe\invcol.exe" `
            -MinimumVersion "13.8.0" -PublisherPattern "^CN=Dell Inc\.,"

        $result.Valid | Should -BeFalse
        $result.Reason | Should -Match "below the approved minimum"
    }

    It "rejects a changed installed module payload" {
        $modulePath = Get-PSModuleInstallPath -ModuleName "LSUClient" -Version "1.8.1"
        New-Item -ItemType Directory -Path $modulePath -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $modulePath "LSUClient.psd1") -Value @"
@{
    RootModule = 'LSUClient.psm1'
    ModuleVersion = '1.8.1'
}
"@
        Set-Content -LiteralPath (Join-Path $modulePath "LSUClient.psm1") -Value "# changed payload"

        $result = Test-InstalledPSModule -ModuleName "LSUClient"

        $result.Valid | Should -BeFalse
        $result.Reason | Should -Match "payload hash"
    }

    It "never changes PowerShell Gallery trust or skips publisher checks" {
        $source = Get-Content -LiteralPath $script:SourceScriptPath -Raw

        $source | Should -Not -Match "Set-PSRepository"
        $source | Should -Not -Match "SkipPublisherCheck"
        $source | Should -Not -Match "(?m)^\s*Install-Module\b"
    }

    It "carries verified dependency provenance into terminal run data" {
        [void](Add-AcquisitionProvenance -Name "LSUClient" -Version "1.8.1" `
            -SourceUri "https://www.powershellgallery.com/api/v2/package/LSUClient/1.8.1" `
            -Sha256 ("A" * 64) -Architecture "neutral" -InstallPath "C:\Modules\LSUClient\1.8.1" `
            -Status "Installed")

        $run = New-RunData -StartedAt (Get-Date).AddSeconds(-1)

        $run.Dependencies.Count | Should -Be 1
        $run.Dependencies[0].Name | Should -Be "LSUClient"
        $run.Dependencies[0].ManifestVersion | Should -Be 1
    }

    It "renders installed provenance in the operator report" {
        $script:LogFile = Join-Path $TestDrive "run.log"
        Mock Start-Process {}
        [void](Add-AcquisitionProvenance -Name "HP Image Assistant" -Version "5.3.6.649" `
            -SourceUri "https://hpia.hpcloud.hp.com/downloads/hpia/hp-hpia-5.3.6.exe" `
            -Sha256 ("B" * 64) -Publisher "CN=HP Inc." -Architecture "x64" `
            -InstallPath "C:\ProgramData\SystemUpdatePro\HPIA\HPImageAssistant.exe")
        $capabilitySystem = New-CapabilitySystemInfo -Manufacturer "HP Inc."
        $script:CapabilityAssessment = Get-CapabilityAssessment -SystemInfo $capabilitySystem `
            -VersionInventory (New-CapabilityVersionInventory -HPStatus "Ready")
        $script:WebhookUrl = "https://example.test/workflow?sig=REPORT-SECRET"
        [void]$script:Errors.Add(
            "Webhook https://example.test/workflow?sig=REPORT-SECRET failed"
        )
        $run = New-RunData -StartedAt (Get-Date).AddSeconds(-1)
        $system = @{
            Manufacturer = "HP"; Model = "EliteBook"; SerialNumber = "REPORT-SERIAL-123"
            OSName = "Windows"; OSBuild = "1"; BIOSVersion = "1"; BIOSDate = Get-Date
            Processor = "CPU"; TotalRAM = 16; InstallationType = "Client"
            Architecture = "x64"; ExecutionContext = "AdministratorUser"
        }

        $reportPath = New-HTMLReport -SysInfo $system -RunData $run
        $report = Get-Content -LiteralPath $reportPath -Raw

        $report | Should -Match "Dependency provenance"
        $report | Should -Match "HP Image Assistant"
        $report | Should -Match "5\.3\.6\.649"
        $report | Should -Match "C:\\ProgramData\\SystemUpdatePro\\HPIA"
        $report | Should -Match "Platform capability"
        $report | Should -Match "Provider capability"
        $report | Should -Match "AdministratorUser"
        $report | Should -Not -Match "REPORT-SECRET|REPORT-SERIAL-123"
        $report | Should -Match "REDACTED"
    }
}

Describe "Protected local evidence store" {
    BeforeEach {
        Initialize-SystemUpdateProTestState
    }

    It "atomically replaces structured evidence and preserves a protected last-known-good copy" {
        $path = Join-Path $script:DataPath "atomic.json"
        $validate = {
            param($data)
            [PSCustomObject]@{
                Valid = ($data.value -in @("first", "second"))
                Reason = "unexpected value"
            }
        }

        Write-ProtectedAtomicJson -Path $path -Data ([ordered]@{ value = "first" }) `
            -DataValidationScript $validate | Should -BeTrue
        Write-ProtectedAtomicJson -Path $path -Data ([ordered]@{ value = "second" }) `
            -DataValidationScript $validate | Should -BeTrue

        ((Get-Content -LiteralPath $path -Raw | ConvertFrom-Json).value) | Should -Be "second"
        ((Get-Content -LiteralPath "$path.previous" -Raw | ConvertFrom-Json).value) |
            Should -Be "first"
        (Test-EvidencePathAccess -Path $path).Valid | Should -BeTrue
        (Test-EvidencePathAccess -Path "$path.previous").Valid | Should -BeTrue
        $writerSids = @((Get-Acl -LiteralPath $path).Access | Where-Object {
            $_.AccessControlType -eq [Security.AccessControl.AccessControlType]::Allow -and
            ($_.FileSystemRights -band [Security.AccessControl.FileSystemRights]::Write)
        } | ForEach-Object {
            $_.IdentityReference.Translate([Security.Principal.SecurityIdentifier]).Value
        })
        $writerSids | Should -Contain "S-1-5-18"
        $writerSids | Should -Contain "S-1-5-32-544"
        @(Get-ChildItem -LiteralPath $script:DataPath -Filter "atomic.json.tmp.*").Count |
            Should -Be 0

        Write-ProtectedAtomicJson -Path $path -Data ([ordered]@{ value = "rejected" }) `
            -DataValidationScript $validate | Should -BeFalse
        ((Get-Content -LiteralPath $path -Raw | ConvertFrom-Json).value) | Should -Be "second"
    }

    It "quarantines a corrupt primary and restores its last-known-good JSON" {
        $path = Join-Path $script:DataPath "recoverable.json"
        $validate = {
            param($data)
            [PSCustomObject]@{ Valid = ($data.Contains("value")); Reason = "value missing" }
        }
        Write-ProtectedAtomicJson -Path $path -Data ([ordered]@{ value = "known-good" }) `
            -DataValidationScript $validate | Should -BeTrue
        Write-ProtectedAtomicJson -Path $path -Data ([ordered]@{ value = "current" }) `
            -DataValidationScript $validate | Should -BeTrue
        [IO.File]::WriteAllText($path, "{broken")

        $read = Read-ProtectedJsonFile -Path $path -ValidationScript $validate

        $read.Success | Should -BeTrue
        $read.Recovered | Should -BeTrue
        $read.Data.value | Should -Be "known-good"
        ((Get-Content -LiteralPath $path -Raw | ConvertFrom-Json).value) |
            Should -Be "known-good"
        @(Get-ChildItem -LiteralPath $script:DataPath -Filter "recoverable.corrupt.*.json").Count |
            Should -Be 1
    }

    It "migrates a v3 continuation state through v5 with evidence and secret-source defaults" {
        $legacy = New-ContinuationState -StageCursor "WindowsUpdate" `
            -ScriptPath $script:SourceScriptPath
        $legacy.SchemaVersion = 3
        [void]$legacy.Parameters.Remove("EvidenceMaxSizeMB")
        [void]$legacy.Parameters.Remove("RedactionMode")
        [void]$legacy.Parameters.Remove("WebhookSecretReference")
        $legacy.Parameters["WebhookUrl"] = "https://example.invalid/legacy-secret"
        [void]$legacy.Remove("ResolvedWebhookUrl")
        Write-ProtectedAtomicJson -Path $script:StateFile -Data $legacy | Should -BeTrue

        $loaded = Get-State

        $loaded.SchemaVersion | Should -Be 5
        $loaded.Parameters.EvidenceMaxSizeMB | Should -Be 512
        $loaded.Parameters.RedactionMode | Should -Be "SecretsAndSerials"
        $loaded.Parameters.WebhookSecretReference | Should -Be ""
        $loaded.ResolvedWebhookUrl | Should -Be "https://example.invalid/legacy-secret"
        $loaded.Contains("_MigrationSourceSchema") | Should -BeFalse
        ((Get-Content -LiteralPath $script:StateFile -Raw | ConvertFrom-Json).SchemaVersion) |
            Should -Be 5
    }

    It "migrates and redacts legacy history without leaving sensitive backup data" {
        $script:WebhookUrl = "https://hooks.slack.com/services/T00/B00/SECRET"
        $legacyEntry = [ordered]@{
            run_id = [guid]::NewGuid().ToString()
            timestamp = (Get-Date).ToString("o")
            exit_code = 0
            serial_number = "SERIAL-123"
            errors = @("delivery https://hooks.slack.com/services/T00/B00/SECRET failed")
        }
        $legacyJson = ConvertTo-Json -InputObject @($legacyEntry) -Depth 8
        Write-ProtectedAtomicFile -Path $script:HistoryFile -Content $legacyJson |
            Should -BeTrue
        Add-SensitiveEvidenceValue -Value "SERIAL-123"

        $history = Read-UpdateHistory
        $backup = Get-Content -LiteralPath "$($script:HistoryFile).previous" -Raw

        $history.schema_version | Should -Be 2
        $history.entries.Count | Should -Be 1
        $history.entries[0].serial_number | Should -Be "[REDACTED]"
        $history.entries[0].errors[0] | Should -Not -Match "SECRET"
        $backup | Should -Not -Match "SERIAL-123|SECRET"
        (Test-HistoryDocument -History (
            ConvertTo-Hashtable -InputObject ($backup | ConvertFrom-Json)
        )).Valid | Should -BeTrue
    }

    It "redacts flushed logs and text artifacts without retaining an unredacted copy" {
        $script:WebhookUrl = "https://example.test/hook?sig=TOPSECRET"
        Add-SensitiveEvidenceValue -Value "SERIAL-ABC"
        $logPath = Join-Path $script:DataPath "SystemUpdatePro_redaction.log"

        Add-ProtectedEvidenceLine -Path $logPath `
            -Line "Serial SERIAL-ABC endpoint https://example.test/hook?sig=TOPSECRET" |
            Should -BeTrue
        $content = Get-Content -LiteralPath $logPath -Raw
        $content | Should -Not -Match "SERIAL-ABC|TOPSECRET"
        $content | Should -Match "REDACTED"
        (Test-EvidencePathAccess -Path $logPath).Valid | Should -BeTrue

        $script:RedactionMode = "Secrets"
        Protect-EvidenceText -Text "SERIAL-ABC https://example.test/hook?sig=TOPSECRET" |
            Should -Match "SERIAL-ABC"
    }

    It "removes only owned expired artifacts and reports exact file, directory, and byte counts" {
        $oldLog = Join-Path $script:LogPath "SystemUpdatePro_20000101_000000.log"
        $oldReport = Join-Path $script:LogPath "SystemUpdatePro_Report_20000101_000000.html"
        $vendorDirectory = Join-Path $script:LogPath "HPIA_20000101_000000"
        [void](Write-ProtectedAtomicFile -Path $oldLog -Content "12345")
        [void](Write-ProtectedAtomicFile -Path $oldReport -Content "1234567")
        [void](New-ProtectedDirectory -Path $vendorDirectory)
        [void](Write-ProtectedAtomicFile -Path (Join-Path $vendorDirectory "one.log") -Content "123")
        [void](Write-ProtectedAtomicFile -Path (Join-Path $vendorDirectory "two.xml") -Content "1234")
        $bundleDirectory = Join-Path $script:DataPath "Bundles"
        [void](New-ProtectedDirectory -Path $bundleDirectory)
        $oldBundle = Join-Path $bundleDirectory `
            "SystemUpdatePro_Diagnostic_20000101_000000_00000000.zip"
        [void](Write-ProtectedAtomicFile -Path $oldBundle -Content "123456")
        [void](New-ProtectedDirectory -Path $script:WebhookDeliveryDirectory)
        $oldDelivery = Join-Path $script:WebhookDeliveryDirectory "old-run.json"
        [void](Write-ProtectedAtomicFile -Path $oldDelivery -Content "123")
        [void](Write-ProtectedAtomicFile -Path $oldDelivery -Content "1234")
        $unowned = Join-Path $script:LogPath "operator-notes.log"
        [IO.File]::WriteAllText($unowned, "keep")

        foreach ($path in @(
            $oldLog,
            $oldReport,
            $vendorDirectory,
            $oldBundle,
            $oldDelivery,
            "$oldDelivery.previous"
        )) {
            (Get-Item -LiteralPath $path).LastWriteTime = (Get-Date).AddDays(-60)
        }
        $result = Invoke-EvidenceRetention -RetentionDays 30 -MaximumSizeMB 512

        $result.DeletedFiles | Should -Be 7
        $result.DeletedDirectories | Should -Be 1
        $result.BytesFreed | Should -Be 32
        $result.Errors.Count | Should -Be 0
        Test-Path -LiteralPath $unowned | Should -BeTrue
        Test-Path -LiteralPath $oldLog | Should -BeFalse
        Test-Path -LiteralPath $vendorDirectory | Should -BeFalse
        Test-Path -LiteralPath $oldBundle | Should -BeFalse
        Test-Path -LiteralPath $oldDelivery | Should -BeFalse
        Test-Path -LiteralPath "$oldDelivery.previous" | Should -BeFalse
    }

    It "enforces the configured evidence size ceiling by deleting oldest owned artifacts first" {
        $oldLargeLog = Join-Path $script:LogPath "SystemUpdatePro_20000101_000001.log"
        $newLargeLog = Join-Path $script:LogPath "SystemUpdatePro_20990101_000001.log"
        [IO.File]::WriteAllBytes($oldLargeLog, (New-Object byte[] (7MB)))
        [IO.File]::WriteAllBytes($newLargeLog, (New-Object byte[] (5MB)))
        [void](Set-EvidencePathAccess -Path $oldLargeLog)
        [void](Set-EvidencePathAccess -Path $newLargeLog)
        (Get-Item -LiteralPath $oldLargeLog).LastWriteTime = (Get-Date).AddDays(-1)

        $result = Invoke-EvidenceRetention -RetentionDays 30 -MaximumSizeMB 10

        $result.DeletedFiles | Should -Be 1
        $result.BytesFreed | Should -Be (7MB)
        $result.RemainingBytes | Should -BeLessOrEqual (10MB)
        Test-Path -LiteralPath $oldLargeLog | Should -BeFalse
        Test-Path -LiteralPath $newLargeLog | Should -BeTrue
    }
}

Describe "Redacted diagnostic and recovery bundle" {
    BeforeEach {
        Initialize-SystemUpdateProTestState
        Mock Get-DiagnosticRuntimeSnapshot {
            param($FallbackCapabilities, $Dependencies, $Errors)
            Add-SensitiveEvidenceValue -Value "SERIAL-DIAGNOSTIC-123"
            return [ordered]@{
                schema_version = 1
                product = "SystemUpdatePro"
                product_version = "4.1.0"
                system = @{
                    serial_number = "SERIAL-DIAGNOSTIC-123"
                    os_build = "26100"
                }
                capabilities = $FallbackCapabilities
                dependency_provenance = @($Dependencies)
            }
        }
        Mock Get-DiagnosticWindowsEvidence {
            param($RawDirectory, $SinceDays, $MaximumEvents)
            return [PSCustomObject]@{
                SchemaVersion = 1
                SinceDays = 7
                MaximumEventsPerChannel = 250
                Events = [ordered]@{
                    "Microsoft-Windows-WindowsUpdateClient/Operational" = @(
                        @{
                            id = 20
                            message = "Install failed with 0x80240022"
                        }
                    )
                }
                WindowsUpdateLogCommand = [PSCustomObject]@{
                    Success = $true
                    ExitCode = 0
                    StandardOutput = ""
                    StandardError = ""
                }
                Files = @()
                Errors = @()
            }
        }
    }

    It "captures a failed run, policy, recovery, Windows, and provider evidence without secrets" {
        $runId = [guid]::NewGuid().ToString()
        $secretEndpoint = "https://user:PRIVATE-PASSWORD@example.test/hook?sig=PRIVATE-SIGNATURE"
        $history = [ordered]@{
            schema_version = 2
            last_updated_at = (Get-Date).ToUniversalTime().ToString("o")
            entries = @(
                [ordered]@{
                    schema_version = 1
                    run_id = $runId
                    timestamp = (Get-Date).ToString("o")
                    status = "Failed"
                    exit_code = 3
                    errors = @("Provider failed at $secretEndpoint")
                    warnings = @()
                    stages = @()
                    dependencies = @(@{ Name = "PSWindowsUpdate"; Version = "2.2.1.5" })
                    mutation_recovery = @(
                        @{ Action = "Restored"; Target = "wuauserv"; Detail = "verified" }
                    )
                    capabilities = @{ Platform = @{ Status = "Ready"; OSBuild = "26100" } }
                    parameters = @{ WebhookUrl = $secretEndpoint; IncludeBIOS = $false }
                }
            )
        }
        Write-ProtectedAtomicJson -Path $script:HistoryFile -Data $history -Depth 20 |
            Should -BeTrue

        $providerLog = Join-Path $script:LogPath "DCU_Apply_20260729_000000.log"
        [IO.File]::WriteAllText(
            $providerLog,
            "Serial SERIAL-DIAGNOSTIC-123 endpoint $secretEndpoint failed"
        )
        [void](Set-EvidencePathAccess -Path $providerLog)
        Mock Get-DiagnosticEvidenceFile {
            @(
                [PSCustomObject]@{
                    Path = $providerLog
                    RelativePath = "evidence/dell/DCU_Apply_20260729_000000.log"
                    Category = "Dell"
                    MaximumBytes = 4194304
                }
            )
        }
        $script:RedactionMode = "Secrets"

        $bundle = New-DiagnosticBundle -MaximumSizeMB 5

        $bundle.Success | Should -BeTrue -Because $bundle.Error
        $script:RedactionMode | Should -Be "Secrets"
        $bundle.Bytes | Should -BeLessOrEqual (5MB)
        Test-Path -LiteralPath $bundle.Path | Should -BeTrue
        (Test-EvidencePathAccess -Path $bundle.Path).Valid | Should -BeTrue

        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $archive = [IO.Compression.ZipFile]::OpenRead($bundle.Path)
        try {
            $entryNames = @($archive.Entries | ForEach-Object { $_.FullName })
            $entryNames | Should -Contain "manifest.json"
            $entryNames | Should -Contain "run/latest_run.json"
            $entryNames | Should -Contain "run/policy.json"
            $entryNames | Should -Contain "inventory/runtime.json"
            $entryNames | Should -Contain "recovery/status.json"
            $entryNames | Should -Contain "windows/events.json"
            $entryNames | Should -Contain "evidence/dell/DCU_Apply_20260729_000000.log"

            $entryText = [System.Collections.ArrayList]::new()
            foreach ($entry in $archive.Entries) {
                $reader = New-Object IO.StreamReader($entry.Open())
                try {
                    [void]$entryText.Add($reader.ReadToEnd())
                } finally {
                    $reader.Dispose()
                }
            }
            $combinedText = $entryText -join "`n"
            $combinedText | Should -Not -Match "SERIAL-DIAGNOSTIC-123"
            $combinedText | Should -Not -Match "PRIVATE-PASSWORD|PRIVATE-SIGNATURE"
            $combinedText | Should -Match "REDACTED"

            $manifestEntry = @($archive.Entries | Where-Object FullName -eq "manifest.json")[0]
            $manifestReader = New-Object IO.StreamReader($manifestEntry.Open())
            try {
                $manifest = $manifestReader.ReadToEnd() | ConvertFrom-Json
            } finally {
                $manifestReader.Dispose()
            }
            $manifest.schema_version | Should -Be 1
            $manifest.latest_run.run_id | Should -Be $runId
            $manifest.latest_run.status | Should -Be "Failed"
            @($manifest.files).Count | Should -BeGreaterThan 5
            foreach ($fileRecord in @($manifest.files)) {
                $fileRecord.sha256 | Should -Match "^[a-f0-9]{64}$"
            }
        } finally {
            $archive.Dispose()
        }
        @(Get-ChildItem -LiteralPath (Join-Path $script:DataPath "Bundles") `
            -Directory -Force -Filter ".staging_*").Count | Should -Be 0
    }

    It "keeps large evidence inside the requested archive ceiling and records truncation" {
        $largeLog = Join-Path $script:LogPath "SystemUpdatePro_20260729_000000.log"
        [IO.File]::WriteAllText($largeLog, ("A" * (6MB)))
        [void](Set-EvidencePathAccess -Path $largeLog)
        Mock Get-DiagnosticEvidenceFile {
            @(
                [PSCustomObject]@{
                    Path = $largeLog
                    RelativePath = "evidence/runlog/large.log"
                    Category = "RunLog"
                    MaximumBytes = 16777216
                }
            )
        }

        $bundle = New-DiagnosticBundle -MaximumSizeMB 5

        $bundle.Success | Should -BeTrue -Because $bundle.Error
        $bundle.Bytes | Should -BeLessOrEqual (5MB)
        $archive = [IO.Compression.ZipFile]::OpenRead($bundle.Path)
        try {
            $manifestEntry = @($archive.Entries | Where-Object FullName -eq "manifest.json")[0]
            $reader = New-Object IO.StreamReader($manifestEntry.Open())
            try {
                $manifest = $reader.ReadToEnd() | ConvertFrom-Json
            } finally {
                $reader.Dispose()
            }
            $largeRecord = @($manifest.files | Where-Object {
                $_.path -eq "evidence/runlog/large.log"
            })[0]
            $largeRecord.truncated | Should -BeTrue
            $largeRecord.original_bytes | Should -BeGreaterThan (5MB)
        } finally {
            $archive.Dispose()
        }
    }

    It "runs diagnostic commands without a shell and returns bounded structured output" {
        $result = Invoke-CapturedCommand -FilePath $env:ComSpec `
            -ArgumentList @("/d", "/c", "echo", "captured output") `
            -TimeoutSeconds 10 -MaximumOutputCharacters 4096

        $result.Success | Should -BeTrue
        $result.ExitCode | Should -Be 0
        $result.TimedOut | Should -BeFalse
        $result.StandardOutput | Should -Match "captured output"
        $result.Command | Should -Be "cmd.exe"
    }

    It "persists the complete redacted run policy for later failed-run collection" {
        $script:WebhookUrl = "https://example.test/workflow?sig=DO-NOT-PERSIST"
        $script:WebhookSecretReference = "env:SYSTEMUPDATEPRO_WEBHOOK_URL"
        $script:EvidenceMaxSizeMB = 768
        $runData = New-RunData -StartedAt (Get-Date).AddSeconds(-1)

        Save-UpdateHistory -RunData $runData | Should -BeTrue
        $history = Read-UpdateHistory
        $policy = @($history.entries)[0].parameters

        foreach ($parameterName in Get-ContinuationParameterName) {
            $policy.Contains($parameterName) | Should -BeTrue
        }
        $policy.EvidenceMaxSizeMB | Should -Be 768
        $policy.WebhookSecretReference | Should -Be "env:SYSTEMUPDATEPRO_WEBHOOK_URL"
        $policy.Contains("WebhookUrl") | Should -BeFalse
        (Get-Content -LiteralPath $script:HistoryFile -Raw) |
            Should -Not -Match "DO-NOT-PERSIST"
    }
}

Describe "Input validation and webhook secret references" {
    BeforeEach {
        Initialize-SystemUpdateProTestState
    }

    It "resolves an HTTPS endpoint from an environment reference and registers it for redaction" {
        $variableName = "SYSTEMUPDATEPRO_TEST_WEBHOOK_$([guid]::NewGuid().ToString('N'))"
        $endpoint = "https://example.test/workflow?sig=ENVIRONMENT-SECRET"
        try {
            [Environment]::SetEnvironmentVariable(
                $variableName,
                $endpoint,
                [EnvironmentVariableTarget]::Process
            )

            $resolved = Resolve-WebhookSecretReference -Reference "env:$variableName"

            $resolved.Success | Should -BeTrue
            $resolved.Source | Should -Be "Environment"
            $resolved.Url | Should -Be $endpoint
            Protect-EvidenceText -Text "Delivery failed for $endpoint" |
                Should -Not -Match "ENVIRONMENT-SECRET"
        } finally {
            [Environment]::SetEnvironmentVariable(
                $variableName,
                $null,
                [EnvironmentVariableTarget]::Process
            )
        }
    }

    It "accepts only a protected schema-versioned config containing an HTTPS endpoint" {
        $configPath = Join-Path $script:DataPath "webhook.json"
        $endpoint = "https://example.test/hook?token=FILE-SECRET"
        Write-ProtectedAtomicJson -Path $configPath -Data ([ordered]@{
            schema_version = 1
            webhook_url = $endpoint
        }) | Should -BeTrue

        $resolved = Resolve-WebhookSecretReference -Reference "file:$configPath"

        $resolved.Success | Should -BeTrue
        $resolved.Source | Should -Be "ProtectedFile"
        $resolved.Url | Should -Be $endpoint

        $unsafeConfig = Join-Path $script:DataPath "unsafe-webhook.json"
        [IO.File]::WriteAllText(
            $unsafeConfig,
            '{"schema_version":1,"webhook_url":"https://example.test/hook"}'
        )
        $acl = Get-Acl -LiteralPath $unsafeConfig
        $acl.SetAccessRuleProtection($false, $true)
        Set-Acl -LiteralPath $unsafeConfig -AclObject $acl

        $unsafe = Resolve-WebhookSecretReference -Reference "file:$unsafeConfig"

        $unsafe.Success | Should -BeFalse
        $unsafe.Error | Should -Match "ACL"
    }

    It "rejects insecure endpoints without echoing their secret value" {
        $variableName = "SYSTEMUPDATEPRO_TEST_WEBHOOK_$([guid]::NewGuid().ToString('N'))"
        $endpoint = "http://example.test/hook?sig=INSECURE-SECRET"
        try {
            [Environment]::SetEnvironmentVariable(
                $variableName,
                $endpoint,
                [EnvironmentVariableTarget]::Process
            )

            $resolved = Resolve-WebhookSecretReference -Reference "env:$variableName"

            $resolved.Success | Should -BeFalse
            $resolved.Error | Should -Match "HTTPS"
            $resolved.Error | Should -Not -Match "INSECURE-SECRET"
        } finally {
            [Environment]::SetEnvironmentVariable(
                $variableName,
                $null,
                [EnvironmentVariableTarget]::Process
            )
        }
    }

    It "rejects invalid ranges, paths, and raw webhook arguments before initialization" {
        $hostExecutable = if ($PSVersionTable.PSEdition -eq "Core") {
            Join-Path $PSHOME "pwsh.exe"
        } else {
            Join-Path $PSHOME "powershell.exe"
        }
        $invalidArguments = @(
            @("-MaxRetries", "0"),
            @("-LogPath", "relative\logs"),
            @("-WebhookSecretReference", "bad-reference"),
            @("-WebhookUrl", "https://example.invalid/hook")
        )
        foreach ($arguments in $invalidArguments) {
            $result = Invoke-CapturedCommand -FilePath $hostExecutable `
                -ArgumentList (@(
                    "-NoLogo", "-NoProfile", "-NonInteractive",
                    "-File", $script:SourceScriptPath
                ) + $arguments) -TimeoutSeconds 20

            $result.Success | Should -BeFalse
            $result.ExitCode | Should -Not -Be 0
            "$($result.StandardOutput)`n$($result.StandardError)" |
                Should -Not -Match "requires administrator privileges"
        }
    }

    It "keeps history as a read-only early command when its ACL permits access" {
        $history = [ordered]@{
            schema_version = 2
            last_updated_at = (Get-Date).ToUniversalTime().ToString("o")
            entries = @()
        }
        Write-ProtectedAtomicJson -Path $script:HistoryFile -Data $history |
            Should -BeTrue
        $hostExecutable = if ($PSVersionTable.PSEdition -eq "Core") {
            Join-Path $PSHOME "pwsh.exe"
        } else {
            Join-Path $PSHOME "powershell.exe"
        }

        $result = Invoke-CapturedCommand -FilePath $hostExecutable -ArgumentList @(
            "-NoLogo", "-NoProfile", "-NonInteractive",
            "-File", $script:SourceScriptPath,
            "-ShowHistory"
        ) -TimeoutSeconds 20

        $result.Success | Should -BeTrue
        $result.ExitCode | Should -Be 0
        "$($result.StandardOutput)`n$($result.StandardError)" |
            Should -Not -Match "requires administrator privileges"
    }
}

Describe "PowerShell 5.1 continuation state machine" {
    BeforeEach {
        Initialize-SystemUpdateProTestState
        Mock Write-Log {}
        $script:TaskActionStub = New-ScheduledTaskAction -Execute "powershell.exe"
        $script:TaskTriggerStub = New-ScheduledTaskTrigger -AtStartup
        $script:TaskPrincipalStub = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
        $script:TaskSettingsStub = New-ScheduledTaskSettingsSet
    }

    It "atomically round-trips a complete versioned state file" {
        $script:SkipOEM = [switch]$true
        $script:IncludeBIOS = [switch]$true
        $script:WebhookSecretReference = "file:C:\ProgramData\SystemUpdatePro\webhook.json"
        $script:WebhookUrl = "https://example.invalid/private-hook"
        [void]$script:StageResults.Add((New-StageResult -Name "OEM" -Status "Succeeded" -Attempted 1 -Installed 1))
        [void](Add-AcquisitionProvenance -Name "LSUClient" -Version "1.8.1" `
            -SourceUri "https://www.powershellgallery.com/api/v2/package/LSUClient/1.8.1" `
            -Sha256 ("A" * 64) -Architecture "neutral" -InstallPath "C:\Modules\LSUClient\1.8.1")
        $state = New-ContinuationState -StageCursor "WindowsUpdate" -ScriptPath $script:SourceScriptPath

        $firstSave = Save-State -State $state
        if (-not $firstSave) { throw "First atomic save failed: $script:LastStateError" }
        $state.Phase = "AwaitingReboot"
        $secondSave = Save-State -State $state
        if (-not $secondSave) { throw "Second atomic save failed: $script:LastStateError" }
        $loaded = Get-State

        $loaded.SchemaVersion | Should -Be 5
        $loaded.Phase | Should -Be "AwaitingReboot"
        $loaded.RunId | Should -Be "11111111-1111-1111-1111-111111111111"
        $loaded.Parameters.Count | Should -Be 30
        $loaded.AcquisitionProvenance.Count | Should -Be 1
        $loaded.AcquisitionProvenance[0].Name | Should -Be "LSUClient"
        $loaded.Parameters.SkipOEM | Should -BeTrue
        $loaded.Parameters.IncludeBIOS | Should -BeTrue
        $loaded.Parameters.WebhookSecretReference |
            Should -Be "file:C:\ProgramData\SystemUpdatePro\webhook.json"
        $loaded.ResolvedWebhookUrl | Should -Be "https://example.invalid/private-hook"
        $loaded.Parameters.MinFirmwareChargePercent | Should -Be 50
        $loaded.StageResults.Count | Should -Be 1
        (Test-ContinuationStateAccess -Path $script:StateFile).Valid | Should -BeTrue
        (Get-Acl -LiteralPath $script:StateFile).AreAccessRulesProtected | Should -BeTrue
        @(Get-ChildItem -LiteralPath (Split-Path -Parent $script:StateFile) -Filter "state.json.tmp.*").Count | Should -Be 0
        Test-Path -LiteralPath "$($script:StateFile).previous" | Should -BeTrue
        (Test-ContinuationStateAccess -Path "$($script:StateFile).previous").Valid | Should -BeTrue
    }

    It "restores every effective parameter, run identity, cursor, and bounded attempt" {
        $script:SkipOEM = [switch]$true
        $script:SkipWindows = [switch]$true
        $script:IncludeBIOS = [switch]$true
        $script:BypassWSUS = [switch]$true
        $script:RepairWindowsUpdate = [switch]$true
        $script:CleanupAfter = [switch]$true
        $script:ResetComponentBase = [switch]$true
        $script:BackupDrivers = [switch]$true
        $script:Force = [switch]$true
        $script:WebhookSecretReference = "env:SYSTEMUPDATEPRO_WEBHOOK_URL"
        $script:WebhookUrl = "https://example.invalid/hook"
        $script:HistoryCount = 25
        $script:MaxRetries = 5
        $script:MaxUpdatePasses = 6
        $script:MinDiskSpaceGB = 20
        $script:MinFirmwareChargePercent = 65
        $script:LogRetentionDays = 60
        $script:EvidenceMaxSizeMB = 256
        $script:RedactionMode = "Secrets"
        [void]$script:StageResults.Add((New-StageResult -Name "WindowsUpdate" -Status "Succeeded" `
            -Attempted 1 -Installed 1 -RebootRequired $true))
        $state = New-ContinuationState -StageCursor "Winget" -ScriptPath $script:SourceScriptPath
        $state.Phase = "AwaitingReboot"
        (Save-State -State $state) | Should -BeTrue
        $loaded = Get-State

        Initialize-SystemUpdateProTestState
        $result = Import-ContinuationState -State $loaded

        $result.Success | Should -BeTrue
        $script:RunId | Should -Be "11111111-1111-1111-1111-111111111111"
        $script:ContinuationAttempt | Should -Be 1
        $script:ResumeStageCursor | Should -Be "Winget"
        $script:SkipOEM.IsPresent | Should -BeTrue
        $script:SkipWindows.IsPresent | Should -BeTrue
        $script:IncludeBIOS.IsPresent | Should -BeTrue
        $script:BypassWSUS.IsPresent | Should -BeTrue
        $script:RepairWindowsUpdate.IsPresent | Should -BeTrue
        $script:CleanupAfter.IsPresent | Should -BeTrue
        $script:ResetComponentBase.IsPresent | Should -BeTrue
        $script:BackupDrivers.IsPresent | Should -BeTrue
        $script:Force.IsPresent | Should -BeTrue
        $script:WebhookSecretReference | Should -Be "env:SYSTEMUPDATEPRO_WEBHOOK_URL"
        $script:WebhookUrl | Should -Be "https://example.invalid/hook"
        $script:HistoryCount | Should -Be 25
        $script:MaxRetries | Should -Be 5
        $script:MaxUpdatePasses | Should -Be 6
        $script:MinDiskSpaceGB | Should -Be 20
        $script:MinFirmwareChargePercent | Should -Be 65
        $script:LogRetentionDays | Should -Be 60
        $script:EvidenceMaxSizeMB | Should -Be 256
        $script:RedactionMode | Should -Be "Secrets"
        $script:StageResults[0].RebootRequired | Should -BeFalse
        $script:StageResults[0].RebootSatisfied | Should -BeTrue
        (Get-State).Phase | Should -Be "Running"
    }

    It "quarantines malformed or incompatible state instead of executing it" {
        Set-Content -LiteralPath $script:StateFile -Value "{not-json"

        $loaded = Get-State

        $loaded.Count | Should -Be 0
        Test-Path -LiteralPath $script:StateFile | Should -BeFalse
        @(Get-ChildItem -LiteralPath (Split-Path -Parent $script:StateFile) -Filter "state.corrupt.*.json").Count | Should -Be 1
    }

    It "rejects resume attempts beyond the persisted bound" {
        $state = New-ContinuationState -StageCursor "WindowsUpdate" -ScriptPath $script:SourceScriptPath
        $state.Phase = "AwaitingReboot"
        $state.AttemptCount = 3
        $state.MaxAttempts = 3

        $result = Import-ContinuationState -State $state

        $result.Success | Should -BeFalse
        $result.Message | Should -Match "limit"
        $script:ContinuationActive | Should -BeFalse
    }

    It "persists the next cursor before a resumed stage can advance" {
        $state = New-ContinuationState -StageCursor "WindowsUpdate" -ScriptPath $script:SourceScriptPath
        $state.Phase = "AwaitingReboot"
        (Import-ContinuationState -State $state).Success | Should -BeTrue
        [void]$script:StageResults.Add((New-StageResult -Name "WindowsUpdate" -Status "Succeeded"))

        (Set-ContinuationCursor -StageCursor "Winget") | Should -BeTrue
        (Get-State).StageCursor | Should -Be "Winget"
        $script:ResumeStageCursor = "Winget"
        (Test-ShouldRunContinuationStage -Stage "WindowsUpdate") | Should -BeFalse
        (Test-ShouldRunContinuationStage -Stage "Winget") | Should -BeTrue
        (Test-ShouldRunContinuationStage -Stage "Cleanup") | Should -BeTrue
    }

    It "rolls back state when scheduled-task registration fails" {
        Mock Get-ScheduledTask { $null }
        Mock Unregister-ScheduledTask {}
        Mock New-ScheduledTaskAction { $script:TaskActionStub }
        Mock New-ScheduledTaskTrigger { $script:TaskTriggerStub }
        Mock New-ScheduledTaskPrincipal { $script:TaskPrincipalStub }
        Mock New-ScheduledTaskSettingsSet { $script:TaskSettingsStub }
        Mock Register-ScheduledTask { throw "registration fault" }

        Register-ContinuationTask | Should -BeFalse

        Test-Path -LiteralPath $script:StateFile | Should -BeFalse
        $script:ContinuationRegistered | Should -BeFalse
        Should -Invoke Register-ScheduledTask -Times 1 -Exactly
        Should -Invoke Unregister-ScheduledTask -Times 0 -Exactly
        Test-Path -LiteralPath (Get-MutationJournalPath) | Should -BeFalse
    }

    It "commits awaiting-reboot state and keeps secrets out of task arguments" {
        $script:TaskPresent = $false
        Mock Get-ScheduledTask {
            if ($script:TaskPresent) { return "task" }
            return $null
        }
        Mock Unregister-ScheduledTask {}
        Mock New-ScheduledTaskAction { $script:TaskActionStub }
        Mock New-ScheduledTaskTrigger { $script:TaskTriggerStub }
        Mock New-ScheduledTaskPrincipal { $script:TaskPrincipalStub }
        Mock New-ScheduledTaskSettingsSet { $script:TaskSettingsStub }
        Mock Register-ScheduledTask { $script:TaskPresent = $true; "task" }
        $script:WebhookSecretReference = "env:SYSTEMUPDATEPRO_WEBHOOK_URL"
        $script:WebhookUrl = "https://example.invalid/private-hook"

        Register-ContinuationTask | Should -BeTrue
        $loaded = Get-State

        $loaded.Phase | Should -Be "AwaitingReboot"
        $loaded.Parameters.WebhookSecretReference |
            Should -Be "env:SYSTEMUPDATEPRO_WEBHOOK_URL"
        $loaded.ResolvedWebhookUrl | Should -Be "https://example.invalid/private-hook"
        $script:ContinuationRegistered | Should -BeTrue
        (Get-MutationJournal).Status | Should -Be "Open"
        Should -Invoke New-ScheduledTaskAction -Times 1 -Exactly -ParameterFilter {
            $Argument -notmatch "WebhookUrl|private-hook" -and
            $Argument -match "-NonInteractive" -and
            $Argument -match "-File"
        }
    }

    It "removes the one-shot task and state on terminal cleanup" {
        $state = New-ContinuationState -ScriptPath $script:SourceScriptPath
        $state.Phase = "Running"
        (Save-State -State $state) | Should -BeTrue
        Mock Get-ScheduledTask { "task" }
        Mock Unregister-ScheduledTask {}

        Unregister-ContinuationTask | Should -BeTrue

        Test-Path -LiteralPath $script:StateFile | Should -BeFalse
        Should -Invoke Unregister-ScheduledTask -Times 1 -Exactly
    }
}

Describe "Privileged mutation journal and crash recovery" {
    BeforeEach {
        Initialize-SystemUpdateProTestState
        Mock Write-Log {}
    }

    It "atomically persists a protected before-image before mutation" {
        $before = [ordered]@{
            Path = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"
            Name = "UseWUServer"; KeyExists = $true; Exists = $true
            Value = 1; Kind = "DWord"
        }

        $entryId = Add-MutationJournalEntry -Type "RegistryValue" `
            -Target "$($before.Path)\$($before.Name)" -Before $before `
            -RecoveryAction "Restore exact policy value" -Scope "WSUS"
        $journal = Get-MutationJournal
        $journalPath = Get-MutationJournalPath

        $journal.Entries.Count | Should -Be 1
        $journal.Entries[0].Id | Should -Be $entryId
        $journal.Entries[0].State | Should -Be "Planned"
        (Test-ContinuationStateAccess -Path $journalPath).Valid | Should -BeTrue
        (Get-Acl -LiteralPath $journalPath).AreAccessRulesProtected | Should -BeTrue
        @(Get-ChildItem -LiteralPath $script:MutationJournalDirectory -Filter "*.tmp.*").Count |
            Should -Be 0
        $script:MutationEvidence[0].Action | Should -Be "Journaled"
    }

    It "restores entries in reverse mutation order" {
        $script:MutationJournal = New-MutationJournal
        $script:MutationJournal.Entries = @(
            [ordered]@{
                Id = [guid]::NewGuid().ToString(); Sequence = 1; Type = "RegistryValue"
                Target = "first"; Scope = "Ordered"; State = "Applied"; RestoreOnFinalize = $true
                RecoveryAction = "first"; Before = [ordered]@{}; Metadata = [ordered]@{}
            },
            [ordered]@{
                Id = [guid]::NewGuid().ToString(); Sequence = 2; Type = "RegistryValue"
                Target = "second"; Scope = "Ordered"; State = "Applied"; RestoreOnFinalize = $true
                RecoveryAction = "second"; Before = [ordered]@{}; Metadata = [ordered]@{}
            }
        )
        $script:RestoreOrder = [System.Collections.ArrayList]::new()
        Mock Restore-MutationJournalEntry {
            [void]$script:RestoreOrder.Add([int]$Entry.Sequence)
            return $true
        }

        Restore-MutationJournalScope -Scope "Ordered" | Should -BeTrue

        @($script:RestoreOrder) | Should -Be @(2, 1)
    }

    It "restores an interrupted Windows Update cache swap on the next run" {
        $cachePath = Join-Path $script:WindowsRoot "SoftwareDistribution"
        New-Item -ItemType Directory -Path $cachePath -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $cachePath "old-cache.txt") -Value "before"
        $staleRunId = $script:RunId

        [void](Invoke-JournaledDirectoryReset -Path $cachePath -Services @())
        Set-Content -LiteralPath (Join-Path $cachePath "new-cache.txt") -Value "after"
        $backupPath = "$cachePath.SystemUpdatePro.$staleRunId.bak"
        Test-Path -LiteralPath $backupPath | Should -BeTrue

        $script:RunId = [guid]::NewGuid().ToString()
        $script:MutationJournal = $null
        $recovery = Invoke-UnfinishedMutationRecovery

        $recovery.Recovered | Should -Be 1
        $recovery.Failed | Should -Be 0
        Test-Path -LiteralPath (Join-Path $cachePath "old-cache.txt") | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $cachePath "new-cache.txt") | Should -BeFalse
        Test-Path -LiteralPath $backupPath | Should -BeFalse
        Test-Path -LiteralPath (Join-Path $script:MutationJournalDirectory "$staleRunId.json") |
            Should -BeFalse
    }

    It "commits a verified cache replacement only on normal finalization" {
        $cachePath = Join-Path $script:WindowsRoot "System32\catroot2"
        New-Item -ItemType Directory -Path $cachePath -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $cachePath "old.cat") -Value "before"
        $backupPath = "$cachePath.SystemUpdatePro.$($script:RunId).bak"

        [void](Invoke-JournaledDirectoryReset -Path $cachePath -Services @())
        Set-Content -LiteralPath (Join-Path $cachePath "new.cat") -Value "after"
        Complete-MutationJournal | Should -BeTrue

        Test-Path -LiteralPath (Join-Path $cachePath "new.cat") | Should -BeTrue
        Test-Path -LiteralPath $backupPath | Should -BeFalse
        Test-Path -LiteralPath (Get-MutationJournalPath) | Should -BeFalse
    }

    It "completes a durably decided cache commit after interruption" {
        $cachePath = Join-Path $script:WindowsRoot "SoftwareDistribution"
        New-Item -ItemType Directory -Path $cachePath -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $cachePath "old-cache.txt") -Value "before"
        $staleRunId = $script:RunId
        $backupPath = "$cachePath.SystemUpdatePro.$staleRunId.bak"

        $entryId = Invoke-JournaledDirectoryReset -Path $cachePath -Services @()
        Set-Content -LiteralPath (Join-Path $cachePath "new-cache.txt") -Value "after"
        [void](Set-MutationJournalEntryState -EntryId $entryId -State "Committing")

        $script:RunId = [guid]::NewGuid().ToString()
        $script:MutationJournal = $null
        $recovery = Invoke-UnfinishedMutationRecovery

        $recovery.Recovered | Should -Be 1
        Test-Path -LiteralPath (Join-Path $cachePath "new-cache.txt") | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $cachePath "old-cache.txt") | Should -BeFalse
        Test-Path -LiteralPath $backupPath | Should -BeFalse
        Test-Path -LiteralPath (Join-Path $script:MutationJournalDirectory "$staleRunId.json") |
            Should -BeFalse
    }

    It "refuses an unfinished journal whose ACL cannot be trusted" {
        $journal = New-MutationJournal
        $journalPath = Get-MutationJournalPath
        $journal | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $journalPath
        Mock Test-EvidencePathAccess {
            [PSCustomObject]@{ Valid = $false; Reason = "untrusted ACL" }
        }

        $result = Invoke-UnfinishedMutationRecovery

        $result.Failed | Should -Be 1
        $result.Messages[0] | Should -Match "untrusted ACL"
        Test-Path -LiteralPath $journalPath | Should -BeFalse
        @(Get-ChildItem -LiteralPath $script:MutationJournalDirectory -Filter "*.corrupt.*").Count |
            Should -Be 1
    }

    It "preserves zero, missing, binary, and multi-string registry before-images exactly" {
        $zero = [ordered]@{
            Path = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"
            Name = "UseWUServer"; KeyExists = $true; Exists = $true
            Value = 0; Kind = "DWord"
        }
        $missing = [ordered]@{
            Path = $zero.Path; Name = $zero.Name; KeyExists = $true
            Exists = $false; Value = $null; Kind = ""
        }
        $binary = ConvertTo-RegistryValueKind -Value @(0, 127, 255) -Kind "Binary"
        $multi = ConvertTo-RegistryValueKind -Value @("first", "second") -Kind "MultiString"

        (Test-RegistrySnapshotEqual -Expected $zero -Actual $zero) | Should -BeTrue
        (Test-RegistrySnapshotEqual -Expected $missing -Actual $missing) | Should -BeTrue
        $binary.GetType().FullName | Should -Be "System.Byte[]"
        @($binary) | Should -Be @(0, 127, 255)
        $multi.GetType().FullName | Should -Be "System.String[]"
        @($multi) | Should -Be @("first", "second")
    }

    It "restores the exact preexisting scheduled-task XML" {
        $xml = "<Task><RegistrationInfo /><Triggers /><Principals /><Settings /><Actions /></Task>"
        $script:TaskExists = $true
        Mock Get-ScheduledTask {
            if ($script:TaskExists) { return "task" }
            return $null
        }
        Mock Unregister-ScheduledTask { $script:TaskExists = $false }
        Mock Register-ScheduledTask { $script:TaskExists = $true; "restored-task" }
        Mock Export-ScheduledTask { $xml }
        $snapshot = [ordered]@{
            Exists = $true; TaskName = $script:TaskName; TaskPath = "\"; Xml = $xml
        }

        Restore-ScheduledTaskSnapshot -Snapshot $snapshot | Should -BeTrue

        Should -Invoke Unregister-ScheduledTask -Times 1 -Exactly
        Should -Invoke Register-ScheduledTask -Times 1 -Exactly -ParameterFilter {
            $Xml -eq $xml -and $TaskPath -eq "\"
        }
    }

    It "restores service status, startup mode, and delayed-start existence" {
        $snapshot = [ordered]@{
            Exists = $true; Name = "wuauserv"; Status = "Running"
            StartupType = "Automatic"; DelayedAutoStartExists = $false
            DelayedAutoStartValue = 0; RestartOnRestore = $false
        }
        Mock Set-ServiceStartupType {}
        Mock Remove-ItemProperty {}
        Mock Get-Service { [PSCustomObject]@{ Status = "Running"; StartType = "Automatic" } }
        Mock Get-ServiceStartupType { "Automatic" }
        Mock Get-ServiceDelayedAutoStartSnapshot {
            [ordered]@{ Exists = $false; Value = 0 }
        }

        Restore-ServiceSnapshot -Snapshot $snapshot | Should -BeTrue

        Should -Invoke Set-ServiceStartupType -Times 1 -Exactly -ParameterFilter {
            $ServiceName -eq "wuauserv" -and $StartupType -eq "Automatic"
        }
        Should -Invoke Remove-ItemProperty -Times 1 -Exactly -ParameterFilter {
            $Name -eq "DelayedAutoStart"
        }
    }

    It "does not mutate a healthy Windows Update runtime" {
        Mock Test-WindowsUpdateRuntime {
            [PSCustomObject]@{ Healthy = $true; Message = "healthy" }
        }
        Mock Set-JournaledServiceState {}
        Mock Invoke-JournaledDirectoryReset {}

        Repair-WindowsUpdateServices | Should -BeTrue

        Should -Invoke Set-JournaledServiceState -Times 0 -Exactly
        Should -Invoke Invoke-JournaledDirectoryReset -Times 0 -Exactly
    }

    It "removes irreversible network resets and cache wildcard deletion from repair" {
        $source = Get-Content -LiteralPath $script:SourceScriptPath -Raw

        $source | Should -Not -Match "regsvr32\.exe"
        $source | Should -Not -Match "netsh\s+winsock\s+reset"
        $source | Should -Not -Match "netsh\s+winhttp\s+reset\s+proxy"
        $source | Should -Not -Match "SoftwareDistribution\\Download\\\*"
        $source | Should -Not -Match "catroot2\\\*"
    }

    It "journals WSUS value absence and service state before enabling bypass" {
        $auPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"
        $script:RegistryWritten = $false
        $script:JournalEntryNumber = 0
        Mock Test-Path { $true } -ParameterFilter { $LiteralPath -eq $auPath }
        Mock Get-ServiceSnapshot {
            [ordered]@{
                Exists = $true; Name = "wuauserv"; Status = "Running"
                StartupType = "Manual"; DelayedAutoStartExists = $false
                DelayedAutoStartValue = 0; RestartOnRestore = $false
            }
        }
        Mock Get-RegistryValueSnapshot {
            [ordered]@{
                Path = $Path; Name = $Name; KeyExists = $true
                Exists = $script:RegistryWritten
                Value = $(if ($script:RegistryWritten) { 0 } else { $null })
                Kind = $(if ($script:RegistryWritten) { "DWord" } else { "" })
            }
        }
        Mock Add-MutationJournalEntry {
            $script:JournalEntryNumber++
            return "entry-$($script:JournalEntryNumber)"
        }
        Mock New-ItemProperty { $script:RegistryWritten = $true }
        Mock Set-MutationJournalEntryState { $true }
        Mock Restart-Service {}
        Mock Get-Service { [PSCustomObject]@{ Status = "Running" } }

        Set-WSUSBypass -Enable | Should -BeTrue

        Should -Invoke Add-MutationJournalEntry -Times 2 -Exactly
        Should -Invoke Add-MutationJournalEntry -Times 1 -Exactly -ParameterFilter {
            $Type -eq "RegistryValue" -and -not $Before.Exists -and $Scope -eq "WSUS"
        }
        Should -Invoke Set-MutationJournalEntryState -Times 2 -Exactly -ParameterFilter {
            $State -eq "Applied"
        }
        Should -Invoke Restart-Service -Times 1 -Exactly -ParameterFilter {
            $Name -eq "wuauserv"
        }
    }

    It "keeps continuation-task recovery durable across a reboot boundary" {
        $snapshot = [ordered]@{
            Exists = $false; TaskName = $script:TaskName; TaskPath = "\"; Xml = ""
        }
        [void](Add-MutationJournalEntry -Type "ScheduledTask" -Target "\$($script:TaskName)" `
            -Before $snapshot -RecoveryAction "Restore original task absence" -Scope "Continuation")
        Complete-MutationJournal -PreserveScopes @("Continuation") | Should -BeTrue

        (Get-MutationJournal).Status | Should -Be "AwaitingContinuation"
        Test-Path -LiteralPath (Get-MutationJournalPath) | Should -BeTrue

        Mock Restore-ScheduledTaskSnapshot { $true }
        Complete-MutationJournal | Should -BeTrue
        Test-Path -LiteralPath (Get-MutationJournalPath) | Should -BeFalse
    }
}

Describe "Fail-closed firmware safety" {
    BeforeEach {
        Initialize-SystemUpdateProTestState
        $script:Force = [switch]$false
        Mock Write-Log {}
        $script:FirmwareSystemInfo = @{
            Manufacturer = "Dell Inc."
            Model = "Latitude 7450"
        }
        $script:ReadyFirmwarePrerequisites = @{
            Disk = @{ Status = "Ready"; Sufficient = $true; FreeGB = 40; RequiredGB = 10; Message = "40 GB free" }
            Power = @{
                Status = "Ready"; HasBattery = $true; OnBattery = $false; OnACPower = $true
                ChargePercent = 80; RequiredChargePercent = 50; Message = "AC online and battery 80%"
            }
            BitLocker = @{
                Status = "Ready"; State = "Protected"; Enabled = $true; ProtectionOn = $true
                Suspended = $false; CanSuspend = $true; Method = "XtsAes256"; Message = "BitLocker active"
            }
        }
    }

    It "reports an unqueryable system drive as Unknown instead of sufficient" {
        Mock Get-CimInstance { throw "CIM unavailable" }

        $disk = Test-DiskSpace -MinGB 10

        $disk.Status | Should -Be "Unknown"
        $disk.Sufficient | Should -BeFalse
        $disk.FreeGB | Should -BeNullOrEmpty
        $disk.Message | Should -Match "Repair CIM/WMI"
    }

    It "requires both AC power and the configured minimum battery charge" {
        Mock Get-SystemPowerStatus {
            [PSCustomObject]@{
                PowerLineStatus = "Online"
                BatteryStatus = "High, Charging"
                BatteryLifePercent = 0.49
            }
        }

        $power = Test-BatteryPower -MinimumChargePercent 50

        $power.Status | Should -Be "Blocked"
        $power.OnACPower | Should -BeTrue
        $power.ChargePercent | Should -Be 49
        $power.Message | Should -Match "at least 50%"
    }

    It "returns Unknown when Windows power telemetry fails" {
        Mock Get-SystemPowerStatus { throw "GetSystemPowerStatus unavailable" }

        $power = Test-BatteryPower -MinimumChargePercent 50

        $power.Status | Should -Be "Unknown"
        $power.OnACPower | Should -BeFalse
        $power.Message | Should -Match "repair the power-status provider"
    }

    It "returns Unknown instead of disabled when BitLocker cannot be queried" {
        Mock Get-Command { $null } -ParameterFilter { $Name -eq "Get-BitLockerVolume" }

        $bitLocker = Test-BitLockerEnabled

        $bitLocker.Status | Should -Be "Unknown"
        $bitLocker.Enabled | Should -BeNullOrEmpty
        $bitLocker.Message | Should -Match "Get-BitLockerVolume"
    }

    It "keeps an unknown power state distinct and unsafe even with Force" {
        $script:Force = [switch]$true
        $unknownPrerequisites = @{
            Disk = $script:ReadyFirmwarePrerequisites.Disk
            Power = @{
                Status = "Unknown"; OnACPower = $false; ChargePercent = $null
                Message = "AC power state is unknown. Connect AC and repair the provider"
            }
            BitLocker = $script:ReadyFirmwarePrerequisites.BitLocker
        }

        $readiness = Test-FirmwareReadiness -Requested -Provider "Dell" `
            -SystemInfo $script:FirmwareSystemInfo -ToolState "Ready" `
            -SupportsBitLockerAutoSuspend $true -Prerequisites $unknownPrerequisites
        $item = New-FirmwareReadinessItem -Readiness $readiness

        $readiness.Status | Should -Be "Unknown"
        $readiness.Safe | Should -BeFalse
        $readiness.Message | Should -Match "-Force cannot override"
        $item.Status | Should -Be "Unknown"
    }

    It "allows protected BitLocker only for a provider with verified auto-suspension" {
        $dell = Test-FirmwareReadiness -Requested -Provider "Dell" `
            -SystemInfo $script:FirmwareSystemInfo -ToolState "Ready" `
            -SupportsBitLockerAutoSuspend $true -Prerequisites $script:ReadyFirmwarePrerequisites
        $hpInfo = @{ Manufacturer = "HP"; Model = "EliteBook 840 G11" }
        $hp = Test-FirmwareReadiness -Requested -Provider "HP" `
            -SystemInfo $hpInfo -ToolState "Ready" -Prerequisites $script:ReadyFirmwarePrerequisites

        $dell.Status | Should -Be "Ready"
        $dell.RequiresBitLockerSuspension | Should -BeTrue
        $hp.Status | Should -Be "Blocked"
        $hp.Message | Should -Match "Suspend protection"
    }

    It "uses one fail-closed update policy for Dell, Lenovo, and HP" {
        $blockedReadiness = [PSCustomObject]@{
            Requested = $true
            Safe = $false
            Status = "Blocked"
            Provider = "Dell"
            RequiresBitLockerSuspension = $false
        }

        $policy = Get-FirmwareUpdatePolicy -Readiness $blockedReadiness

        $policy.IncludeFirmware | Should -BeFalse
        @($policy.DellUpdateTypes) | Should -Not -Contain "bios"
        @($policy.DellUpdateTypes) | Should -Not -Contain "firmware"
        @($policy.HPCategories) | Should -Be @("Drivers")
        $policy.LenovoExcludeFirmware | Should -BeTrue
    }

    It "adds firmware only after every readiness check succeeds" {
        $readiness = Test-FirmwareReadiness -Requested -Provider "Dell" `
            -SystemInfo $script:FirmwareSystemInfo -ToolState "Ready" `
            -SupportsBitLockerAutoSuspend $true -Prerequisites $script:ReadyFirmwarePrerequisites

        $policy = Get-FirmwareUpdatePolicy -Readiness $readiness

        $policy.IncludeFirmware | Should -BeTrue
        @($policy.DellUpdateTypes) | Should -Contain "bios"
        @($policy.DellUpdateTypes) | Should -Contain "firmware"
        $policy.DellAutoSuspendBitLocker | Should -BeTrue
        @($policy.HPCategories) | Should -Contain "BIOS"
        $policy.LenovoExcludeFirmware | Should -BeFalse
    }

    It "detects Lenovo BIOS, UEFI, and firmware items consistently" {
        (Test-OEMUpdateIsFirmware -Update @{ Title = "ThinkPad BIOS Update"; Category = "Recommended"; Type = "Package" }) |
            Should -BeTrue
        (Test-OEMUpdateIsFirmware -Update @{ Title = "Dock"; Category = "Firmware"; Type = "Package" }) |
            Should -BeTrue
        (Test-OEMUpdateIsFirmware -Update @{ Title = "Intel display driver"; Category = "Driver"; Type = "Driver" }) |
            Should -BeFalse
    }

    It "treats missing model inventory or an unverified tool as Unknown" {
        $missingModel = Test-FirmwareReadiness -Requested -Provider "Dell" `
            -SystemInfo @{ Manufacturer = "Dell"; Model = "" } -ToolState "Ready" `
            -SupportsBitLockerAutoSuspend $true -Prerequisites $script:ReadyFirmwarePrerequisites
        $unknownTool = Test-FirmwareReadiness -Requested -Provider "Dell" `
            -SystemInfo $script:FirmwareSystemInfo -ToolState "Unknown" `
            -SupportsBitLockerAutoSuspend $true -Prerequisites $script:ReadyFirmwarePrerequisites

        $missingModel.Status | Should -Be "Unknown"
        $missingModel.Message | Should -Match "Win32_ComputerSystem"
        $unknownTool.Status | Should -Be "Unknown"
        $unknownTool.Message | Should -Match "OEM tool"
    }

    It "removes Dell BIOS and firmware types when a prerequisite is unknown" {
        $unknownPrerequisites = @{
            Disk = $script:ReadyFirmwarePrerequisites.Disk
            Power = @{ Status = "Unknown"; Message = "Power provider unavailable" }
            BitLocker = $script:ReadyFirmwarePrerequisites.BitLocker
        }
        Mock Get-DCUPath { "C:\Program Files\Dell\CommandUpdate\dcu-cli.exe" }
        Mock Test-DellInventoryCollector { $true }
        Mock Repair-DellServices { $true }
        Mock Get-Service { @() }
        Mock Restore-MutationJournalScope { $true }
        Mock Start-Process { [PSCustomObject]@{ ExitCode = 0 } }

        $result = Invoke-DellUpdate -IncludeBIOS -SystemInfo $script:FirmwareSystemInfo `
            -FirmwarePrerequisites $unknownPrerequisites

        $result.Status | Should -Be "Partial" -Because $result.Message
        $result.FirmwareReadiness.Status | Should -Be "Unknown"
        Should -Invoke Start-Process -Times 1 -Exactly -ParameterFilter {
            $ArgumentList -match "/applyUpdates" -and
            $ArgumentList -match "-updateType=driver,application,others" -and
            $ArgumentList -notmatch "-autoSuspendBitLocker"
        }
        Should -Invoke Restore-MutationJournalScope -Times 1 -Exactly -ParameterFilter {
            $Scope -eq "Dell"
        }
    }

    It "filters Lenovo firmware items while preserving safe driver discovery" {
        $script:DryRun = [switch]$true
        $unknownPrerequisites = @{
            Disk = $script:ReadyFirmwarePrerequisites.Disk
            Power = @{ Status = "Unknown"; Message = "Power provider unavailable" }
            BitLocker = $script:ReadyFirmwarePrerequisites.BitLocker
        }
        Mock Install-PSModuleWithRetry { $true }
        Mock Get-VerifiedPSModulePath { "C:\Program Files\WindowsPowerShell\Modules\LSUClient\1.8.1\LSUClient.psd1" }
        Mock Import-Module {}
        Mock Remove-Item {}
        Mock Get-LSUpdate {
            @(
                [PSCustomObject]@{
                    ID = "bios-1"; Title = "ThinkPad UEFI BIOS"; Category = "BIOS"; Type = "BIOS"
                    Installer = [PSCustomObject]@{ Unattended = $true }
                },
                [PSCustomObject]@{
                    ID = "driver-1"; Title = "Intel Ethernet Driver"; Category = "Driver"; Type = "Driver"
                    Installer = [PSCustomObject]@{ Unattended = $true }
                }
            )
        }

        $result = Invoke-LenovoUpdate -IncludeBIOS `
            -SystemInfo @{ Manufacturer = "Lenovo"; Model = "ThinkPad T14 Gen 5"; SerialNumber = "TEST" } `
            -FirmwarePrerequisites $unknownPrerequisites

        $result.Status | Should -Be "Partial" -Because $result.Message
        $result.Available | Should -Be 1
        @($result.Items | Where-Object Name -eq "ThinkPad UEFI BIOS").Status | Should -Be "Unknown"
        @($result.Items | Where-Object Name -eq "Intel Ethernet Driver").Status | Should -Be "Available"
    }

    It "limits HP application categories to drivers when BitLocker cannot be auto-suspended" {
        $hpInfo = @{ Manufacturer = "HP"; Model = "EliteBook 840 G11"; SerialNumber = "TEST" }
        Mock Get-HPIAPath { "C:\Program Files\HP\HPIA\HPImageAssistant.exe" }
        Mock Get-Process { @() }
        Mock New-Item {}
        Mock Remove-Item {}
        Mock New-ProtectedDirectory { $true }
        Mock Protect-EvidenceTree { $true }
        Mock Start-Process { [PSCustomObject]@{ ExitCode = 0 } }

        $result = Invoke-HPUpdate -IncludeBIOS -SystemInfo $hpInfo `
            -FirmwarePrerequisites $script:ReadyFirmwarePrerequisites

        $result.Status | Should -Be "Partial"
        $result.FirmwareReadiness.Status | Should -Be "Blocked"
        Should -Invoke Start-Process -Times 1 -Exactly -ParameterFilter {
            $ArgumentList -match "/Action:Install" -and
            $ArgumentList -match "/Category:Drivers" -and
            $ArgumentList -notmatch "/Category:[^ ]*(Firmware|BIOS)"
        }
    }
}

Describe "Explicit component cleanup safety" {
    BeforeEach {
        Initialize-SystemUpdateProTestState
        Mock Write-Log {}
    }

    It "retains update uninstallability during normal cleanup" {
        $script:DryRun = [switch]$true
        $script:CleanupAfter = [switch]$true

        $result = Invoke-ComponentCleanup

        $result.Success | Should -BeTrue
        $result.Status | Should -Be "Succeeded"
        $result.ResetBase | Should -BeFalse
        $result.RollbackImpact | Should -Match "retained"
        $result.Evidence[0] | Should -Not -Match "/ResetBase"
        $result.Items.Count | Should -Be 2
    }

    It "uses ResetBase only after the explicit high-risk switch" {
        $script:DryRun = [switch]$true
        $script:ResetComponentBase = [switch]$true

        $result = Invoke-ComponentCleanup

        $result.Success | Should -BeTrue
        $result.ResetBase | Should -BeTrue
        $result.RollbackImpact | Should -Match "IRREVERSIBLE"
        $result.Evidence[0] | Should -Match "/ResetBase"
        Should -Invoke Write-Log -ParameterFilter { $Level -eq "WARNING" -and $Message -match "IRREVERSIBLE" }
    }

    It "restores every temporary cleanmgr registry flag" {
        Mock Start-Process { [PSCustomObject]@{ ExitCode = 0 } }
        Mock Test-Path { $true }
        $script:WrittenRegistryPaths = @{}
        Mock Get-RegistryValueSnapshot {
            $written = $script:WrittenRegistryPaths.ContainsKey([string]$Path)
            [ordered]@{
                Path = $Path; Name = $Name; KeyExists = $true; Exists = $true
                Value = $(if ($written) { 2 } else { 7 }); Kind = "DWord"
            }
        }
        Mock New-ItemProperty { $script:WrittenRegistryPaths[[string]$LiteralPath] = $true }
        Mock Add-MutationJournalEntry { [guid]::NewGuid().ToString() }
        Mock Set-MutationJournalEntryState { $true }
        Mock Restore-MutationJournalScope { $true }

        $result = Invoke-ComponentCleanup

        $result.Success | Should -BeTrue
        $result.Status | Should -Be "Succeeded"
        Should -Invoke Add-MutationJournalEntry -Times 5 -Exactly -ParameterFilter {
            $Type -eq "RegistryValue" -and $Scope -eq "Cleanup"
        }
        Should -Invoke Set-MutationJournalEntryState -Times 5 -Exactly -ParameterFilter {
            $State -eq "Applied"
        }
        Should -Invoke Restore-MutationJournalScope -Times 1 -Exactly -ParameterFilter {
            $Scope -eq "Cleanup"
        }
        Should -Invoke New-ItemProperty -Times 5 -Exactly -ParameterFilter {
            $Name -eq "StateFlags0100" -and $Value -eq 2
        }
    }

    It "reports Disk Cleanup or restoration faults as partial success" {
        Mock Start-Process {
            if ($FilePath -eq "dism.exe") { return [PSCustomObject]@{ ExitCode = 0 } }
            return [PSCustomObject]@{ ExitCode = 5 }
        }
        Mock Test-Path { $true }
        $script:WrittenRegistryPaths = @{}
        Mock Get-RegistryValueSnapshot {
            $written = $script:WrittenRegistryPaths.ContainsKey([string]$Path)
            [ordered]@{
                Path = $Path; Name = $Name; KeyExists = $true; Exists = $written
                Value = $(if ($written) { 2 } else { $null })
                Kind = $(if ($written) { "DWord" } else { "" })
            }
        }
        Mock New-ItemProperty { $script:WrittenRegistryPaths[[string]$LiteralPath] = $true }
        Mock Add-MutationJournalEntry { [guid]::NewGuid().ToString() }
        Mock Set-MutationJournalEntryState { $true }
        Mock Restore-MutationJournalScope { $false }

        $result = Invoke-ComponentCleanup
        $stage = ConvertTo-StageResult -Name "Cleanup" -Provider "DISM and cleanmgr" -Result $result

        $result.Success | Should -BeFalse
        $result.Status | Should -Be "Partial"
        $stage.Status | Should -Be "Partial"
        $stage.Failed | Should -BeGreaterThan 0
        $stage.Message | Should -Match "failures"
        Should -Invoke Add-MutationJournalEntry -Times 5 -Exactly
        Should -Invoke Restore-MutationJournalScope -Times 1 -Exactly -ParameterFilter {
            $Scope -eq "Cleanup"
        }
    }

    It "states irreversible rollback impact in the HTML report" {
        $script:ResetComponentBase = [switch]$true
        $script:LogFile = Join-Path $TestDrive "run.log"
        Mock Start-Process {}
        $run = @{
            ExitCode = 0; DurationSeconds = 5; OEMUpdates = 0; WindowsUpdates = 0; WingetUpdates = 0
            RebootRequired = $false; Errors = @(); Warnings = @(); TotalInstalled = 0
        }
        $system = @{
            Manufacturer = "Test"; Model = "Device"; SerialNumber = "123"
            OSName = "Windows"; OSBuild = "1"; BIOSVersion = "1"; BIOSDate = Get-Date
            Processor = "CPU"; TotalRAM = 16
        }

        $reportPath = New-HTMLReport -SysInfo $system -RunData $run
        $report = Get-Content -LiteralPath $reportPath -Raw

        $report | Should -Match "Component rollback"
        $report | Should -Match "Disabled by irreversible /ResetBase"
    }
}

Describe "Versioned retryable webhook delivery" {
    BeforeEach {
        Initialize-SystemUpdateProTestState
        Mock Write-Log {}
    }

    BeforeAll {
        function New-WebhookTestRun {
            param([string]$ReportPath = "")

            return @{
                SchemaVersion = 1
                RunId = "22222222-2222-2222-2222-222222222222"
                StartedAt = "2026-07-29T20:00:00.0000000Z"
                CompletedAt = "2026-07-29T20:05:00.0000000Z"
                Status = "Partial"
                OEMUpdates = 1
                WindowsUpdates = 2
                WingetUpdates = 3
                TotalInstalled = 6
                TotalAvailable = 7
                TotalFailed = 1
                RebootRequired = $true
                ExitCode = 2
                DurationSeconds = 300
                Errors = @("One update failed")
                Warnings = @()
                Stages = @(
                    [ordered]@{
                        Name = "WindowsUpdate"
                        Provider = "Windows Update"
                        Status = "Partial"
                        Attempted = 3
                        Available = 3
                        Installed = 2
                        Failed = 1
                        Skipped = 0
                        RebootRequired = $true
                        ProviderExitCode = 2
                    }
                )
                Dependencies = @()
                MutationRecovery = @()
                Capabilities = $null
                Retention = $null
                EvidenceDelivery = @{
                    Report = @{
                        Status = $(if ($ReportPath) { "Succeeded" } else { "Failed" })
                        Detail = $ReportPath
                    }
                }
            }
        }
    }

    It "builds a correlated v2 contract and Teams Workflow Adaptive Card" {
        $reportPath = Join-Path $TestDrive "SystemUpdatePro_Report.html"
        [IO.File]::WriteAllText($reportPath, "<html></html>")
        $run = New-WebhookTestRun -ReportPath $reportPath

        $payload = New-WebhookPayload -RunData $run
        $channel = Get-WebhookChannel -Url (
            "https://prod-00.westus.logic.azure.com:443/workflows/abc/" +
            "triggers/manual/paths/invoke?api-version=2016-10-01&sig=test"
        )
        $request = ConvertTo-WebhookRequest -Payload $payload -Channel $channel
        $card = $request.Body | ConvertFrom-Json

        $payload.schema_version | Should -Be 2
        $payload.run_id | Should -Be $run.RunId
        $payload.idempotency_key | Should -Match "^[a-f0-9]{64}$"
        $payload.started_at | Should -Be $run.StartedAt
        $payload.completed_at | Should -Be $run.CompletedAt
        $payload.evidence_uri | Should -Match "^file:/"
        $payload.stage_summary.Count | Should -Be 1
        $payload.stage_summary[0].failed | Should -Be 1
        $channel | Should -Be "TeamsWorkflow"
        $card.type | Should -Be "message"
        $card.attachments[0].contentType |
            Should -Be "application/vnd.microsoft.card.adaptive"
        $card.attachments[0].content.type | Should -Be "AdaptiveCard"
        $request.Body | Should -Match $run.RunId
        $request.Body | Should -Match $payload.idempotency_key
        $request.Body | Should -Match "file:"
    }

    It "honors Retry-After for 429 and persists every attempt with terminal success" {
        $script:WebhookResponses = [System.Collections.Queue]::new()
        $script:WebhookResponses.Enqueue([PSCustomObject]@{
            Success = $false
            StatusCode = 429
            RetryAfterSeconds = 7
            Error = "rate limited"
        })
        $script:WebhookResponses.Enqueue([PSCustomObject]@{
            Success = $true
            StatusCode = 202
            RetryAfterSeconds = $null
            Error = ""
        })
        Mock Invoke-WebhookRequest { $script:WebhookResponses.Dequeue() }
        Mock Start-Sleep {}
        $run = New-WebhookTestRun

        $result = Send-WebhookNotification -Url "https://example.test/webhook?sig=SECRET" `
            -RunData $run -MaximumAttempts 3

        $result.TerminalStatus | Should -Be "Succeeded"
        $result.AttemptCount | Should -Be 2
        $result.Attempts[0].StatusCode | Should -Be 429
        $result.Attempts[0].Outcome | Should -Be "Retrying"
        $result.Attempts[0].DelayBeforeNextSeconds | Should -Be 7
        $result.Attempts[1].StatusCode | Should -Be 202
        $result.LocalRecordStatus | Should -Be "Succeeded"
        Should -Invoke Invoke-WebhookRequest -Times 2 -Exactly -ParameterFilter {
            $Headers["Idempotency-Key"] -eq $result.IdempotencyKey -and
            $Headers["X-SystemUpdatePro-Run-Id"] -eq $run.RunId -and
            $Body -match '"schema_version":2'
        }
        Should -Invoke Start-Sleep -Times 1 -Exactly -ParameterFilter {
            $Seconds -eq 7
        }

        $record = Get-Content -LiteralPath $result.LocalRecordPath -Raw |
            ConvertFrom-Json
        $record.TerminalStatus | Should -Be "Succeeded"
        $record.AttemptCount | Should -Be 2
        $record.Attempts[0].StatusCode | Should -Be 429
        (Get-Content -LiteralPath $result.LocalRecordPath -Raw) |
            Should -Not -Match "SECRET"
        (Test-EvidencePathAccess -Path $result.LocalRecordPath).Valid |
            Should -BeTrue
        Test-Path -LiteralPath "$($result.LocalRecordPath).previous" |
            Should -BeTrue
        (Test-EvidencePathAccess -Path "$($result.LocalRecordPath).previous").Valid |
            Should -BeTrue
    }

    It "uses bounded exponential delays and stops after the configured attempts" {
        Mock Invoke-WebhookRequest {
            [PSCustomObject]@{
                Success = $false
                StatusCode = 503
                RetryAfterSeconds = $null
                Error = "temporarily unavailable"
            }
        }
        Mock Start-Sleep {}
        $run = New-WebhookTestRun

        $result = Send-WebhookNotification -Url "https://example.test/webhook" `
            -RunData $run -MaximumAttempts 3

        $result.TerminalStatus | Should -Be "Failed"
        $result.AttemptCount | Should -Be 3
        $result.Attempts[0].DelayBeforeNextSeconds | Should -Be 2
        $result.Attempts[1].DelayBeforeNextSeconds | Should -Be 4
        $result.Attempts[2].DelayBeforeNextSeconds | Should -Be 0
        Should -Invoke Invoke-WebhookRequest -Times 3 -Exactly
        Should -Invoke Start-Sleep -Times 2 -Exactly
    }

    It "does not retry a non-transient client rejection" {
        Mock Invoke-WebhookRequest {
            [PSCustomObject]@{
                Success = $false
                StatusCode = 400
                RetryAfterSeconds = $null
                Error = "bad request"
            }
        }
        Mock Start-Sleep {}

        $result = Send-WebhookNotification -Url "https://example.test/webhook" `
            -RunData (New-WebhookTestRun) -MaximumAttempts 5

        $result.TerminalStatus | Should -Be "Failed"
        $result.AttemptCount | Should -Be 1
        Should -Invoke Invoke-WebhookRequest -Times 1 -Exactly
        Should -Invoke Start-Sleep -Times 0 -Exactly
    }
}

Describe "Provider and terminal failure handling" {
    BeforeEach {
        Initialize-SystemUpdateProTestState
    }

    It "preserves a failed Windows Update pass with zero item failures" {
        Mock Write-Log {}
        Mock Install-PSModuleWithRetry { $true }
        Mock Invoke-WindowsUpdatePSWU {
            @{
                Success = $false; RebootRequired = $false
                Available = 0; Attempted = 0; Installed = 0; Failed = 0
                ExitCode = 55; HResult = -1; Items = @(); Evidence = @()
                Message = "provider failed before enumeration"
            }
        }

        $result = Invoke-WindowsUpdate -MaxPasses 3

        $result.Success | Should -BeFalse
        $result.Passes | Should -Be 1
        $result.ExitCode | Should -Be 55
        $result.Message | Should -Match "failed"
    }

    It "attempts each evidence sink once and finalizes idempotently" {
        $script:WebhookUrl = "https://example.invalid/hook"
        $reportPath = Join-Path $TestDrive "report.html"
        Set-Content -LiteralPath $reportPath -Value "<html></html>"

        Mock New-HTMLReport { $reportPath }
        Mock Initialize-EventLog { $true }
        Mock Write-EventLogEntry { $true }
        Mock Send-WebhookNotification {
            [PSCustomObject]@{
                SchemaVersion = 1
                PayloadSchemaVersion = 2
                Channel = "Generic"
                IdempotencyKey = "idempotency"
                EvidenceUri = "file:///report.html"
                MaximumAttempts = 3
                AttemptCount = 1
                Attempts = @(@{ Attempt = 1; StatusCode = 200; Outcome = "Succeeded" })
                TerminalStatus = "Succeeded"
                LocalRecordStatus = "Succeeded"
                LocalRecordPath = "delivery.json"
                Error = ""
            }
        }
        Mock Save-UpdateHistory { $true }

        $run = @{
            SchemaVersion = 1; RunId = "test-run-id"; Status = "Succeeded"
            TotalInstalled = 0; TotalAvailable = 0; TotalFailed = 0
            RebootRequired = $false; ExitCode = 0; DurationSeconds = 1
            EvidenceDelivery = @{}
        }
        $system = @{ Manufacturer = "Test"; Model = "Device" }

        $first = Invoke-TerminalEvidence -SysInfo $system -RunData $run -WebhookEndpoint $script:WebhookUrl
        $second = Invoke-TerminalEvidence -SysInfo $system -RunData $first -WebhookEndpoint $script:WebhookUrl

        foreach ($sink in @("Report", "Event", "Webhook", "History")) {
            $second.EvidenceDelivery[$sink].Status | Should -Be "Succeeded"
        }
        $second.EvidenceDelivery.Webhook.PayloadSchemaVersion | Should -Be 2
        $second.EvidenceDelivery.Webhook.TerminalStatus | Should -Be "Succeeded"
        $second.EvidenceDelivery.Webhook.AttemptCount | Should -Be 1
        Should -Invoke New-HTMLReport -Times 1 -Exactly
        Should -Invoke Write-EventLogEntry -Times 1 -Exactly
        Should -Invoke Send-WebhookNotification -Times 1 -Exactly
        Should -Invoke Save-UpdateHistory -Times 1 -Exactly
    }

    It "records independent sink failures without skipping later sinks" {
        $script:WebhookUrl = "https://example.invalid/hook"

        Mock New-HTMLReport { throw "report fault" }
        Mock Initialize-EventLog { $true }
        Mock Write-EventLogEntry { $false }
        Mock Send-WebhookNotification {
            [PSCustomObject]@{
                SchemaVersion = 1
                PayloadSchemaVersion = 2
                Channel = "Generic"
                IdempotencyKey = "idempotency"
                EvidenceUri = ""
                MaximumAttempts = 3
                AttemptCount = 1
                Attempts = @(@{ Attempt = 1; StatusCode = 500; Outcome = "Failed" })
                TerminalStatus = "Failed"
                LocalRecordStatus = "Succeeded"
                LocalRecordPath = "delivery.json"
                Error = "Webhook returned HTTP status 500"
            }
        }
        Mock Save-UpdateHistory { $true }

        $run = @{
            SchemaVersion = 1; RunId = "test-run-id"; Status = "Failed"
            TotalInstalled = 0; TotalAvailable = 0; TotalFailed = 1
            RebootRequired = $false; ExitCode = 3; DurationSeconds = 1
            EvidenceDelivery = @{}
        }
        $system = @{ Manufacturer = "Test"; Model = "Device" }

        $result = Invoke-TerminalEvidence -SysInfo $system -RunData $run -WebhookEndpoint $script:WebhookUrl

        $result.EvidenceDelivery.Report.Status | Should -Be "Failed"
        $result.EvidenceDelivery.Event.Status | Should -Be "Failed"
        $result.EvidenceDelivery.Webhook.Status | Should -Be "Failed"
        $result.EvidenceDelivery.Webhook.TerminalStatus | Should -Be "Failed"
        $result.EvidenceDelivery.Webhook.AttemptCount | Should -Be 1
        $result.EvidenceDelivery.History.Status | Should -Be "Succeeded"
        Should -Invoke Save-UpdateHistory -Times 1 -Exactly
    }
}
