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
        "Get-AcquisitionManifest",
        "Get-AcquisitionManifestEntry",
        "Get-SystemArchitecture",
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
        "Send-WebhookNotification",
        "Save-UpdateHistory",
        "Invoke-TerminalEvidence",
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
        $script:ResultSchemaVersion = 1
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
        $script:StateSchemaVersion = 3
        $script:MaxContinuationAttempts = 3
        $script:RunStartedAt = Get-Date
        $stateTestDirectory = Join-Path $TestDrive ([guid]::NewGuid().ToString("N"))
        New-Item -ItemType Directory -Path $stateTestDirectory -Force | Out-Null
        $script:StateFile = Join-Path $stateTestDirectory "state.json"
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
        $script:AcquisitionManifestVersion = 1
        $script:AcquisitionManifest = Get-AcquisitionManifest
        $script:AcquisitionProvenance = [System.Collections.ArrayList]::new()
        $script:MutationJournalSchemaVersion = 1
        $script:MutationJournalDirectory = Join-Path $stateTestDirectory "Journals"
        New-Item -ItemType Directory -Path $script:MutationJournalDirectory -Force | Out-Null
        $script:MutationJournal = $null
        $script:MutationEvidence = [System.Collections.ArrayList]::new()
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
        $script:EntryScriptPath = $script:SourceScriptPath
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
        $run = New-RunData -StartedAt (Get-Date).AddSeconds(-1)
        $system = @{
            Manufacturer = "HP"; Model = "EliteBook"; SerialNumber = "123"
            OSName = "Windows"; OSBuild = "1"; BIOSVersion = "1"; BIOSDate = Get-Date
            Processor = "CPU"; TotalRAM = 16
        }

        $reportPath = New-HTMLReport -SysInfo $system -RunData $run
        $report = Get-Content -LiteralPath $reportPath -Raw

        $report | Should -Match "Dependency provenance"
        $report | Should -Match "HP Image Assistant"
        $report | Should -Match "5\.3\.6\.649"
        $report | Should -Match "C:\\ProgramData\\SystemUpdatePro\\HPIA"
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

        $loaded.SchemaVersion | Should -Be 3
        $loaded.Phase | Should -Be "AwaitingReboot"
        $loaded.RunId | Should -Be "11111111-1111-1111-1111-111111111111"
        $loaded.Parameters.Count | Should -Be 22
        $loaded.AcquisitionProvenance.Count | Should -Be 1
        $loaded.AcquisitionProvenance[0].Name | Should -Be "LSUClient"
        $loaded.Parameters.SkipOEM | Should -BeTrue
        $loaded.Parameters.IncludeBIOS | Should -BeTrue
        $loaded.Parameters.WebhookUrl | Should -Be "https://example.invalid/private-hook"
        $loaded.Parameters.MinFirmwareChargePercent | Should -Be 50
        $loaded.StageResults.Count | Should -Be 1
        (Test-ContinuationStateAccess -Path $script:StateFile).Valid | Should -BeTrue
        (Get-Acl -LiteralPath $script:StateFile).AreAccessRulesProtected | Should -BeTrue
        @(Get-ChildItem -LiteralPath (Split-Path -Parent $script:StateFile) -Filter "state.json.tmp.*").Count | Should -Be 0
        Test-Path -LiteralPath "$($script:StateFile).previous" | Should -BeFalse
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
        $script:WebhookUrl = "https://example.invalid/hook"
        $script:HistoryCount = 25
        $script:MaxRetries = 5
        $script:MaxUpdatePasses = 6
        $script:MinDiskSpaceGB = 20
        $script:MinFirmwareChargePercent = 65
        $script:LogRetentionDays = 60
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
        $script:WebhookUrl | Should -Be "https://example.invalid/hook"
        $script:HistoryCount | Should -Be 25
        $script:MaxRetries | Should -Be 5
        $script:MaxUpdatePasses | Should -Be 6
        $script:MinDiskSpaceGB | Should -Be 20
        $script:MinFirmwareChargePercent | Should -Be 65
        $script:LogRetentionDays | Should -Be 60
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
        $script:WebhookUrl = "https://example.invalid/private-hook"

        Register-ContinuationTask | Should -BeTrue
        $loaded = Get-State

        $loaded.Phase | Should -Be "AwaitingReboot"
        $loaded.Parameters.WebhookUrl | Should -Be "https://example.invalid/private-hook"
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
        Mock Test-ContinuationStateAccess {
            [PSCustomObject]@{ Valid = $false; Reason = "untrusted ACL" }
        }

        $result = Invoke-UnfinishedMutationRecovery

        $result.Failed | Should -Be 1
        $result.Messages[0] | Should -Match "untrusted ACL"
        Test-Path -LiteralPath $journalPath | Should -BeTrue
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
        Mock Send-WebhookNotification { $true }
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
        Mock Send-WebhookNotification { $false }
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
        $result.EvidenceDelivery.History.Status | Should -Be "Succeeded"
        Should -Invoke Save-UpdateHistory -Times 1 -Exactly
    }
}
