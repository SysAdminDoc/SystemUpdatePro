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
        "Get-DCUPath",
        "Install-DellCommandUpdate",
        "Repair-DellServices",
        "Invoke-DellUpdate",
        "Invoke-LenovoUpdate",
        "Get-HPIAPath",
        "Install-HPIA",
        "Invoke-HPUpdate",
        "Get-RegistryValueSnapshot",
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
        Should -Invoke Unregister-ScheduledTask -Times 2 -Exactly
    }

    It "commits awaiting-reboot state and keeps secrets out of task arguments" {
        Mock Unregister-ScheduledTask {}
        Mock New-ScheduledTaskAction { $script:TaskActionStub }
        Mock New-ScheduledTaskTrigger { $script:TaskTriggerStub }
        Mock New-ScheduledTaskPrincipal { $script:TaskPrincipalStub }
        Mock New-ScheduledTaskSettingsSet { $script:TaskSettingsStub }
        Mock Register-ScheduledTask { "task" }
        $script:WebhookUrl = "https://example.invalid/private-hook"

        Register-ContinuationTask | Should -BeTrue
        $loaded = Get-State

        $loaded.Phase | Should -Be "AwaitingReboot"
        $loaded.Parameters.WebhookUrl | Should -Be "https://example.invalid/private-hook"
        $script:ContinuationRegistered | Should -BeTrue
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
        Mock Repair-DellServices { $true }
        Mock Get-Service { @() }
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
    }

    It "filters Lenovo firmware items while preserving safe driver discovery" {
        $script:DryRun = [switch]$true
        $unknownPrerequisites = @{
            Disk = $script:ReadyFirmwarePrerequisites.Disk
            Power = @{ Status = "Unknown"; Message = "Power provider unavailable" }
            BitLocker = $script:ReadyFirmwarePrerequisites.BitLocker
        }
        Mock Install-PSModuleWithRetry { $true }
        Mock Get-Module { [PSCustomObject]@{ Name = "LSUClient"; Version = [Version]"1.8.1" } }
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
        Mock Get-RegistryValueSnapshot {
            [ordered]@{ Path = $Path; Name = $Name; Exists = $true; Value = 7; Kind = "DWord" }
        }
        Mock New-ItemProperty {}
        Mock Restore-RegistryValueSnapshot { $true }

        $result = Invoke-ComponentCleanup

        $result.Success | Should -BeTrue
        $result.Status | Should -Be "Succeeded"
        Should -Invoke Restore-RegistryValueSnapshot -Times 5 -Exactly
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
        Mock Get-RegistryValueSnapshot {
            [ordered]@{ Path = $Path; Name = $Name; Exists = $false; Value = $null; Kind = "" }
        }
        Mock New-ItemProperty {}
        Mock Restore-RegistryValueSnapshot { $false }

        $result = Invoke-ComponentCleanup
        $stage = ConvertTo-StageResult -Name "Cleanup" -Provider "DISM and cleanmgr" -Result $result

        $result.Success | Should -BeFalse
        $result.Status | Should -Be "Partial"
        $stage.Status | Should -Be "Partial"
        $stage.Failed | Should -BeGreaterThan 0
        $stage.Message | Should -Match "failures"
        Should -Invoke Restore-RegistryValueSnapshot -Times 5 -Exactly
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
