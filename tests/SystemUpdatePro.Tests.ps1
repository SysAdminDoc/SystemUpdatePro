BeforeAll {
    $scriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) "SystemUpdatePro.ps1"
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

    function Initialize-SystemUpdateProTestState {
        $script:DryRun = [switch]$false
        $script:WebhookUrl = ""
        $script:ResultSchemaVersion = 1
        $script:RunId = "test-run-id"
        $script:StageResults = [System.Collections.ArrayList]::new()
        $script:Errors = [System.Collections.ArrayList]::new()
        $script:Warnings = [System.Collections.ArrayList]::new()
        $script:RebootRequired = $false
        $script:RunFinalized = $false
        $script:LastEvidenceDelivery = $null
        $script:WindowsUpdates = [System.Collections.ArrayList]::new()
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
