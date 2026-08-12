$script:SystemUpdateProEntryPoint = Join-Path $PSScriptRoot 'SystemUpdatePro.ps1'

function ConvertTo-SystemUpdateProProcessArgument {
    param([AllowEmptyString()][string]$Value)

    if ($Value -notmatch '[\s"]') { return $Value }
    return '"' + ($Value -replace '(\\*)"', '$1$1\"' -replace '(\\+)$', '$1$1') + '"'
}

function Invoke-SystemUpdatePro {
    [CmdletBinding()]
    [OutputType([int])]
    param(
        [Parameter(Position = 0, ValueFromRemainingArguments = $true)]
        [AllowEmptyCollection()][string[]]$ArgumentList = @()
    )

    if (-not (Test-Path -LiteralPath $script:SystemUpdateProEntryPoint -PathType Leaf)) {
        throw "SystemUpdatePro entry point was not found: $script:SystemUpdateProEntryPoint"
    }
    $hostCommand = if ([string]$PSVersionTable.PSEdition -eq 'Core') { 'pwsh.exe' } else { 'powershell.exe' }
    $processArguments = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $script:SystemUpdateProEntryPoint)
    $processArguments += @($ArgumentList | ForEach-Object { ConvertTo-SystemUpdateProProcessArgument -Value ([string]$_) })
    $process = Start-Process -FilePath $hostCommand -ArgumentList $processArguments -Wait -NoNewWindow -PassThru -ErrorAction Stop
    return [int]$process.ExitCode
}

Export-ModuleMember -Function Invoke-SystemUpdatePro
