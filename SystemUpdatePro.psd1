@{
    RootModule = 'SystemUpdatePro.psm1'
    ModuleVersion = '4.2.0'
    GUID = '8bc11208-2e04-45dc-93a1-48d7a919ab10'
    Author = 'SysAdminDoc'
    CompanyName = 'SysAdminDoc'
    Copyright = '(c) SysAdminDoc. All rights reserved.'
    Description = 'Enterprise Multi-OEM Windows and application update orchestration.'
    PowerShellVersion = '5.1'
    CompatiblePSEditions = @('Desktop', 'Core')
    FunctionsToExport = @('Invoke-SystemUpdatePro')
    CmdletsToExport = @()
    VariablesToExport = @()
    AliasesToExport = @()
    FileList = @('SystemUpdatePro.ps1', 'SystemUpdatePro.psm1', 'SystemUpdatePro.psd1', 'README.md', 'LICENSE')
    PrivateData = @{
        PSData = @{
            Tags = @('Windows', 'updates', 'OEM', 'MSP', 'RMM', 'PowerShell')
            LicenseUri = 'https://github.com/SysAdminDoc/SystemUpdatePro/blob/main/LICENSE'
            ProjectUri = 'https://github.com/SysAdminDoc/SystemUpdatePro'
            ReleaseNotes = 'https://github.com/SysAdminDoc/SystemUpdatePro/blob/main/CHANGELOG.md'
        }
    }
}
