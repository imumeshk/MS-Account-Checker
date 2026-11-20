@{
    RootModule           = 'MSAccountChecker.psm1'
    ModuleVersion        = '2.0.0'
    GUID                 = '6f4c2d9e-3b7a-4c1e-9a52-8d0f3e7b6a41'
    Author               = 'imumeshk'
    Copyright            = 'Released under CC0-1.0'
    Description          = 'Identifies the Microsoft account type behind an email address across the Commercial, GCC High and China clouds, with tenant and federation details.'
    PowerShellVersion    = '5.1'
    CompatiblePSEditions = @('Desktop', 'Core')
    FunctionsToExport    = @('Test-MSAccount', 'Clear-MSAccountCache', 'Get-MSAccountExitCode')
    CmdletsToExport      = @()
    VariablesToExport    = @()
    AliasesToExport      = @()
    PrivateData          = @{
        PSData = @{
            Tags         = @('Microsoft365', 'Entra', 'AzureAD', 'Identity', 'GCCHigh', 'Audit')
            LicenseUri   = 'https://github.com/imumeshk/MS-Account-Checker/blob/main/LICENSE'
            ProjectUri   = 'https://github.com/imumeshk/MS-Account-Checker'
            ReleaseNotes = 'https://github.com/imumeshk/MS-Account-Checker/blob/main/CHANGELOG.md'
        }
    }
}
