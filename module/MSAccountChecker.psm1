<#
    MSAccountChecker module: the lookup logic of Microsoft Account Checker for scripts and CI.

    The rules are read from ..\script\MS_Account_Checker.ps1 when the module is imported, so the
    module and the desktop app always classify accounts the same way.

        Import-Module .\module\MSAccountChecker.psd1
        Test-MSAccount -Email user@contoso.com
        Import-Csv users.csv | Test-MSAccount | Export-Csv results.csv -NoTypeInformation
#>

$scriptPath = Join-Path $PSScriptRoot '..\script\MS_Account_Checker.ps1'
if (-not (Test-Path -LiteralPath $scriptPath)) { throw "MSAccountChecker: cannot find $scriptPath. Keep the module folder next to the script folder." }
$parseTokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile((Resolve-Path -LiteralPath $scriptPath).Path, [ref]$parseTokens, [ref]$parseErrors)
if ($parseErrors) { throw "MSAccountChecker: the script has syntax errors: $($parseErrors[0].Message)" }

function Get-ScriptBlockAssignment {
    param([string]$Name)
    $assignment = $ast.Find({
        param($node)
        $node -is [System.Management.Automation.Language.AssignmentStatementAst] -and $node.Left.Extent.Text -eq ('$script:' + $Name)
    }, $true)
    $block = $assignment.Right.Find({ param($node) $node -is [System.Management.Automation.Language.ScriptBlockExpressionAst] }, $true)
    $text = $block.Extent.Text.Trim()
    return $text.Substring(1, $text.Length - 2)
}
function Get-FunctionDefinition {
    param([string]$Name)
    $definition = $ast.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $Name }, $true)
    return $definition.Extent.Text
}

. ([scriptblock]::Create((Get-ScriptBlockAssignment 'CoreLogic')))
. ([scriptblock]::Create((Get-FunctionDefinition 'New-SharedState')))
. ([scriptblock]::Create((Get-FunctionDefinition 'Get-CategoryLabelText')))
. ([scriptblock]::Create((Get-FunctionDefinition 'Get-CliExitCode')))

$script:ModuleShared = New-SharedState
$CoreStats = $script:ModuleShared.Stats

function Test-MSAccount {
    <#
    .SYNOPSIS
        Identifies the Microsoft account type behind one or more email addresses.
    .DESCRIPTION
        Returns one object per address with the cloud, account type, status (OK, Inconclusive, Throttled,
        Unavailable, Invalid), recommended next step and tenant details. Pipe in strings, or objects with an
        Email, Mail, UserPrincipalName or EmailAddress property.
    .PARAMETER Email
        The addresses to check.
    .PARAMETER NoCache
        Ask Microsoft every time instead of reusing answers from earlier calls in this session.
    .PARAMETER DelayMs
        Pause between live lookups, to stay clear of Microsoft throttling. Default 200 ms.
    .EXAMPLE
        Test-MSAccount -Email user@contoso.com, other@fabrikam.com
    .EXAMPLE
        Import-Csv users.csv | Test-MSAccount | Where-Object status -ne 'OK'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('Mail', 'UserPrincipalName', 'EmailAddress', 'UPN')]
        [string[]]$Email,
        [switch]$NoCache,
        [ValidateRange(0, 10000)][int]$DelayMs = 200
    )
    process {
        foreach ($address in $Email) {
            if ([string]::IsNullOrWhiteSpace($address)) { continue }
            $cache = if ($NoCache) { $null } else { $script:ModuleShared.Cache }
            $result = Get-AccountCheckCached -Email $address -Cache $cache -DomainCache $script:ModuleShared.DomainCache -Policy $script:ModuleShared['CachePolicy']
            $output = [pscustomobject](ConvertTo-ApiResult $result)
            $output.PSObject.TypeNames.Insert(0, 'MSAccountChecker.Result')
            $output
            if ($result.Source -eq 'Live' -and $DelayMs -gt 0) { Start-Sleep -Milliseconds $DelayMs }
        }
    }
}

function Clear-MSAccountCache {
    <#
    .SYNOPSIS
        Forgets the answers cached by Test-MSAccount in this session.
    #>
    [CmdletBinding()]
    param()
    $script:ModuleShared.Cache.Clear()
    $script:ModuleShared.DomainCache.Clear()
}

function Get-MSAccountExitCode {
    <#
    .SYNOPSIS
        Turns Test-MSAccount results into a CI exit code: 0 all definitive, 1 some inconclusive, throttled
        or invalid, 2 some unavailable.
    .EXAMPLE
        $results = Test-MSAccount -Email $list
        exit (Get-MSAccountExitCode -Results $results)
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][object[]]$Results)
    return (Get-CliExitCode -Results $Results)
}

Export-ModuleMember -Function Test-MSAccount, Clear-MSAccountCache, Get-MSAccountExitCode
