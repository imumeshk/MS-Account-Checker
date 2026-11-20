#requires -Version 5.1
<#
    Pester 5 tests for script/MS_Account_Checker.ps1.

    The app is never started. The lookup logic, the API server script and the helper functions are
    read out of the script with the PowerShell parser, so no window opens. Microsoft is never called:
    Invoke-RestMethod is replaced by a simulated Microsoft that answers from the $script:Fake table.

    Run from the repository folder:
        Invoke-Pester -Path .\tests -Output Detailed
#>

BeforeAll {
    $script:ScriptPath = (Resolve-Path (Join-Path $PSScriptRoot '..\script\MS_Account_Checker.ps1')).Path
    $parseTokens = $null
    $parseErrors = $null
    $script:Ast = [System.Management.Automation.Language.Parser]::ParseFile($script:ScriptPath, [ref]$parseTokens, [ref]$parseErrors)
    $script:ParseErrors = $parseErrors
    $script:SourceText  = Get-Content -LiteralPath $script:ScriptPath -Raw

    function Find-Assignment {
        param([string]$Name)
        $assignment = $script:Ast.Find({
            param($node)
            $node -is [System.Management.Automation.Language.AssignmentStatementAst] -and $node.Left.Extent.Text -eq ('$script:' + $Name)
        }, $true)
        if (-not $assignment) { throw "No assignment to `$script:$Name was found in the script." }
        return $assignment
    }
    function Get-AssignedScriptBlockText {
        param([string]$Name)
        $block = (Find-Assignment $Name).Right.Find({ param($node) $node -is [System.Management.Automation.Language.ScriptBlockExpressionAst] }, $true)
        $text = $block.Extent.Text.Trim()
        return $text.Substring(1, $text.Length - 2)
    }
    function Get-AssignedString {
        param([string]$Name)
        return [string](Find-Assignment $Name).Right.Expression.Value
    }
    function Get-FunctionText {
        param([string]$Name)
        $definition = $script:Ast.Find({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $Name
        }, $true)
        if (-not $definition) { throw "Function $Name was not found in the script." }
        return $definition.Extent.Text
    }

    # Load the shared lookup logic and the helpers under test into this scope.
    $script:CoreText = Get-AssignedScriptBlockText 'CoreLogic'
    . ([scriptblock]::Create($script:CoreText))
    foreach ($functionName in @('New-SharedState', 'New-Counts', 'Get-Bucket', 'Add-ToCounts', 'Get-BatchTokens', 'Get-ExportRows', 'New-HtmlReport', 'Test-HistoryItem',
                                'Get-CategoryLabelText', 'Test-IsFederated', 'Get-ResultChanges', 'Get-DomainSummary', 'ConvertTo-RedactedEmail', 'Import-EmailFile',
                                'Get-CliExitCode', 'Get-ThemePalette')) {
        . ([scriptblock]::Create((Get-FunctionText $functionName)))
    }
    $script:AppName    = 'Microsoft Account Checker'
    $script:AppVersion = 'test'

    # ---- Simulated Microsoft -------------------------------------------------------------------
    function Reset-Fake {
        $script:Fake = @{
            Commercial = 0; GCCHigh = 1; China = 1
            Throttle = 0; Fail429 = 0; Network = $false
            Realm = 'Managed'; Brand = 'Contoso'; AuthUrl = ''; FedUrl = ''; DomainType = 3
            TenantId = '11111111-2222-3333-4444-555555555555'
            Calls = New-Object System.Collections.Generic.List[string]
        }
    }
    function Invoke-FakeMicrosoft {
        param([string]$Uri)
        $script:Fake.Calls.Add($Uri)
        if ($Uri -like '*GetCredentialType*') {
            $cloud = if ($Uri -like '*microsoftonline.us*') { 'GCCHigh' } elseif ($Uri -like '*microsoftonline.cn*') { 'China' } else { 'Commercial' }
            if ($cloud -eq 'Commercial' -and $script:Fake.Network) { throw 'Simulated network failure' }
            if ($cloud -eq 'Commercial' -and $script:Fake.Fail429 -gt 0) {
                $script:Fake.Fail429 = $script:Fake.Fail429 - 1
                throw 'Simulated HTTP 429'
            }
            $throttle = 0
            if ($cloud -eq 'Commercial' -and $script:Fake.Throttle -gt 0) {
                $script:Fake.Throttle = $script:Fake.Throttle - 1
                $throttle = 1
            }
            return [pscustomobject]@{
                IfExistsResult = [int]$script:Fake[$cloud]
                ThrottleStatus = $throttle
                EstsProperties = [pscustomobject]@{ DomainType = $script:Fake.DomainType }
                Credentials    = [pscustomobject]@{ FederationRedirectUrl = $script:Fake.FedUrl }
            }
        }
        if ($Uri -like '*getuserrealm*') {
            return [pscustomobject]@{ NameSpaceType = $script:Fake.Realm; FederationBrandName = $script:Fake.Brand; AuthURL = $script:Fake.AuthUrl; CloudInstanceName = 'microsoftonline.com' }
        }
        if ($Uri -like '*openid-configuration*') {
            if (-not $script:Fake.TenantId) { throw 'Simulated HTTP 400' }
            return [pscustomobject]@{
                token_endpoint      = "https://login.microsoftonline.com/$($script:Fake.TenantId)/oauth2/token"
                tenant_region_scope = 'NA'
                cloud_instance_name = 'microsoftonline.com'
            }
        }
        throw "Unexpected request to $Uri"
    }
    function Get-CallCount {
        param([string]$Pattern)
        return @($script:Fake.Calls | Where-Object { $_ -like $Pattern }).Count
    }
}

Describe 'Script file' {
    It 'parses without syntax errors' {
        $script:ParseErrors | Should -BeNullOrEmpty
    }
    It 'has a well-formed window layout' {
        { [xml](Get-AssignedString 'Xaml') } | Should -Not -Throw
    }
    It 'declares version 2.0.0' {
        Get-AssignedString 'AppVersion' | Should -Be '2.0.0'
    }
    It 'never sends a wildcard CORS header' {
        $script:SourceText | Should -Not -Match 'Access-Control-Allow-Origin'
    }
    It 'contains only ASCII characters, so Windows PowerShell 5.1 reads it correctly' {
        $script:SourceText | Should -Not -Match '[^\x00-\x7F]'
    }
}

Describe 'Email validation' {
    It 'accepts <Email>' -TestCases @(
        @{ Email = 'user@contoso.com' }
        @{ Email = 'first.last+tag@sub.contoso.co.uk' }
        @{ Email = 'a_b-c%d@x.io' }
        @{ Email = '  padded@contoso.com  ' }
    ) {
        param($Email)
        Test-EmailFormat $Email | Should -BeTrue
    }
    It 'rejects "<Email>"' -TestCases @(
        @{ Email = '' }
        @{ Email = 'user' }
        @{ Email = 'user@' }
        @{ Email = 'user@contoso' }
        @{ Email = 'user name@contoso.com' }
        @{ Email = '@contoso.com' }
        @{ Email = 'user@contoso.c' }
    ) {
        param($Email)
        Test-EmailFormat $Email | Should -BeFalse
    }
}

Describe 'Lookup and classification' {
    BeforeEach {
        Reset-Fake
        Mock Invoke-RestMethod { Invoke-FakeMicrosoft -Uri $Uri }
        Mock Start-Sleep { }
        Mock Get-HttpStatus { if ("$($ErrorRecord.Exception.Message)" -match 'HTTP (\d{3})') { [int]$Matches[1] } else { 0 } }
        $CoreStats = (New-SharedState).Stats
    }

    It 'classifies a managed work account as Entra in the Commercial cloud' {
        $r = Get-AccountCheck -Email 'user@contoso.com'
        $r.Category     | Should -Be 'Entra'
        $r.AccountType  | Should -Be 'Likely Entra Microsoft Account (Work/School).'
        $r.FoundInCloud | Should -Be 'Commercial'
        $r.CloudURL     | Should -Be 'https://login.microsoftonline.com'
        $r.Status       | Should -Be 'OK'
        $r.StatusLabel  | Should -Be 'Definitive'
        $r.Organization | Should -Be 'Contoso'
        $r.TenantId     | Should -Be '11111111-2222-3333-4444-555555555555'
        $r.TenantRegion | Should -Be 'North America (NA)'
    }

    It 'produces the four-line result format' {
        $lines = (Get-AccountCheck -Email 'user@contoso.com').Summary -split "`r`n"
        $lines.Count | Should -Be 4
        $lines[0] | Should -Be 'Username: user@contoso.com'
        $lines[1] | Should -Be 'FoundInCloud: Commercial'
        $lines[2] | Should -Be 'Cloud URL: https://login.microsoftonline.com'
        $lines[3] | Should -Be 'Account Type: Likely Entra Microsoft Account (Work/School).'
    }

    It 'detects federation from GetUserRealm and names the provider' {
        $script:Fake.Realm   = 'Federated'
        $script:Fake.AuthUrl = 'https://contoso.okta.com/app/office365/sso/wsfed/passive'
        $r = Get-AccountCheck -Email 'user@contoso.com'
        $r.Category              | Should -Be 'Federated'
        $r.FederationRedirectUrl | Should -Be $script:Fake.AuthUrl
        $r.FederationProvider    | Should -BeLike 'Okta*'
        ($r.Summary -split "`r`n").Count | Should -Be 5
    }

    It 'detects federation from EstsProperties.DomainType 4 alone' {
        $script:Fake.DomainType = 4
        (Get-AccountCheck -Email 'user@contoso.com').Category | Should -Be 'Federated'
    }

    It 'detects federation from FederationRedirectUrl alone' {
        $script:Fake.FedUrl = 'https://sso.godaddy.com/?realm=pass'
        $r = Get-AccountCheck -Email 'user@contoso.com'
        $r.Category           | Should -Be 'Federated'
        $r.FederationProvider | Should -BeLike 'GoDaddy*'
    }

    It 'classifies <Code> for a personal domain as MSA without tenant lookups' -TestCases @(@{ Code = 0 }, @{ Code = 5 }) {
        param($Code)
        $script:Fake.Commercial = $Code
        $r = Get-AccountCheck -Email 'someone@outlook.com'
        $r.Category | Should -Be 'MSA'
        $r.CloudURL | Should -Be 'https://login.live.com'
        Get-CallCount '*getuserrealm*'         | Should -Be 0
        Get-CallCount '*openid-configuration*' | Should -Be 0
    }

    It 'classifies code 5 for a work domain as Federated' {
        $script:Fake.Commercial = 5
        (Get-AccountCheck -Email 'user@contoso.com').Category | Should -Be 'Federated'
    }

    It 'classifies code 6 as Dual identity' {
        $script:Fake.Commercial = 6
        (Get-AccountCheck -Email 'user@contoso.com').AccountType | Should -Be 'Dual Identity (Work/School & Personal) Account.'
    }

    It 'finds a GCC High account when Commercial says not found' {
        $script:Fake.Commercial = 1
        $script:Fake.GCCHigh    = 0
        $r = Get-AccountCheck -Email 'user@agency.gov'
        $r.Category     | Should -Be 'GCCHigh'
        $r.FoundInCloud | Should -Be 'GCCHigh'
        $r.CloudURL     | Should -Be 'https://login.microsoftonline.us'
        Get-CallCount '*microsoftonline.cn*' | Should -Be 0
    }

    It 'finds a China account after GCC High says not found' {
        $script:Fake.Commercial = 1
        $script:Fake.China      = 5
        $r = Get-AccountCheck -Email 'user@contoso.cn'
        $r.Category | Should -Be 'China'
        $r.CloudURL | Should -Be 'https://login.partner.microsoftonline.cn'
    }

    It 'does not check sovereign clouds for personal domains' {
        $script:Fake.Commercial = 1
        $null = Get-AccountCheck -Email 'someone@hotmail.com'
        Get-CallCount '*microsoftonline.us*' | Should -Be 0
    }

    It 'reports a definitive "Not a Microsoft Account" when the domain has no tenant' {
        $script:Fake.Commercial = 1
        $script:Fake.TenantId   = ''
        $script:Fake.Realm      = 'Unknown'
        $r = Get-AccountCheck -Email 'user@example.org'
        $r.Category     | Should -Be 'NotMicrosoft'
        $r.Status       | Should -Be 'OK'
        $r.FoundInCloud | Should -Be 'N/A'
    }

    It 'reports Inconclusive when the address is not found but the domain has a tenant' {
        $script:Fake.Commercial = 1
        $r = Get-AccountCheck -Email 'alias@contoso.com'
        $r.Category   | Should -Be 'NotMicrosoft'
        $r.Status     | Should -Be 'Inconclusive'
        $r.Note       | Should -BeLike '*Microsoft tenant*'
        $r.NextAction | Should -Not -BeNullOrEmpty
        $CoreStats['Inconclusive'] | Should -Be 1
    }

    It 'reports an unknown result code as Inconclusive' {
        $script:Fake.Commercial = 9
        $r = Get-AccountCheck -Email 'user@contoso.com'
        $r.Status      | Should -Be 'Inconclusive'
        $r.AccountType | Should -BeLike 'Unknown Result Code: 9*'
    }

    It 'marks an invalid address without calling Microsoft' {
        $r = Get-AccountCheck -Email 'not-an-address'
        $r.Status | Should -Be 'Invalid'
        $script:Fake.Calls.Count | Should -Be 0
    }
}

Describe 'Throttling, retries and outages' {
    BeforeEach {
        Reset-Fake
        Mock Invoke-RestMethod { Invoke-FakeMicrosoft -Uri $Uri }
        Mock Start-Sleep { }
        Mock Get-HttpStatus { if ("$($ErrorRecord.Exception.Message)" -match 'HTTP (\d{3})') { [int]$Matches[1] } else { 0 } }
        $CoreStats = (New-SharedState).Stats
    }

    It 'retries once after ThrottleStatus and then succeeds' {
        $script:Fake.Throttle = 1
        $r = Get-AccountCheck -Email 'user@contoso.com'
        $r.Status | Should -Be 'OK'
        Get-CallCount '*microsoftonline.com/common/GetCredentialType' | Should -Be 2
        $CoreStats['Retries']        | Should -Be 1
        $CoreStats['ThrottleEvents'] | Should -Be 1
    }

    It 'flags the result as Throttled when every attempt is throttled' {
        $script:Fake.Throttle = 3
        $r = Get-AccountCheck -Email 'user@contoso.com'
        $r.Status      | Should -Be 'Throttled'
        $r.StatusLabel | Should -Be 'Throttled'
        $r.NextAction  | Should -BeLike '*Wait a few minutes*'
        Get-CallCount '*microsoftonline.com/common/GetCredentialType' | Should -Be 3
    }

    It 'retries after HTTP 429 and then succeeds' {
        $script:Fake.Fail429 = 1
        (Get-AccountCheck -Email 'user@contoso.com').Status | Should -Be 'OK'
        $CoreStats['Retries'] | Should -Be 1
    }

    It 'reports Temporarily unavailable when HTTP 429 never clears' {
        $script:Fake.Fail429 = 5
        $r = Get-AccountCheck -Email 'user@contoso.com'
        $r.Status      | Should -Be 'Unavailable'
        $r.StatusLabel | Should -Be 'Temporarily unavailable'
        $CoreStats['Unavailable'] | Should -Be 1
    }

    It 'reports Temporarily unavailable on a network failure, with a next step' {
        $script:Fake.Network = $true
        $r = Get-AccountCheck -Email 'user@contoso.com'
        $r.Status     | Should -Be 'Unavailable'
        $r.Category   | Should -Be 'Error'
        $r.NextAction | Should -BeLike '*network*'
        $CoreStats['LastError'] | Should -BeLike '*user@contoso.com*'
    }

    It 'backs off for about 2 and then 4 seconds' {
        Get-BackoffMs 1 | Should -BeGreaterOrEqual 2000
        Get-BackoffMs 1 | Should -BeLessThan 2500
        Get-BackoffMs 2 | Should -BeGreaterOrEqual 4000
        Get-BackoffMs 2 | Should -BeLessThan 4500
    }

    It 'records per-endpoint timings' {
        $null = Get-AccountCheck -Email 'user@contoso.com'
        $snapshot = Get-CoreStatsSnapshot -Stats $CoreStats
        $snapshot.LiveLookups | Should -Be 1
        ($snapshot.Endpoints | Where-Object { $_.Endpoint -eq 'GetCredentialType (Commercial)' }).Calls | Should -Be 1
        ($snapshot.Endpoints | Where-Object { $_.Endpoint -eq 'GetUserRealm' }).Calls | Should -Be 1
    }
}

Describe 'Cache' {
    BeforeEach {
        Reset-Fake
        Mock Invoke-RestMethod { Invoke-FakeMicrosoft -Uri $Uri }
        Mock Start-Sleep { }
        Mock Get-HttpStatus { if ("$($ErrorRecord.Exception.Message)" -match 'HTTP (\d{3})') { [int]$Matches[1] } else { 0 } }
        $CoreStats = (New-SharedState).Stats
        $cache = [hashtable]::Synchronized(@{})
    }

    It 'answers a repeat lookup from the cache' {
        $null = Get-AccountCheckCached -Email 'user@contoso.com' -Cache $cache
        $before = $script:Fake.Calls.Count
        $second = Get-AccountCheckCached -Email 'USER@contoso.com' -Cache $cache
        $second.Source | Should -Be 'Cache'
        $second.Category | Should -Be 'Entra'
        $script:Fake.Calls.Count | Should -Be $before
        $CoreStats['CacheHits'] | Should -Be 1
    }

    It 'looks up again after 24 hours' {
        $null = Get-AccountCheckCached -Email 'user@contoso.com' -Cache $cache
        $cache['user@contoso.com'].Time = [datetime]::UtcNow.AddHours(-25)
        (Get-AccountCheckCached -Email 'user@contoso.com' -Cache $cache).Source | Should -Be 'Live'
    }

    It 'skips the cache when asked' {
        $null = Get-AccountCheckCached -Email 'user@contoso.com' -Cache $cache
        (Get-AccountCheckCached -Email 'user@contoso.com' -Cache $cache -SkipCache).Source | Should -Be 'Live'
    }

    It 'does not cache throttled results' {
        $script:Fake.Throttle = 3
        $null = Get-AccountCheckCached -Email 'user@contoso.com' -Cache $cache
        $cache.Count | Should -Be 0
    }

    It 'does not cache unavailable results' {
        $script:Fake.Network = $true
        $null = Get-AccountCheckCached -Email 'user@contoso.com' -Cache $cache
        $cache.Count | Should -Be 0
    }

    It 'caches inconclusive results' {
        $script:Fake.Commercial = 1
        $null = Get-AccountCheckCached -Email 'alias@contoso.com' -Cache $cache
        $cache.Count | Should -Be 1
    }
}

Describe 'Helpers' {
    It 'names the federation provider for <Url>' -TestCases @(
        @{ Url = 'https://contoso.okta.com/sso'; Expected = 'Okta (contoso.okta.com)' }
        @{ Url = 'https://sso.godaddy.com/?realm=pass'; Expected = 'GoDaddy (sso.godaddy.com)' }
        @{ Url = 'https://login.contoso.com/adfs/ls/'; Expected = 'AD FS (login.contoso.com)' }
        @{ Url = 'https://sts.contoso.com/'; Expected = 'AD FS, likely (sts.contoso.com)' }
        @{ Url = 'https://idp.example.net/login'; Expected = 'idp.example.net' }
        @{ Url = ''; Expected = '' }
    ) {
        param($Url, $Expected)
        Get-FederationProvider $Url | Should -Be $Expected
    }

    It 'formats tenant regions' {
        Format-TenantRegion -Scope 'NA' -SubScope 'GCC' | Should -Be 'North America (NA), GCC'
        Format-TenantRegion -Scope 'EU' -SubScope ''    | Should -Be 'Europe (EU)'
        Format-TenantRegion -Scope 'ZZ' -SubScope ''    | Should -Be 'ZZ'
        Format-TenantRegion -Scope ''   -SubScope ''    | Should -Be ''
    }

    It 'gives every non-definitive status a next step' -TestCases @(
        @{ Status = 'Inconclusive' }, @{ Status = 'Throttled' }, @{ Status = 'Unavailable' }, @{ Status = 'Invalid' }
    ) {
        param($Status)
        (Get-StatusAdvice -Status $Status).Action | Should -Not -BeNullOrEmpty
    }

    It 'gives definitive results no next step' {
        (Get-StatusAdvice -Status 'OK').Label  | Should -Be 'Definitive'
        (Get-StatusAdvice -Status 'OK').Action | Should -Be ''
    }

    It 'splits, de-duplicates and validates a batch list' {
        $tokens = Get-BatchTokens -Text "a@x.com, b@y.com;A@X.com`r`nnot-an-address <c@z.org>"
        @($tokens.Valid) | Should -Be @('a@x.com', 'b@y.com', 'c@z.org')
        @($tokens.Invalid) | Should -Be @('not-an-address')
        $tokens.Duplicates | Should -Be 1
    }

    It 'counts results by account type and status, and can subtract' {
        $counts = New-Counts
        Add-ToCounts -Counts $counts -Category 'China' -Status 'OK'
        Add-ToCounts -Counts $counts -Category 'Entra' -Status 'Throttled'
        Add-ToCounts -Counts $counts -Category 'Error' -Status 'Error'
        $counts.Sovereign   | Should -Be 1
        $counts.Throttled   | Should -Be 1
        $counts.Unavailable | Should -Be 1
        Add-ToCounts -Counts $counts -Category 'Entra' -Status 'Throttled' -Delta -1
        $counts.Throttled | Should -Be 0
        $counts.Total     | Should -Be 2
    }
}

Describe 'History filters' {
    BeforeAll {
        $script:Item = [pscustomobject]@{
            Timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'); Username = 'a@contoso.com'; FoundInCloud = 'Commercial'
            AccountType = 'Federated Entra ID (Work/School) Account.'; Category = 'Federated'; Organization = 'Contoso'
            FederationProvider = 'Okta (contoso.okta.com)'; Status = 'Error'; StatusLabel = 'Error'; Source = 'Live'; Origin = 'Batch'
        }
        function New-Filter {
            param([hashtable]$Change = @{})
            $state = @{ Text = ''; Type = 'All'; Cloud = 'All'; Status = 'All'; Federation = 'All'; Source = 'All'; Since = $null }
            foreach ($key in $Change.Keys) { $state[$key] = $Change[$key] }
            $script:HistoryFilterState = $state
        }
    }
    It 'keeps the item when <Name>' -TestCases @(
        @{ Name = 'no filter is set'; Change = @{} }
        @{ Name = 'the account type matches'; Change = @{ Type = 'Federated' } }
        @{ Name = 'the cloud matches'; Change = @{ Cloud = 'Commercial' } }
        @{ Name = 'an old "Error" status is filtered as Unavailable'; Change = @{ Status = 'Unavailable' } }
        @{ Name = 'federated only is chosen'; Change = @{ Federation = 'Yes' } }
        @{ Name = 'the date is today'; Change = @{ Since = (Get-Date).Date } }
        @{ Name = 'the search matches the organization'; Change = @{ Text = 'contoso' } }
    ) {
        param($Name, $Change)
        New-Filter $Change
        Test-HistoryItem -Item $script:Item | Should -BeTrue
    }
    It 'hides the item when <Name>' -TestCases @(
        @{ Name = 'the account type differs'; Change = @{ Type = 'Entra' } }
        @{ Name = 'not federated is chosen'; Change = @{ Federation = 'No' } }
        @{ Name = 'only cache answers are shown'; Change = @{ Source = 'Cache' } }
        @{ Name = 'the date is in the future'; Change = @{ Since = (Get-Date).Date.AddDays(1) } }
        @{ Name = 'the search does not match'; Change = @{ Text = 'zzz' } }
    ) {
        param($Name, $Change)
        New-Filter $Change
        Test-HistoryItem -Item $script:Item | Should -BeFalse
    }
}

Describe 'Export' {
    BeforeAll {
        $script:Rows = @(
            [pscustomobject]@{ Timestamp = '2026-11-17 10:00:00'; Username = 'quote"and,comma@contoso.com'; FoundInCloud = 'Commercial'; CloudURL = 'https://login.microsoftonline.com'
                FederationRedirectUrl = ''; AccountType = 'Likely Entra Microsoft Account (Work/School).'; Category = 'Entra'; Organization = 'Contoso'
                TenantId = '11111111-2222-3333-4444-555555555555'; TenantRegion = 'North America (NA)'; NamespaceType = 'Managed'; FederationProvider = ''
                Status = 'OK'; StatusLabel = 'Definitive'; NextAction = ''; Note = ''; Source = 'Live'; Origin = 'Single' }
            [pscustomobject]@{ Timestamp = '2026-11-17 10:01:00'; Username = '<script>@contoso.com'; FoundInCloud = 'N/A'; CloudURL = ''
                FederationRedirectUrl = ''; AccountType = 'Not a Microsoft Account.'; Category = 'NotMicrosoft'; Organization = ''
                TenantId = ''; TenantRegion = ''; NamespaceType = ''; FederationProvider = ''
                Status = 'Throttled'; StatusLabel = 'Throttled'; NextAction = 'Wait'; Note = ''; Source = 'Live'; Origin = 'Batch' }
        )
    }
    It 'round-trips rows through CSV with every column' {
        $path = Join-Path $TestDrive 'export.csv'
        Get-ExportRows $script:Rows | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding UTF8
        $back = @(Import-Csv -LiteralPath $path)
        $back.Count | Should -Be 2
        $back[0].Username    | Should -Be 'quote"and,comma@contoso.com'
        $back[0].TenantId    | Should -Be '11111111-2222-3333-4444-555555555555'
        $back[1].StatusLabel | Should -Be 'Throttled'
        $back[0].PSObject.Properties.Name | Should -Contain 'NextAction'
    }
    It 'HTML-encodes values in the report' {
        $html = New-HtmlReport -Items $script:Rows -Title 'Test'
        $html | Should -Match '&lt;script&gt;@contoso.com'
        $html | Should -Not -Match '<script>@'
        $html | Should -Match 'class="warn">Throttled'
    }
}

Describe 'Local API server' {
    BeforeAll {
        # The real API script, with Microsoft replaced by a stub that always answers "managed Entra account".
        $stub = @'
function Invoke-CoreRequest {
    param([string]$Uri, [string]$Method = 'Get', [string]$Body = '', [string]$Endpoint = '')
    if ($Uri -like '*GetCredentialType*') { return [pscustomobject]@{ IfExistsResult = 0; ThrottleStatus = 0 } }
    if ($Uri -like '*getuserrealm*') { return [pscustomobject]@{ NameSpaceType = 'Managed'; FederationBrandName = 'Contoso' } }
    return [pscustomobject]@{ token_endpoint = 'https://login.microsoftonline.com/11111111-2222-3333-4444-555555555555/oauth2/token'; tenant_region_scope = 'NA' }
}
'@
        $probe = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Loopback, 0)
        $probe.Start(); $script:ApiPort = $probe.LocalEndpoint.Port; $probe.Stop()

        $script:ApiShared = New-SharedState
        $script:ApiShared['RateLimitPerMinute'] = 0
        $script:ApiShared['BatchDelayMs'] = 0
        $script:Listener = [System.Net.HttpListener]::new()
        $script:Listener.Prefixes.Add("http://localhost:$($script:ApiPort)/")
        $script:Listener.Start()

        $script:ApiShell = [powershell]::Create()
        $script:ApiShell.Runspace = [runspacefactory]::CreateRunspace()
        $script:ApiShell.Runspace.Open()
        [void]$script:ApiShell.AddScript((Get-AssignedScriptBlockText 'ApiScript'))
        [void]$script:ApiShell.AddParameter('CoreText', $script:CoreText + "`n" + $stub)
        [void]$script:ApiShell.AddParameter('Listener', $script:Listener)
        [void]$script:ApiShell.AddParameter('Shared', $script:ApiShared)
        [void]$script:ApiShell.AddParameter('StartedUtc', [datetime]::UtcNow)
        [void]$script:ApiShell.AddParameter('TtlHours', 24)
        [void]$script:ApiShell.AddParameter('Version', 'test')
        [void]$script:ApiShell.AddParameter('WebPage', (Get-AssignedString 'WebPageTemplate'))
        [void]$script:ApiShell.AddParameter('Port', $script:ApiPort)
        $script:ApiHandle = $script:ApiShell.BeginInvoke()

        $script:RunnerControl = [hashtable]::Synchronized(@{ Stop = $false })
        $script:RunnerShell = [powershell]::Create()
        $script:RunnerShell.Runspace = [runspacefactory]::CreateRunspace()
        $script:RunnerShell.Runspace.Open()
        [void]$script:RunnerShell.AddScript((Get-AssignedScriptBlockText 'JobRunnerScript'))
        [void]$script:RunnerShell.AddParameter('CoreText', $script:CoreText + "`n" + $stub)
        [void]$script:RunnerShell.AddParameter('Shared', $script:ApiShared)
        [void]$script:RunnerShell.AddParameter('TtlHours', 24)
        [void]$script:RunnerShell.AddParameter('Control', $script:RunnerControl)
        $script:RunnerHandle = $script:RunnerShell.BeginInvoke()

        Add-Type -AssemblyName System.Net.Http
        $script:Http = [System.Net.Http.HttpClient]::new()
        $script:SelfOrigin = "http://localhost:$($script:ApiPort)"

        function Send-Api {
            param([string]$Path, [string]$Method = 'GET', [hashtable]$Headers = @{}, [string]$Body = '')
            $request = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::new($Method), "$($script:SelfOrigin)$Path")
            foreach ($name in $Headers.Keys) { [void]$request.Headers.TryAddWithoutValidation($name, [string]$Headers[$name]) }
            if ($Body) { $request.Content = [System.Net.Http.StringContent]::new($Body, [System.Text.Encoding]::UTF8, 'application/json') }
            $response = $script:Http.SendAsync($request).GetAwaiter().GetResult()
            $text = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
            $json = $null
            try { $json = ConvertFrom-Json -InputObject $text } catch { }
            return [pscustomobject]@{
                Status      = [int]$response.StatusCode
                Text        = $text
                Json        = $json
                Headers     = $response.Headers
                ContentType = [string]$response.Content.Headers.ContentType
                RetryAfter  = [string]$response.Headers.RetryAfter
            }
        }
    }
    AfterAll {
        try { $script:Listener.Stop(); $script:Listener.Close() } catch { }
        try { [void]$script:ApiHandle.AsyncWaitHandle.WaitOne(3000); $script:ApiShell.Dispose() } catch { }
        try { $script:RunnerControl.Stop = $true; [void]$script:RunnerHandle.AsyncWaitHandle.WaitOne(3000); $script:RunnerShell.Dispose() } catch { }
        try { $script:Http.Dispose() } catch { }
    }

    Context 'responses' {
        It 'answers /api/health' {
            $r = Send-Api '/api/health'
            $r.Status | Should -Be 200
            $r.Json.status | Should -Be 'ok'
            $r.Json.tokenRequired | Should -BeFalse
        }
        It 'checks an address' {
            $r = Send-Api '/api/check?email=user%40contoso.com'
            $r.Status | Should -Be 200
            $r.Json.category     | Should -Be 'Entra'
            $r.Json.organization | Should -Be 'Contoso'
            $r.Json.statusLabel  | Should -Be 'Definitive'
            $script:ApiShared.ApiResults.Count | Should -BeGreaterThan 0
        }
        It 'sends no CORS header and sends nosniff' {
            $r = Send-Api '/api/health'
            $r.Headers.Contains('Access-Control-Allow-Origin') | Should -BeFalse
            $r.Headers.Contains('X-Content-Type-Options') | Should -BeTrue
        }
        It 'rejects a missing address with 400' {
            (Send-Api '/api/check').Status | Should -Be 400
        }
        It 'rejects an invalid address with 400' {
            $r = Send-Api '/api/check?email=nope'
            $r.Status | Should -Be 400
            $r.Json.status | Should -Be 'Invalid'
        }
        It 'checks a batch given as an array' {
            $r = Send-Api '/api/batch' -Method 'POST' -Body '["a@contoso.com","b@contoso.com"]'
            $r.Status | Should -Be 200
            $r.Json.count | Should -Be 2
        }
        It 'checks a batch given as {"emails": [...]}' {
            (Send-Api '/api/batch' -Method 'POST' -Body '{"emails":["c@contoso.com"]}').Json.count | Should -Be 1
        }
        It 'rejects more than 200 addresses with 413' {
            $many = ConvertTo-Json -InputObject @(1..201 | ForEach-Object { "u$_@contoso.com" })
            (Send-Api '/api/batch' -Method 'POST' -Body $many).Status | Should -Be 413
        }
        It 'rejects an empty batch with 400' {
            (Send-Api '/api/batch' -Method 'POST' -Body '{}').Status | Should -Be 400
        }
        It 'rejects GET on /api/batch with 405' {
            (Send-Api '/api/batch').Status | Should -Be 405
        }
        It 'returns 404 for an unknown path' {
            (Send-Api '/api/nothing').Status | Should -Be 404
        }
        It 'serves the web page with protective headers' {
            $r = Send-Api '/'
            $r.Status | Should -Be 200
            $r.ContentType | Should -BeLike 'text/html*'
            $r.Text | Should -Match 'Microsoft Account Checker'
            $r.Headers.Contains('X-Frame-Options') | Should -BeTrue
            $r.Headers.Contains('Content-Security-Policy') | Should -BeTrue
        }
        It 'returns history with search' {
            [void]$script:ApiShared.History.Add([pscustomobject]@{ Timestamp = '2026-11-17 10:00:00'; Username = 'one@contoso.com'; Category = 'Entra'; AccountType = 'x'; FoundInCloud = 'Commercial'; Organization = 'Contoso'; Origin = 'Single'; Status = 'OK'; StatusLabel = 'Definitive' })
            [void]$script:ApiShared.History.Add([pscustomobject]@{ Timestamp = '2026-11-17 10:01:00'; Username = 'two@fabrikam.com'; Category = 'MSA'; AccountType = 'y'; FoundInCloud = 'Commercial'; Organization = 'Fabrikam'; Origin = 'API'; Status = 'OK'; StatusLabel = 'Definitive' })
            $all = Send-Api '/api/history'
            $all.Json.total | Should -Be 2
            $all.Json.items[0].username | Should -Be 'two@fabrikam.com'
            (Send-Api '/api/history?search=contoso').Json.matched | Should -Be 1
        }
        It 'filters history by <Name>' -TestCases @(
            @{ Name = 'account type';         Query = '&type=Federated';    Expected = 'f1@filtertest.com' }
            @{ Name = 'sovereign cloud type'; Query = '&type=Sovereign';    Expected = 'f2@filtertest.com' }
            @{ Name = 'cloud';                Query = '&cloud=GCCHigh';     Expected = 'f2@filtertest.com' }
            @{ Name = 'status';               Query = '&status=Throttled';  Expected = 'f1@filtertest.com' }
            @{ Name = 'old Error status';     Query = '&status=Unavailable'; Expected = 'f3@filtertest.com' }
            @{ Name = 'federated only';       Query = '&federation=yes';    Expected = 'f1@filtertest.com' }
            @{ Name = 'source';               Query = '&source=Cache';      Expected = 'f2@filtertest.com' }
            @{ Name = 'date';                 Query = '&days=30&type=Federated'; Expected = 'f1@filtertest.com' }
        ) {
            param($Name, $Query, $Expected)
            if (-not ($script:ApiShared.History | Where-Object { $_.Username -eq 'f1@filtertest.com' })) {
                $now = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
                [void]$script:ApiShared.History.Add([pscustomobject]@{ Timestamp = $now; Username = 'f1@filtertest.com'; Category = 'Federated'; AccountType = 'Federated'; FoundInCloud = 'Commercial'; Organization = 'Filter'; Origin = 'Batch'; Status = 'Throttled'; StatusLabel = 'Throttled'; Source = 'Live'; FederationProvider = 'Okta (filter.okta.com)' })
                [void]$script:ApiShared.History.Add([pscustomobject]@{ Timestamp = '2020-01-01 00:00:00'; Username = 'f2@filtertest.com'; Category = 'GCCHigh'; AccountType = 'GCC'; FoundInCloud = 'GCCHigh'; Organization = 'Filter'; Origin = 'API'; Status = 'OK'; StatusLabel = 'Definitive'; Source = 'Cache'; FederationProvider = '' })
                [void]$script:ApiShared.History.Add([pscustomobject]@{ Timestamp = '2020-01-01 00:00:00'; Username = 'f3@filtertest.com'; Category = 'Error'; AccountType = 'Network'; FoundInCloud = 'N/A'; Organization = ''; Origin = 'Single'; Status = 'Error'; StatusLabel = 'Error'; Source = 'Live'; FederationProvider = '' })
            }
            $r = Send-Api ("/api/history?search=filtertest$Query")
            $r.Status | Should -Be 200
            $r.Json.matched | Should -Be 1
            $r.Json.items[0].username | Should -Be $Expected
        }
        It 'keeps the statistic counts for the whole history when filtering' {
            $all = (Send-Api '/api/history').Json.counts.Total
            (Send-Api '/api/history?type=MSA').Json.counts.Total | Should -Be $all
        }
    }

    Context 'requests from other websites' {
        It 'refuses Sec-Fetch-Site <Site>' -TestCases @(@{ Site = 'cross-site' }, @{ Site = 'same-site' }) {
            param($Site)
            (Send-Api '/api/check?email=user%40contoso.com' -Headers @{ 'Sec-Fetch-Site' = $Site }).Status | Should -Be 403
        }
        It 'refuses a foreign Origin header' {
            (Send-Api '/api/batch' -Method 'POST' -Body '["a@contoso.com"]' -Headers @{ Origin = 'https://evil.example' }).Status | Should -Be 403
        }
        It 'accepts its own origin' {
            (Send-Api '/api/health' -Headers @{ Origin = $script:SelfOrigin; 'Sec-Fetch-Site' = 'same-origin' }).Status | Should -Be 200
        }
        It 'still lets a link from another site open the page' {
            (Send-Api '/' -Headers @{ 'Sec-Fetch-Site' = 'cross-site'; 'Sec-Fetch-Mode' = 'navigate' }).Status | Should -Be 200
        }
        It 'counts refused requests' {
            [int]$script:ApiShared['RejectedCount'] | Should -BeGreaterThan 0
        }
    }

    Context 'operations endpoints' {
        It 'answers /health without a token' {
            $r = Send-Api '/health'
            $r.Status | Should -Be 200
            $r.Json.status | Should -Be 'ok'
        }
        It 'serves Prometheus metrics' {
            $r = Send-Api '/metrics'
            $r.Status | Should -Be 200
            $r.ContentType | Should -BeLike 'text/plain*'
            $r.Text | Should -Match '(?m)^mac_api_requests_total \d+'
            $r.Text | Should -Match '(?m)^# TYPE mac_lookups_total counter'
            $r.Text | Should -Match 'mac_lookups_total\{source="live"\}'
        }
        It 'serves a valid OpenAPI 3 document' {
            $r = Send-Api '/api/openapi.json'
            $r.Status | Should -Be 200
            $r.Json.openapi | Should -Be '3.0.3'
            $r.Json.paths.PSObject.Properties.Name | Should -Contain '/api/check'
            $r.Json.paths.PSObject.Properties.Name | Should -Contain '/api/jobs/{id}'
            $r.Json.components.securitySchemes.apiKey.name | Should -Be 'X-API-Key'
        }
        It 'filters history to changed addresses' {
            [void]$script:ApiShared.History.Add([pscustomobject]@{ Timestamp = '2026-11-19 10:00:00'; Username = 'moved@changedtest.com'; Category = 'Federated'; AccountType = 'x'; FoundInCloud = 'Commercial'; Organization = ''; Origin = 'API'; Status = 'OK'; StatusLabel = 'Definitive'; Changes = 'now federated' })
            [void]$script:ApiShared.History.Add([pscustomobject]@{ Timestamp = '2026-11-19 10:00:00'; Username = 'same@changedtest.com'; Category = 'Entra'; AccountType = 'y'; FoundInCloud = 'Commercial'; Organization = ''; Origin = 'API'; Status = 'OK'; StatusLabel = 'Definitive'; Changes = '' })
            $r = Send-Api '/api/history?search=changedtest&status=Changed'
            $r.Json.matched | Should -Be 1
            $r.Json.items[0].changes | Should -Be 'now federated'
        }
    }

    Context 'background jobs' {
        It 'runs a submitted job to completion' {
            $submit = Send-Api '/api/jobs' -Method 'POST' -Body '["j1@contoso.com","j2@contoso.com","j3@contoso.com"]'
            $submit.Status | Should -Be 202
            $id = $submit.Json.id
            $submit.Json.statusUrl | Should -Be "/api/jobs/$id"
            $deadline = (Get-Date).AddSeconds(20)
            do { Start-Sleep -Milliseconds 300; $job = (Send-Api "/api/jobs/$id").Json } while ($job.status -ne 'done' -and (Get-Date) -lt $deadline)
            $job.status | Should -Be 'done'
            $job.done | Should -Be 3
            @($job.results).Count | Should -Be 3
            @($job.results)[0].category | Should -Be 'Entra'
        }
        It 'lists jobs' {
            (Send-Api '/api/jobs').Json.count | Should -BeGreaterThan 0
        }
        It 'returns 404 for an unknown job' {
            (Send-Api '/api/jobs/doesnotexist').Status | Should -Be 404
        }
        It 'rejects more than 5,000 addresses with 413' {
            $many = ConvertTo-Json -InputObject @(1..5001 | ForEach-Object { "u$_@contoso.com" })
            (Send-Api '/api/jobs' -Method 'POST' -Body $many).Status | Should -Be 413
        }
        It 'accepts a cancel request' {
            $id = (Send-Api '/api/jobs' -Method 'POST' -Body '["c1@contoso.com"]').Json.id
            (Send-Api "/api/jobs/$id" -Method 'DELETE').Json.cancelRequested | Should -BeTrue
        }
    }

    Context 'rate limit' {
        BeforeAll { $script:ApiShared['RateLimitPerMinute'] = 3 }
        AfterAll { $script:ApiShared['RateLimitPerMinute'] = 0 }
        It 'answers 429 with Retry-After once the limit is reached' {
            1..3 | ForEach-Object { (Send-Api '/api/health').Status | Should -Be 200 }
            $limited = Send-Api '/api/health'
            $limited.Status | Should -Be 429
            $limited.RetryAfter | Should -Not -BeNullOrEmpty
        }
        It 'never limits /health' {
            (Send-Api '/health').Status | Should -Be 200
        }
    }

    Context 'API token' {
        BeforeAll { $script:ApiShared['ApiToken'] = '0123456789abcdef0123456789abcdef' }
        AfterAll { $script:ApiShared['ApiToken'] = '' }

        It 'still answers /health without the token' {
            (Send-Api '/health').Status | Should -Be 200
        }
        It 'protects /metrics with the token' {
            (Send-Api '/metrics').Status | Should -Be 401
        }
        It 'refuses requests without the token' {
            (Send-Api '/api/health').Status | Should -Be 401
        }
        It 'refuses a wrong token' {
            (Send-Api '/api/health' -Headers @{ 'X-API-Key' = 'wrong' }).Status | Should -Be 401
        }
        It 'accepts the token in X-API-Key' {
            (Send-Api '/api/health' -Headers @{ 'X-API-Key' = '0123456789abcdef0123456789abcdef' }).Status | Should -Be 200
        }
        It 'accepts the token as a Bearer token' {
            (Send-Api '/api/health' -Headers @{ Authorization = 'Bearer 0123456789abcdef0123456789abcdef' }).Status | Should -Be 200
        }
        It 'puts the token in the page opened from this computer' {
            (Send-Api '/').Text | Should -Match '0123456789abcdef0123456789abcdef'
        }
        It 'leaves the token out of a page opened from another site' {
            (Send-Api '/' -Headers @{ 'Sec-Fetch-Site' = 'cross-site'; 'Sec-Fetch-Mode' = 'navigate' }).Text | Should -Not -Match '0123456789abcdef0123456789abcdef'
        }
    }
}


Describe 'Colour contrast (WCAG 2.1)' {
    BeforeAll {
        function Get-ContrastRatio {
            param([string]$Foreground, [string]$Background)
            $luminance = {
                param($hex)
                $h = $hex.TrimStart('#')
                $h = $h.Substring($h.Length - 6)
                $channels = foreach ($i in 0, 2, 4) {
                    $v = [Convert]::ToInt32($h.Substring($i, 2), 16) / 255
                    if ($v -le 0.03928) { $v / 12.92 } else { [Math]::Pow(($v + 0.055) / 1.055, 2.4) }
                }
                0.2126 * $channels[0] + 0.7152 * $channels[1] + 0.0722 * $channels[2]
            }
            $a = & $luminance $Foreground
            $b = & $luminance $Background
            return ([Math]::Max($a, $b) + 0.05) / ([Math]::Min($a, $b) + 0.05)
        }
        $script:TextPairs = @(
            @('Text', 'Bg'), @('Text', 'Surface'), @('Muted', 'Surface'), @('Muted', 'Bg'), @('Muted', 'InputBg'),
            @('PrimaryText', 'Primary'), @('SecondaryText', 'Secondary'), @('DangerText', 'Danger'), @('Link', 'Surface'),
            @('HeaderText', 'HeaderStart'), @('HeaderText', 'HeaderEnd'), @('HeaderMuted', 'HeaderStart'), @('HeaderMuted', 'HeaderEnd'),
            @('CatEntra', 'Surface'), @('CatMSA', 'Surface'), @('CatFederated', 'Surface'), @('CatDual', 'Surface'), @('CatGCCHigh', 'Surface'),
            @('CatChina', 'Surface'), @('CatNotMicrosoft', 'Surface'), @('CatError', 'Surface'), @('CatUnknown', 'Surface'))
    }
    It 'meets 4.5:1 for every text colour in the <Theme> theme' -TestCases @(@{ Theme = 'light'; Dark = $false }, @{ Theme = 'dark'; Dark = $true }) {
        param($Theme, $Dark)
        $palette = Get-ThemePalette -Dark $Dark
        $failures = foreach ($pair in $script:TextPairs) {
            $ratio = Get-ContrastRatio -Foreground $palette[$pair[0]] -Background $palette[$pair[1]]
            if ($ratio -lt 4.5) { '{0} on {1}: {2:0.00}' -f $pair[0], $pair[1], $ratio }
        }
        $failures | Should -BeNullOrEmpty
    }
}

Describe 'Cache policy by outcome' {
    BeforeEach {
        Reset-Fake
        Mock Invoke-RestMethod { Invoke-FakeMicrosoft -Uri $Uri }
        Mock Start-Sleep { }
        Mock Get-HttpStatus { if ("$($ErrorRecord.Exception.Message)" -match 'HTTP (\d{3})') { [int]$Matches[1] } else { 0 } }
        $CoreStats = (New-SharedState).Stats
        $cache = [hashtable]::Synchronized(@{})
        $policy = @{ StableHours = 24.0; InconclusiveHours = 6.0; FailureMinutes = 5.0 }
    }
    It 'keeps throttled results for the short failure time' {
        $script:Fake.Throttle = 3
        $null = Get-AccountCheckCached -Email 'user@contoso.com' -Cache $cache -Policy $policy
        $cache['user@contoso.com'].TtlHours | Should -Be (5 / 60)
        (Get-AccountCheckCached -Email 'user@contoso.com' -Cache $cache -Policy $policy).Source | Should -Be 'Cache'
    }
    It 'uses the inconclusive time for inconclusive answers' {
        $script:Fake.Commercial = 1
        $null = Get-AccountCheckCached -Email 'alias@contoso.com' -Cache $cache -Policy $policy
        $cache['alias@contoso.com'].TtlHours | Should -Be 6
    }
    It 'never caches when the stable time is 0' {
        $null = Get-AccountCheckCached -Email 'user@contoso.com' -Cache $cache -Policy @{ StableHours = 0; InconclusiveHours = 0; FailureMinutes = 0 }
        $cache.Count | Should -Be 0
    }
    It 'reuses fresh answers but refreshes stale ones with MaxAgeHours' {
        $null = Get-AccountCheckCached -Email 'user@contoso.com' -Cache $cache -Policy $policy
        $cache['user@contoso.com'].Time = [datetime]::UtcNow.AddHours(-13)
        (Get-AccountCheckCached -Email 'user@contoso.com' -Cache $cache -Policy $policy -MaxAgeHours 24).Source | Should -Be 'Cache'
        (Get-AccountCheckCached -Email 'user@contoso.com' -Cache $cache -Policy $policy -MaxAgeHours 12).Source | Should -Be 'Live'
    }
}

Describe 'Change detection' {
    BeforeAll {
        function New-Result {
            param([string]$Category = 'Entra', [string]$Cloud = 'Commercial', [string]$Tenant = 't1', [string]$Provider = '', [string]$Status = 'OK')
            [pscustomobject]@{ Category = $Category; FoundInCloud = $Cloud; TenantId = $Tenant; FederationProvider = $Provider; Status = $Status }
        }
    }
    It 'reports nothing when nothing changed' {
        Get-ResultChanges -Previous (New-Result) -Current (New-Result) | Should -Be ''
    }
    It 'reports an account type change' {
        Get-ResultChanges -Previous (New-Result) -Current (New-Result -Category 'NotMicrosoft') | Should -BeLike 'account type Work/School (Entra ID) -> Not a Microsoft account*'
    }
    It 'reports a cloud change' {
        Get-ResultChanges -Previous (New-Result) -Current (New-Result -Category 'Entra' -Cloud 'GCCHigh') | Should -Match 'cloud Commercial -> GCCHigh'
    }
    It 'reports a tenant change' {
        Get-ResultChanges -Previous (New-Result) -Current (New-Result -Tenant 't2') | Should -Be 'tenant t1 -> t2'
    }
    It 'reports a federation provider change' {
        Get-ResultChanges -Previous (New-Result -Category 'Federated' -Provider 'Okta (a)') -Current (New-Result -Category 'Federated' -Provider 'GoDaddy (b)') |
            Should -Be 'federation provider Okta (a) -> GoDaddy (b)'
    }
    It 'reports federation turning on for the same account type' {
        Get-ResultChanges -Previous (New-Result -Category 'Dual') -Current (New-Result -Category 'Dual' -Provider 'Okta (a)') | Should -Be 'now federated'
    }
    It 'ignores throttled or failed checks' {
        Get-ResultChanges -Previous (New-Result) -Current (New-Result -Category 'Error' -Status 'Throttled') | Should -Be ''
        Get-ResultChanges -Previous (New-Result -Status 'Unavailable') -Current (New-Result -Category 'MSA') | Should -Be ''
    }
}

Describe 'Domain summary' {
    It 'groups results by domain with tenant, cloud and account-type mix' {
        $items = @(
            [pscustomobject]@{ Username = 'a@contoso.com'; Category = 'Entra'; Organization = 'Contoso'; TenantId = 't1'; FoundInCloud = 'Commercial'; FederationProvider = ''; Status = 'OK'; Changes = '' }
            [pscustomobject]@{ Username = 'b@contoso.com'; Category = 'Entra'; Organization = 'Contoso'; TenantId = 't1'; FoundInCloud = 'Commercial'; FederationProvider = ''; Status = 'Throttled'; Changes = 'now federated' }
            [pscustomobject]@{ Username = 'c@CONTOSO.com'; Category = 'NotMicrosoft'; Organization = 'Contoso'; TenantId = 't1'; FoundInCloud = 'N/A'; FederationProvider = ''; Status = 'Inconclusive'; Changes = '' }
            [pscustomobject]@{ Username = 'd@fabrikam.com'; Category = 'Federated'; Organization = 'Fabrikam'; TenantId = 't2'; FoundInCloud = 'Commercial'; FederationProvider = 'Okta (f)'; Status = 'OK'; Changes = '' }
        )
        $summary = @(Get-DomainSummary -Items $items)
        $summary.Count | Should -Be 2
        $summary[0].Domain | Should -Be 'contoso.com'
        $summary[0].Addresses | Should -Be 3
        $summary[0].TenantId | Should -Be 't1'
        $summary[0].Cloud | Should -Be 'Commercial'
        $summary[0].AccountTypes | Should -Be 'Work/School (Entra ID) 2, Not a Microsoft account 1'
        $summary[0].NotFound | Should -Be 1
        $summary[0].NeedsAttention | Should -Be 2
        $summary[0].Changed | Should -Be 1
        $summary[1].FederationProvider | Should -Be 'Okta (f)'
    }
}

Describe 'CSV import and export with extra columns' {
    BeforeAll {
        $script:CsvPath = Join-Path $TestDrive 'people.csv'
        @('Email,EmployeeId,Display Name,Status', 'ada@contoso.com,1001,Ada Lovelace,Active', 'grace@fabrikam.com,1002,"Hopper, Grace",Leaver') | Set-Content -LiteralPath $script:CsvPath -Encoding UTF8
        $script:ByContent = Join-Path $TestDrive 'contacts.csv'
        @('Id,Contact,Team', '1,ada@contoso.com,Eng', '2,grace@fabrikam.com,Research') | Set-Content -LiteralPath $script:ByContent -Encoding UTF8
        $script:TextPath = Join-Path $TestDrive 'list.txt'
        'Please check ada@contoso.com and grace@fabrikam.com today.' | Set-Content -LiteralPath $script:TextPath -Encoding UTF8
    }
    It 'finds the email column by name and keeps the other columns' {
        $import = Import-EmailFile -Path $script:CsvPath
        $import.EmailColumn | Should -Be 'Email'
        @($import.Emails) | Should -Be @('ada@contoso.com', 'grace@fabrikam.com')
        @($import.KeepColumns) | Should -Be @('EmployeeId', 'Display Name', 'Status')
        $import.Extras['grace@fabrikam.com']['Display Name'] | Should -Be 'Hopper, Grace'
    }
    It 'finds the email column by content when no column name matches' {
        (Import-EmailFile -Path $script:ByContent).EmailColumn | Should -Be 'Contact'
    }
    It 'keeps only the chosen columns' {
        @((Import-EmailFile -Path $script:CsvPath -KeepColumns @('EmployeeId')).KeepColumns) | Should -Be @('EmployeeId')
    }
    It 'reads addresses out of plain text' {
        @((Import-EmailFile -Path $script:TextPath).Emails) | Should -Be @('ada@contoso.com', 'grace@fabrikam.com')
    }
    It 'adds extra columns to exports and renames clashes' {
        $row = [pscustomobject]@{ Username = 'ada@contoso.com'; Status = 'OK'; Extra = @{ c0 = '1001'; c1 = 'Active' } }
        $export = @(Get-ExportRows -Items @($row) -ExtraColumns @('EmployeeId', 'Status'))[0]
        $export.EmployeeId | Should -Be '1001'
        $export.'Input Status' | Should -Be 'Active'
        $export.Status | Should -Be 'OK'
    }
    It 'redacts addresses and extra columns when asked' {
        $row = [pscustomobject]@{ Username = 'ada@contoso.com'; Status = 'OK'; Extra = @{ c0 = 'Ada Lovelace' } }
        $export = @(Get-ExportRows -Items @($row) -ExtraColumns @('Display Name') -Redact)[0]
        $export.Username | Should -Be 'a***@contoso.com'
        $export.'Display Name' | Should -Be '[redacted]'
    }
    It 'redacts addresses in the HTML report' {
        $row = [pscustomobject]@{ Timestamp = '2026-11-19 10:00:00'; Username = 'secret.person@contoso.com'; Category = 'Entra'; Status = 'OK'; AccountType = 'x' }
        $html = New-HtmlReport -Items @($row) -Title 'Test' -Redact
        $html | Should -Not -Match 'secret.person'
        $html | Should -Match 's\*\*\*@contoso.com'
    }
}

Describe 'Endpoint health advice' {
    It 'suggests <Concurrency> at a time, <Delay> ms apart, when <Case>' -TestCases @(
        @{ Case = 'there is too little data'; Calls = 2; Rate = 0; Availability = 100; P95 = 300; Concurrency = 1; Delay = 200 }
        @{ Case = 'throttling is heavy'; Calls = 50; Rate = 15; Availability = 100; P95 = 300; Concurrency = 1; Delay = 1000 }
        @{ Case = 'availability is poor'; Calls = 50; Rate = 0; Availability = 80; P95 = 300; Concurrency = 1; Delay = 1000 }
        @{ Case = 'there is some throttling'; Calls = 50; Rate = 3; Availability = 100; P95 = 300; Concurrency = 1; Delay = 500 }
        @{ Case = 'responses are fast and unthrottled'; Calls = 50; Rate = 0; Availability = 100; P95 = 500; Concurrency = 3; Delay = 200 }
        @{ Case = 'responses are slow'; Calls = 50; Rate = 0; Availability = 100; P95 = 1500; Concurrency = 2; Delay = 300 }
    ) {
        param($Case, $Calls, $Rate, $Availability, $P95, $Concurrency, $Delay)
        $advice = Get-SafeBatchAdvice -Calls $Calls -ThrottleRate $Rate -Availability $Availability -P95 $P95
        $advice.Concurrency | Should -Be $Concurrency
        $advice.DelayMs | Should -Be $Delay
    }
    It 'computes recent latency, throttling rate and a per-minute timeline' {
        $CoreStats = (New-SharedState).Stats
        Add-CoreTiming -Endpoint 'GetCredentialType (Commercial)' -Milliseconds 100 -Failed $false
        Add-CoreTiming -Endpoint 'GetCredentialType (Commercial)' -Milliseconds 300 -Failed $false
        Add-CoreTiming -Endpoint 'GetUserRealm' -Milliseconds 50 -Failed $true
        Add-CoreThrottle
        $recent = (Get-CoreStatsSnapshot -Stats $CoreStats).Recent
        $recent.LookupCalls | Should -Be 2
        $recent.Calls | Should -Be 3
        $recent.MedianMs | Should -Be 100
        $recent.ThrottleRatePct | Should -Be 50
        [Math]::Round($recent.AvailabilityPct, 1) | Should -Be 66.7
        $timeline = (Get-CoreStatsSnapshot -Stats $CoreStats).Timeline
        $timeline.Count | Should -Be 30
        $timeline[29].AverageMs | Should -Be 200
        $timeline[29].Calls | Should -Be 2
    }
}

Describe 'Command-line mode' {
    BeforeAll {
        $script:HostExe = (Get-Process -Id $PID).Path
    }
    It 'returns exit code 3 for a missing input file' {
        $null = & $script:HostExe -NoProfile -NonInteractive -File $script:ScriptPath -InputFile (Join-Path $TestDrive 'missing.csv') 2>&1
        $LASTEXITCODE | Should -Be 3
    }
    It 'returns exit code 1 and JSON for an invalid address, without contacting Microsoft' {
        $output = & $script:HostExe -NoProfile -NonInteractive -File $script:ScriptPath -Email 'not-an-address' -Format Json
        $LASTEXITCODE | Should -Be 1
        @((($output | Out-String) | ConvertFrom-Json))[0].status | Should -Be 'Invalid'
    }
    It 'maps statuses to exit codes' {
        Get-CliExitCode -Results @([pscustomobject]@{ status = 'OK' }) | Should -Be 0
        Get-CliExitCode -Results @([pscustomobject]@{ status = 'OK' }, [pscustomobject]@{ status = 'Inconclusive' }) | Should -Be 1
        Get-CliExitCode -Results @([pscustomobject]@{ status = 'Throttled' }, [pscustomobject]@{ status = 'Unavailable' }) | Should -Be 2
    }
}

Describe 'PowerShell module' {
    BeforeAll {
        Import-Module (Join-Path $PSScriptRoot '..\module\MSAccountChecker.psd1') -Force
    }
    AfterAll {
        Remove-Module MSAccountChecker -ErrorAction SilentlyContinue
    }
    It 'exports its commands' {
        $names = (Get-Command -Module MSAccountChecker).Name
        $names | Should -Contain 'Test-MSAccount'
        $names | Should -Contain 'Clear-MSAccountCache'
        $names | Should -Contain 'Get-MSAccountExitCode'
    }
    It 'checks addresses from the pipeline' {
        Mock -ModuleName MSAccountChecker Invoke-RestMethod {
            if ($Uri -like '*GetCredentialType*') { return [pscustomobject]@{ IfExistsResult = 0; ThrottleStatus = 0 } }
            if ($Uri -like '*getuserrealm*') { return [pscustomobject]@{ NameSpaceType = 'Managed'; FederationBrandName = 'Contoso' } }
            return [pscustomobject]@{ token_endpoint = 'https://login.microsoftonline.com/11111111-2222-3333-4444-555555555555/oauth2/token'; tenant_region_scope = 'EU' }
        }
        Mock -ModuleName MSAccountChecker Start-Sleep { }
        $results = @([pscustomobject]@{ UserPrincipalName = 'a@contoso.com' }, [pscustomobject]@{ UserPrincipalName = 'b@contoso.com' }) | Test-MSAccount -NoCache
        @($results).Count | Should -Be 2
        $results[0].category | Should -Be 'Entra'
        $results[0].organization | Should -Be 'Contoso'
        $results[0].tenantRegion | Should -Be 'Europe (EU)'
        Get-MSAccountExitCode -Results $results | Should -Be 0
    }
}

Describe 'Encryption at rest' -Skip:(-not ($IsWindows -or $PSVersionTable.PSEdition -eq 'Desktop')) {
    BeforeAll {
        if (-not ('MsAccCheckerNative3' -as [type])) {
            $assignment = $script:Ast.Find({
                param($node)
                $node -is [System.Management.Automation.Language.AssignmentStatementAst] -and $node.Left.Extent.Text -eq '$nativeSource'
            }, $true)
            Add-Type -TypeDefinition $assignment.Right.Expression.Value
        }
    }
    It 'round-trips data through Windows data protection' {
        $plain = '{"Username":"user@contoso.com"}'
        $protected = [MsAccCheckerNative3]::Protect([System.Text.Encoding]::UTF8.GetBytes($plain))
        [System.Text.Encoding]::UTF8.GetString($protected) | Should -Not -Match 'contoso'
        [System.Text.Encoding]::UTF8.GetString([MsAccCheckerNative3]::Unprotect($protected)) | Should -Be $plain
    }
}
