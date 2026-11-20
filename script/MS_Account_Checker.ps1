#requires -Version 5.1
<#
.SYNOPSIS
    Microsoft Account Checker - identify the Microsoft account type behind an email address.

.DESCRIPTION
    Queries Microsoft's public sign-in discovery endpoints to find out whether an email
    address belongs to a Microsoft account and, if so, which kind and in which cloud
    (Commercial, GCC High or China).

    Lookup order:
      1. GetCredentialType in the Commercial cloud.
      2. GCC High, then China, only when Commercial returns IfExistsResult 1 for a
         non-personal domain. A sovereign result of 0 or 5 is trusted.

    New in 2.0.0
      - WPF interface with live light/dark theme, batch checks, history and statistics
      - Federation confirmed with GetUserRealm and the credential response, not one field
      - Throttling detection (ThrottleStatus, HTTP 429/503) with automatic retry and back-off
      - Tenant details: organization name, tenant ID, tenant region, federation provider
      - 24-hour lookup cache; CSV, JSON and HTML export
      - Optional local REST API on http://localhost:<port>/ with a web page for checking in a browser
      - System tray mode: keeps the API running when the window is closed (-ApiMode starts in the tray)

.PARAMETER ApiMode
    Starts the API server immediately and runs from the system tray without showing the window.
    Right-click the tray icon to open the web page, show the window or exit.

.PARAMETER Port
    Port for the API server. Overrides the port saved in the settings.

.PARAMETER NoBrowser
    Does not open the web page in the browser when the API server starts.

.PARAMETER Email
    Command-line mode: checks these addresses and prints the results instead of opening the window.
    Exit code: 0 all definitive, 1 some inconclusive, throttled or invalid, 2 some unavailable, 3 bad input.

.PARAMETER InputFile
    Command-line mode: checks the addresses in a .csv or .txt file. Other CSV columns are kept in the output.

.PARAMETER OutFile
    Command-line mode: writes the results to this file (CSV or JSON, see -Format).

.PARAMETER Format
    Command-line mode output: Table (default), Object (for the pipeline), Json or Csv.

.PARAMETER NoCache
    Command-line mode: do not reuse answers within the run.

.PARAMETER Console
    Keeps the app attached to the console it was started from, so errors are visible there.
    By default the app restarts itself without a console window.

.EXAMPLE
    .\MS_Account_Checker.ps1
    Opens the desktop window.

.EXAMPLE
    .\MS_Account_Checker.ps1 -Email user@contoso.com, other@fabrikam.com
    Command-line mode: prints a table and sets the exit code.

.EXAMPLE
    .\MS_Account_Checker.ps1 -InputFile .\users.csv -OutFile .\results.csv -Format Csv
    Command-line mode: checks a CSV file and saves the results with the original columns.

.EXAMPLE
    powershell.exe -STA -WindowStyle Hidden -File .\MS_Account_Checker.ps1 -ApiMode -Port 8080
    Runs the API server from the system tray and opens http://localhost:8080/ in the browser.

.NOTES
    Name       : MS_Account_Checker.ps1
    Version    : 2.0.0
    Author     : imumeshk
    Repository : https://github.com/imumeshk/MS-Account-Checker
    License    : CC0-1.0
    Requires   : Windows; Windows PowerShell 5.1 or PowerShell 7+; STA thread
    Settings   : .\config\AccChecker.config (Key=Value)
    Logs       : .\logs\MS_Account_Checker_Logs_YYYY-MM-DD.txt (tab-separated)

.LINK
    https://github.com/imumeshk/MS-Account-Checker
#>
[CmdletBinding()]
param(
    [switch]$ApiMode,
    [ValidateRange(0, 65535)][int]$Port = 0,
    [switch]$NoBrowser,
    [switch]$Console,
    [string[]]$Email,
    [string]$InputFile,
    [string]$OutFile,
    [ValidateSet('Table', 'Object', 'Json', 'Csv')][string]$Format = 'Table',
    [switch]$NoCache
)
$script:CliMode = [bool]($Email -or $InputFile)

#region Native helpers and running without a console window
if (-not ('MsAccCheckerNative3' -as [type])) {
    $nativeSource = @'
using System;
using System.Runtime.InteropServices;
public static class MsAccCheckerNative3
{
    [DllImport("dwmapi.dll")]
    public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attribute, ref int value, int size);
    [DllImport("kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool IsWindowVisible(IntPtr hwnd);
    [DllImport("user32.dll")]
    public static extern uint GetClipboardSequenceNumber();

    // Windows data protection (DPAPI), current user scope: only this Windows account can read the data.
    [StructLayout(LayoutKind.Sequential)]
    private struct DATA_BLOB { public int cbData; public IntPtr pbData; }
    [DllImport("crypt32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern bool CryptProtectData(ref DATA_BLOB pDataIn, string szDataDescr, IntPtr pOptionalEntropy, IntPtr pvReserved, IntPtr pPromptStruct, int dwFlags, ref DATA_BLOB pDataOut);
    [DllImport("crypt32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern bool CryptUnprotectData(ref DATA_BLOB pDataIn, IntPtr ppszDataDescr, IntPtr pOptionalEntropy, IntPtr pvReserved, IntPtr pPromptStruct, int dwFlags, ref DATA_BLOB pDataOut);
    [DllImport("kernel32.dll")]
    private static extern IntPtr LocalFree(IntPtr hMem);

    public static byte[] Protect(byte[] data) { return Transform(data, true); }
    public static byte[] Unprotect(byte[] data) { return Transform(data, false); }

    private static byte[] Transform(byte[] data, bool protect)
    {
        DATA_BLOB input = new DATA_BLOB();
        DATA_BLOB output = new DATA_BLOB();
        input.cbData = data.Length;
        input.pbData = Marshal.AllocHGlobal(Math.Max(1, data.Length));
        try
        {
            Marshal.Copy(data, 0, input.pbData, data.Length);
            const int UiForbidden = 0x1;
            bool ok = protect
                ? CryptProtectData(ref input, "MS Account Checker", IntPtr.Zero, IntPtr.Zero, IntPtr.Zero, UiForbidden, ref output)
                : CryptUnprotectData(ref input, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero, UiForbidden, ref output);
            if (!ok) { throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error()); }
            byte[] result = new byte[output.cbData];
            Marshal.Copy(output.pbData, result, 0, output.cbData);
            return result;
        }
        finally
        {
            Marshal.FreeHGlobal(input.pbData);
            if (output.pbData != IntPtr.Zero) { LocalFree(output.pbData); }
        }
    }
}
'@
    try { Add-Type -TypeDefinition $nativeSource -ErrorAction Stop } catch { }
}

# When started from a visible console (Run with PowerShell, Windows Terminal, a shortcut without
# -WindowStyle Hidden), restart in a background process that has no console, so only the app window
# appears and closing the console cannot close the app. launcher.vbs already starts the script hidden,
# so in that case it simply continues here. Use -Console to stay in the console for troubleshooting.
if (-not $script:CliMode -and -not $Console -and $env:MSAC_NO_CONSOLE -ne '1' -and $PSCommandPath) {
    $consoleShown = $false
    try {
        $consoleHandle = [MsAccCheckerNative3]::GetConsoleWindow()
        $consoleShown = ($consoleHandle -ne [IntPtr]::Zero) -and [MsAccCheckerNative3]::IsWindowVisible($consoleHandle)
    } catch { }
    # Windows Terminal and the VS Code terminal host the console in a hidden pseudo-window.
    if ($env:WT_SESSION -or $env:TERM_PROGRAM) { $consoleShown = $true }

    if ($consoleShown) {
        try {
            $arguments = @('-NoProfile', '-NoLogo', '-STA', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $PSCommandPath))
            if ($ApiMode)      { $arguments += '-ApiMode' }
            if ($Port -gt 0)   { $arguments += @('-Port', [string]$Port) }
            if ($NoBrowser)    { $arguments += '-NoBrowser' }
            $startInfo = New-Object System.Diagnostics.ProcessStartInfo
            $startInfo.FileName         = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
            $startInfo.Arguments        = $arguments -join ' '
            $startInfo.UseShellExecute  = $false
            $startInfo.CreateNoWindow   = $true
            $startInfo.WorkingDirectory = Split-Path -Parent $PSCommandPath
            $startInfo.EnvironmentVariables['MSAC_NO_CONSOLE'] = '1'
            [void][System.Diagnostics.Process]::Start($startInfo)
            return
        } catch {
            # If the restart fails for any reason, keep running in this console.
        }
    }
}
Remove-Item -Path Env:\MSAC_NO_CONSOLE -ErrorAction SilentlyContinue
#endregion

#region Startup checks, assemblies, paths
if (-not $script:CliMode) { Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase }

if (-not $script:CliMode -and [System.Threading.Thread]::CurrentThread.GetApartmentState() -ne [System.Threading.ApartmentState]::STA) {
    $staMessage = 'Microsoft Account Checker needs an STA thread. Start it with launcher.vbs, or run: powershell.exe -STA -File .\MS_Account_Checker.ps1'
    Write-Error $staMessage
    try { [void][System.Windows.MessageBox]::Show($staMessage, 'Microsoft Account Checker') } catch { }
    return
}

try {
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12
} catch { }


$script:AppName      = 'Microsoft Account Checker'
$script:AppVersion   = '2.0.0'
$script:AppAuthor    = 'imumeshk'
$script:AppRepo      = 'https://github.com/imumeshk/MS-Account-Checker'
$script:AppIssues    = 'https://github.com/imumeshk/MS-Account-Checker/issues'
$script:AppLicense   = 'CC0-1.0'
$script:LogHeader    = "Timestamp`tUsername`tFoundInCloud`tCloudURL`tFederationRedirectUrl`tAccountType`tOrganization`tTenantId`tTenantRegion`tFederationProvider`tStatus"
$script:AppDir       = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).ProviderPath }
$script:ScriptFile   = $PSCommandPath
$script:ConfigDir    = Join-Path $script:AppDir 'config'
$script:LogsDir      = Join-Path $script:AppDir 'logs'
$script:ConfigFile   = Join-Path $script:ConfigDir 'AccChecker.config'
$script:HistoryFile  = Join-Path $script:ConfigDir 'check_history.json'
$script:CacheFile    = Join-Path $script:ConfigDir 'lookup_cache.json'
$script:CacheHours   = 24
$script:HistoryLimit = 5000
$script:ApiLogLimit  = 500
$script:BatchDelayMs = 200

foreach ($dir in @($script:ConfigDir, $script:LogsDir)) {
    if (-not (Test-Path -LiteralPath $dir)) { [void](New-Item -ItemType Directory -Path $dir -Force) }
}
#endregion

#region Core lookup logic (shared by the window, the batch workers and the API server)
# This block is dot-sourced into the main script, into every background runspace and into the tests,
# so the exact same rules are used everywhere. $CoreStats, when set by the host, collects diagnostics.
$script:CoreLogic = {
    $CorePersonalDomains = @('outlook.com', 'hotmail.com', 'live.com', 'msn.com')
    $CoreEmailRegex      = [regex]'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    $CoreGuidRegex       = [regex]'[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}'
    $CoreMaxAttempts     = 3
    $CoreCloudBase       = @{
        Commercial = 'https://login.microsoftonline.com'
        GCCHigh    = 'https://login.microsoftonline.us'
        China      = 'https://login.partner.microsoftonline.cn'
    }
    $CoreRegionNames = @{
        NA = 'North America'; EU = 'Europe'; AS = 'Asia Pacific'; OC = 'Oceania'
        AF = 'Africa'; SA = 'South America'; WW = 'Worldwide'
    }
    $CoreFederationProviders = [ordered]@{
        'godaddy'      = 'GoDaddy'
        'okta'         = 'Okta'
        'onelogin'     = 'OneLogin'
        'pingone'      = 'PingOne'
        'pingidentity' = 'Ping Identity'
        'duosecurity'  = 'Duo'
        'auth0'        = 'Auth0'
        'jumpcloud'    = 'JumpCloud'
        'secureauth'   = 'SecureAuth'
        'centrify'     = 'Centrify'
        'idaptive'     = 'CyberArk Identity'
        'cyberark'     = 'CyberArk'
        'google'       = 'Google'
        'shibboleth'   = 'Shibboleth'
        'citrix'       = 'Citrix'
    }
    if (-not (Test-Path variable:CoreStats)) { $CoreStats = $null }

    # ---- Diagnostics -------------------------------------------------------------------------
    function Add-CoreCounter {
        param([string]$Name, [int]$By = 1)
        $stats = $CoreStats
        if ($null -eq $stats) { return }
        [System.Threading.Monitor]::Enter($stats.SyncRoot)
        try { $stats[$Name] = [int]$stats[$Name] + $By } finally { [System.Threading.Monitor]::Exit($stats.SyncRoot) }
    }

    function Set-CoreValue {
        param([string]$Name, $Value)
        $stats = $CoreStats
        if ($null -eq $stats) { return }
        [System.Threading.Monitor]::Enter($stats.SyncRoot)
        try { $stats[$Name] = $Value } finally { [System.Threading.Monitor]::Exit($stats.SyncRoot) }
    }

    function Add-CoreTiming {
        param([string]$Endpoint, [double]$Milliseconds, [bool]$Failed)
        $stats = $CoreStats
        if ($null -eq $stats -or -not $Endpoint) { return }
        [System.Threading.Monitor]::Enter($stats.SyncRoot)
        try {
            $table = $stats['Endpoints']
            $entry = $table[$Endpoint]
            if ($null -eq $entry) {
                $entry = @{ Calls = 0; Failures = 0; TotalMs = 0.0; LastMs = 0.0 }
                $table[$Endpoint] = $entry
            }
            $entry['Calls']   = $entry['Calls'] + 1
            if ($Failed) { $entry['Failures'] = $entry['Failures'] + 1 }
            $entry['TotalMs'] = $entry['TotalMs'] + $Milliseconds
            $entry['LastMs']  = $Milliseconds
            # Rolling samples for the recent-health view (latency over time, availability, throttling rate).
            $samples = $stats['Samples']
            if ($null -ne $samples) {
                $samples.Add(@{ T = [datetime]::UtcNow; E = $Endpoint; Ms = $Milliseconds; F = $Failed })
                if ($samples.Count -gt 3000) { $samples.RemoveRange(0, $samples.Count - 3000) }
            }
        } finally {
            [System.Threading.Monitor]::Exit($stats.SyncRoot)
        }
    }

    function Add-CoreThrottle {
        Add-CoreCounter 'ThrottleEvents'
        $stats = $CoreStats
        if ($null -eq $stats) { return }
        [System.Threading.Monitor]::Enter($stats.SyncRoot)
        try {
            $marks = $stats['ThrottleTimes']
            if ($null -ne $marks) {
                $marks.Add([datetime]::UtcNow)
                if ($marks.Count -gt 2000) { $marks.RemoveRange(0, $marks.Count - 2000) }
            }
        } finally {
            [System.Threading.Monitor]::Exit($stats.SyncRoot)
        }
    }

    # Suggested safe batch settings from recent throttling, availability and latency.
    function Get-SafeBatchAdvice {
        param([int]$Calls, [double]$ThrottleRate, $Availability, [double]$P95)
        if ($Calls -lt 5) {
            return @{ Concurrency = 1; DelayMs = 200; Reason = 'Not enough recent lookups to judge. One at a time, 200 ms apart, is the safe default.' }
        }
        if ($ThrottleRate -ge 10 -or ($null -ne $Availability -and $Availability -lt 90)) {
            return @{ Concurrency = 1; DelayMs = 1000; Reason = 'Microsoft is throttling or failing often. Slow down to one lookup a second.' }
        }
        if ($ThrottleRate -ge 2) {
            return @{ Concurrency = 1; DelayMs = 500; Reason = 'Some throttling in the last 5 minutes. One at a time, 500 ms apart, avoids most of it.' }
        }
        if ($P95 -gt 0 -and $P95 -lt 800) {
            return @{ Concurrency = 3; DelayMs = 200; Reason = 'No throttling and fast responses. Up to 3 parallel lookups should be safe.' }
        }
        return @{ Concurrency = 2; DelayMs = 300; Reason = 'No throttling, but responses are slow. 2 parallel lookups, 300 ms apart.' }
    }

    function Get-CoreStatsSnapshot {
        param($Stats)
        if ($null -eq $Stats) { return $null }
        [System.Threading.Monitor]::Enter($Stats.SyncRoot)
        try {
            $endpoints = New-Object System.Collections.Generic.List[object]
            foreach ($name in @($Stats['Endpoints'].Keys | Sort-Object)) {
                $e = $Stats['Endpoints'][$name]
                $calls = [int]$e['Calls']
                $endpoints.Add([pscustomobject][ordered]@{
                    Endpoint     = $name
                    Calls        = $calls
                    Failures     = [int]$e['Failures']
                    AverageMs    = if ($calls) { [int]($e['TotalMs'] / $calls) } else { 0 }
                    LastMs       = [int]$e['LastMs']
                    Availability = if ($calls) { '{0:0.0}%' -f (100.0 * ($calls - [int]$e['Failures']) / $calls) } else { '' }
                })
            }

            # Last 5 minutes, and a per-minute timeline of the last 30 minutes.
            $now = [datetime]::UtcNow
            $windowStart = $now.AddMinutes(-5)
            $recentCalls = 0; $recentFailures = 0; $credentialCalls = 0
            $latencies = New-Object System.Collections.Generic.List[double]
            $buckets = @(); for ($b = 0; $b -lt 30; $b++) { $buckets += ,@{ Sum = 0.0; N = 0; Failures = 0 } }
            $samples = $Stats['Samples']
            if ($null -ne $samples) {
                foreach ($sample in $samples) {
                    $minutesAgo = ($now - $sample.T).TotalMinutes
                    if ($minutesAgo -lt 30 -and $sample.E -like 'GetCredentialType*') {
                        $bucket = $buckets[29 - [int][Math]::Floor($minutesAgo)]
                        if ($sample.F) { $bucket['Failures'] = $bucket['Failures'] + 1 } else { $bucket['Sum'] = $bucket['Sum'] + $sample.Ms; $bucket['N'] = $bucket['N'] + 1 }
                    }
                    if ($sample.T -lt $windowStart) { continue }
                    $recentCalls++
                    if ($sample.F) { $recentFailures++ }
                    if ($sample.E -like 'GetCredentialType*') {
                        $credentialCalls++
                        if (-not $sample.F) { $latencies.Add([double]$sample.Ms) }
                    }
                }
            }
            $recentThrottles = 0
            if ($null -ne $Stats['ThrottleTimes']) { foreach ($mark in $Stats['ThrottleTimes']) { if ($mark -ge $windowStart) { $recentThrottles++ } } }
            $sorted = @($latencies | Sort-Object)
            $p50 = if ($sorted.Count) { $sorted[[int][Math]::Floor(($sorted.Count - 1) * 0.5)] } else { 0 }
            $p95 = if ($sorted.Count) { $sorted[[int][Math]::Floor(($sorted.Count - 1) * 0.95)] } else { 0 }
            $availability = if ($recentCalls) { 100.0 * ($recentCalls - $recentFailures) / $recentCalls } else { $null }
            $throttleRate = if ($credentialCalls) { 100.0 * $recentThrottles / $credentialCalls } else { 0.0 }
            $advice = Get-SafeBatchAdvice -Calls $credentialCalls -ThrottleRate $throttleRate -Availability $availability -P95 $p95
            $timeline = New-Object System.Collections.Generic.List[object]
            for ($b = 0; $b -lt 30; $b++) {
                $bucket = $buckets[$b]
                $timeline.Add([pscustomobject]@{
                    MinutesAgo = 29 - $b
                    AverageMs  = if ($bucket['N']) { [int]($bucket['Sum'] / $bucket['N']) } else { $null }
                    Calls      = $bucket['N'] + $bucket['Failures']
                    Failures   = $bucket['Failures']
                })
            }

            $live = [int]$Stats['LiveLookups']
            $hits = [int]$Stats['CacheHits']
            return [pscustomobject][ordered]@{
                LiveLookups    = $live
                CacheHits      = $hits
                CacheHitRate   = if (($live + $hits) -gt 0) { '{0:0}%' -f (100.0 * $hits / ($live + $hits)) } else { '' }
                Retries        = [int]$Stats['Retries']
                ThrottleEvents = [int]$Stats['ThrottleEvents']
                Unavailable    = [int]$Stats['Unavailable']
                Inconclusive   = [int]$Stats['Inconclusive']
                LastSuccess    = [string]$Stats['LastSuccess']
                LastError      = [string]$Stats['LastError']
                Since          = [string]$Stats['Since']
                Endpoints      = $endpoints.ToArray()
                Recent         = [pscustomobject][ordered]@{
                    Calls             = $recentCalls
                    LookupCalls       = $credentialCalls
                    AvailabilityPct   = $availability
                    MedianMs          = [int]$p50
                    P95Ms             = [int]$p95
                    ThrottleRatePct   = $throttleRate
                    Throttles         = $recentThrottles
                    SuggestedParallel = $advice.Concurrency
                    SuggestedDelayMs  = $advice.DelayMs
                    Suggestion        = $advice.Reason
                }
                Timeline       = $timeline.ToArray()
            }
        } finally {
            [System.Threading.Monitor]::Exit($Stats.SyncRoot)
        }
    }

    function Reset-CoreStats {
        param($Stats)
        if ($null -eq $Stats) { return }
        [System.Threading.Monitor]::Enter($Stats.SyncRoot)
        try {
            foreach ($name in @('LiveLookups', 'CacheHits', 'Retries', 'ThrottleEvents', 'Unavailable', 'Inconclusive')) { $Stats[$name] = 0 }
            $Stats['LastSuccess'] = ''
            $Stats['LastError']   = ''
            $Stats['Since']       = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            $Stats['Endpoints']   = @{}
            if ($null -ne $Stats['Samples']) { $Stats['Samples'].Clear() }
            if ($null -ne $Stats['ThrottleTimes']) { $Stats['ThrottleTimes'].Clear() }
        } finally {
            [System.Threading.Monitor]::Exit($Stats.SyncRoot)
        }
    }

    # ---- Validation and HTTP -------------------------------------------------------------------
    function Test-EmailFormat {
        param([string]$Email)
        if ([string]::IsNullOrWhiteSpace($Email)) { return $false }
        return $CoreEmailRegex.IsMatch($Email.Trim())
    }

    function Get-BackoffMs {
        param([int]$Attempt)
        return [int](2000 * [Math]::Pow(2, $Attempt - 1)) + (Get-Random -Minimum 0 -Maximum 500)
    }

    function Get-HttpStatus {
        param($ErrorRecord)
        try {
            $response = $ErrorRecord.Exception.Response
            if ($null -ne $response) { return [int]$response.StatusCode }
        } catch { }
        return 0
    }

    # One HTTP call with retry and back-off when Microsoft answers 429 (too many requests) or 503.
    function Invoke-CoreRequest {
        param([string]$Uri, [string]$Method = 'Get', [string]$Body = '', [string]$Endpoint = '')
        for ($attempt = 1; ; $attempt++) {
            $watch = [System.Diagnostics.Stopwatch]::StartNew()
            try {
                if ($Body) {
                    $reply = Invoke-RestMethod -Uri $Uri -Method $Method -Body $Body -ContentType 'application/json' -TimeoutSec 15 -ErrorAction Stop
                } else {
                    $reply = Invoke-RestMethod -Uri $Uri -Method $Method -TimeoutSec 15 -ErrorAction Stop
                }
                Add-CoreTiming -Endpoint $Endpoint -Milliseconds $watch.Elapsed.TotalMilliseconds -Failed $false
                return $reply
            } catch {
                $status = Get-HttpStatus $_
                # 400/404 from tenant discovery mean "no tenant for this domain": the endpoint itself worked.
                Add-CoreTiming -Endpoint $Endpoint -Milliseconds $watch.Elapsed.TotalMilliseconds -Failed (-not ($status -eq 400 -or $status -eq 404))
                if ($status -eq 429) { Add-CoreThrottle }
                if (($status -eq 429 -or $status -eq 503) -and $attempt -lt $CoreMaxAttempts) {
                    Add-CoreCounter 'Retries'
                    Start-Sleep -Milliseconds (Get-BackoffMs $attempt)
                    continue
                }
                throw
            }
        }
    }

    # GetCredentialType with throttling detection. A response with ThrottleStatus other than 0 is
    # not trustworthy, so it is retried; if every attempt is throttled the result is flagged.
    function Invoke-CredentialTypeLookup {
        param([string]$Email, [string]$Cloud)
        $url  = '{0}/common/GetCredentialType' -f $CoreCloudBase[$Cloud]
        $body = @{ Username = $Email; isOtherIdpSupported = $true } | ConvertTo-Json
        $response = $null
        for ($attempt = 1; $attempt -le $CoreMaxAttempts; $attempt++) {
            try {
                $response = Invoke-CoreRequest -Uri $url -Method 'Post' -Body $body -Endpoint "GetCredentialType ($Cloud)"
            } catch {
                throw "Network/API Error: $($_.Exception.Message)"
            }
            if ($null -eq $response) {
                throw "API Response Error: Received null response from Microsoft API for $Cloud cloud."
            }
            if ($null -eq $response.PSObject.Properties['IfExistsResult']) {
                throw "API Response Error: 'IfExistsResult' property missing from Microsoft API response for $Cloud cloud."
            }
            $throttled = $false
            if ($null -ne $response.PSObject.Properties['ThrottleStatus']) {
                $throttleValue = 0
                [void][int]::TryParse([string]$response.ThrottleStatus, [ref]$throttleValue)
                $throttled = ($throttleValue -ne 0)
            }
            if (-not $throttled) {
                return [pscustomobject]@{ Response = $response; Throttled = $false; Attempts = $attempt }
            }
            Add-CoreThrottle
            if ($attempt -lt $CoreMaxAttempts) {
                Add-CoreCounter 'Retries'
                Start-Sleep -Milliseconds (Get-BackoffMs $attempt)
            }
        }
        return [pscustomobject]@{ Response = $response; Throttled = $true; Attempts = $CoreMaxAttempts }
    }

    function Get-FederationProvider {
        param([string]$Url)
        if ([string]::IsNullOrWhiteSpace($Url)) { return '' }
        try { $uri = [Uri]$Url } catch { return '' }
        $hostName = $uri.Host.ToLowerInvariant()
        $path     = $uri.AbsolutePath.ToLowerInvariant()
        foreach ($key in $CoreFederationProviders.Keys) {
            if ($hostName -like "*$key*") { return "$($CoreFederationProviders[$key]) ($hostName)" }
        }
        if ($path -like '*/adfs/*') { return "AD FS ($hostName)" }
        if ($hostName -like 'adfs.*' -or $hostName -like 'sts.*' -or $hostName -like 'fs.*') { return "AD FS, likely ($hostName)" }
        return $hostName
    }

    function Format-TenantRegion {
        param([string]$Scope, [string]$SubScope)
        if (-not $Scope) { return '' }
        $name = $CoreRegionNames[$Scope]
        $text = if ($name) { "$name ($Scope)" } else { $Scope }
        if ($SubScope) { $text += ", $SubScope" }
        return $text
    }

    # Label and recommended next step for each status. "OK" means Microsoft gave a definitive answer.
    function Get-StatusAdvice {
        param([string]$Status)
        switch ($Status) {
            'OK'           { return @{ Label = 'Definitive'; Action = '' } }
            'Inconclusive' { return @{ Label = 'Inconclusive'; Action = "Microsoft's answer is not definitive. Confirm the user's primary sign-in address (UPN) with the organization's administrator; this address may be an alias, a guest or a mail contact." } }
            'Throttled'    { return @{ Label = 'Throttled'; Action = 'Microsoft rate-limited every attempt, so the answer may be wrong. Wait a few minutes and check again, or use Retry failed on the Batch check tab.' } }
            'Unavailable'  { return @{ Label = 'Temporarily unavailable'; Action = 'Microsoft could not be reached. Check the network connection, proxy or firewall, then check again.' } }
            'Invalid'      { return @{ Label = 'Invalid address'; Action = 'Correct the email address format and check again.' } }
            default        { return @{ Label = [string]$Status; Action = '' } }
        }
    }

    function Set-StatusFields {
        param($Result)
        $advice = Get-StatusAdvice -Status $Result.Status
        $Result.StatusLabel = $advice.Label
        $Result.NextAction  = $advice.Action
    }

    # Per-domain details: GetUserRealm (Commercial only) and the tenant's OpenID discovery document.
    function Get-DomainDetails {
        param([string]$Email, [string]$Domain, [string]$Cloud, $DomainCache, [double]$TtlHours = 24)
        $key = "$Cloud|$Domain"
        if ($null -ne $DomainCache) {
            $entry = $DomainCache[$key]
            if ($null -ne $entry -and ([datetime]::UtcNow - $entry.Time).TotalHours -lt $TtlHours) { return $entry.Details }
        }

        $details = [pscustomobject][ordered]@{
            NamespaceType = ''; Organization = ''; RealmCloudInstance = ''; RealmAuthUrl = ''
            TenantId = ''; TenantRegion = ''; TenantCloudInstance = ''
            RealmRaw = $null; DiscoveryRaw = $null; Errors = ''; Complete = $true
        }
        $errors = @()

        if ($Cloud -eq 'Commercial') {
            try {
                $realmUrl = 'https://login.microsoftonline.com/getuserrealm.srf?login={0}&json=1' -f [Uri]::EscapeDataString($Email)
                $realm = Invoke-CoreRequest -Uri $realmUrl -Endpoint 'GetUserRealm'
                $details.RealmRaw           = $realm
                $details.NamespaceType      = [string]$realm.NameSpaceType
                $details.Organization       = [string]$realm.FederationBrandName
                $details.RealmCloudInstance = [string]$realm.CloudInstanceName
                $details.RealmAuthUrl       = [string]$realm.AuthURL
            } catch {
                $errors += "GetUserRealm: $($_.Exception.Message)"
                $details.Complete = $false
            }
        }

        try {
            $discoveryUrl = '{0}/{1}/.well-known/openid-configuration' -f $CoreCloudBase[$Cloud], [Uri]::EscapeDataString($Domain)
            $document = Invoke-CoreRequest -Uri $discoveryUrl -Endpoint "Tenant discovery ($Cloud)"
            $details.DiscoveryRaw = $document
            $match = $CoreGuidRegex.Match([string]$document.token_endpoint)
            if (-not $match.Success) { $match = $CoreGuidRegex.Match([string]$document.issuer) }
            if ($match.Success) { $details.TenantId = $match.Value.ToLowerInvariant() }
            $details.TenantRegion        = Format-TenantRegion -Scope ([string]$document.tenant_region_scope) -SubScope ([string]$document.tenant_region_sub_scope)
            $details.TenantCloudInstance = [string]$document.cloud_instance_name
        } catch {
            $status = Get-HttpStatus $_
            if ($status -ne 400 -and $status -ne 404) {
                $errors += "Tenant discovery: $($_.Exception.Message)"
                $details.Complete = $false
            }
        }

        $details.Errors = $errors -join ' | '
        if ($null -ne $DomainCache -and $details.Complete) {
            $DomainCache[$key] = @{ Time = [datetime]::UtcNow; Details = $details }
        }
        return $details
    }

    function Get-AccountCheck {
        param([string]$Email, $DomainCache = $null, [double]$TtlHours = 24)

        $address = ([string]$Email).Trim()
        $result = [pscustomobject][ordered]@{
            Timestamp             = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            Username              = $address
            FoundInCloud          = 'N/A'
            CloudURL              = ''
            FederationRedirectUrl = ''
            AccountType           = ''
            Category              = 'Unknown'
            Organization          = ''
            TenantId              = ''
            TenantRegion          = ''
            NamespaceType         = ''
            FederationProvider    = ''
            Code                  = ''
            Status                = 'OK'
            StatusLabel           = ''
            NextAction            = ''
            Source                = 'Live'
            Throttled             = $false
            Note                  = ''
            Summary               = ''
            Raw                   = ''
            Error                 = ''
        }

        if (-not (Test-EmailFormat $address)) {
            $result.Status      = 'Invalid'
            $result.Category    = 'Error'
            $result.AccountType = 'Invalid email format.'
            $result.Error       = 'Invalid email format. Please check and try again.'
            $result.Summary     = "Username: $address`r`nError: $($result.Error)"
            Set-StatusFields $result
            return $result
        }

        Add-CoreCounter 'LiveLookups'
        $domain           = $address.Split('@')[1].ToLowerInvariant()
        $isPersonalDomain = $CorePersonalDomains -contains $domain
        $raw              = [ordered]@{}

        # Pass 1: always Commercial.
        try {
            $commercial = Invoke-CredentialTypeLookup -Email $address -Cloud 'Commercial'
            $raw['Commercial'] = $commercial.Response
        } catch {
            $message            = $_.Exception.Message
            $result.Status      = 'Unavailable'
            $result.Category    = 'Error'
            $result.AccountType = $message
            $result.Error       = $message
            $result.Summary     = "Username: $address`r`nError: No valid API response. $message"
            Set-StatusFields $result
            Add-CoreCounter 'Unavailable'
            Set-CoreValue 'LastError' ("{0}  {1}: {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $address, $message)
            return $result
        }

        $final      = $commercial
        $finalCloud = 'Commercial'

        # Pass 2: sovereign clouds only when Commercial says "not found" (1) for a non-personal domain.
        # Throttled sovereign answers are never trusted, because a throttled reply can claim the account exists.
        if ([int]$commercial.Response.IfExistsResult -eq 1 -and -not $isPersonalDomain) {
            foreach ($cloud in @('GCCHigh', 'China')) {
                $lookup = $null
                try {
                    $lookup = Invoke-CredentialTypeLookup -Email $address -Cloud $cloud
                    $raw[$cloud] = $lookup.Response
                } catch {
                    $raw[$cloud] = "Ignored error: $($_.Exception.Message)"
                }
                if ($null -ne $lookup -and -not $lookup.Throttled -and ([int]$lookup.Response.IfExistsResult -in @(0, 5))) {
                    $final      = $lookup
                    $finalCloud = $cloud
                    break
                }
            }
        }

        $response = $final.Response
        $code = [int]$response.IfExistsResult
        $result.Code      = [string]$code
        $result.Throttled = [bool]$final.Throttled

        # Tenant details (not for consumer domains such as outlook.com).
        $details = $null
        if (-not $isPersonalDomain) {
            try { $details = Get-DomainDetails -Email $address -Domain $domain -Cloud $finalCloud -DomainCache $DomainCache -TtlHours $TtlHours } catch { $details = $null }
        }
        if ($null -ne $details) {
            $raw['UserRealm']       = $details.RealmRaw
            $raw['TenantDiscovery'] = $details.DiscoveryRaw
            if ($details.Errors) { $raw['DetailErrors'] = $details.Errors }
            $result.Organization  = $details.Organization
            $result.TenantId      = $details.TenantId
            $result.TenantRegion  = $details.TenantRegion
            $result.NamespaceType = $details.NamespaceType
        }

        # Federation: any of these signals is enough.
        $federationUrl = [string]$response.Credentials.FederationRedirectUrl
        $domainType = 0
        try { $domainType = [int]$response.EstsProperties.DomainType } catch { $domainType = 0 }
        $realmFederated = ($null -ne $details -and $details.NamespaceType -eq 'Federated')
        $isFederated = [bool]$response.IsFederated -or [bool]$federationUrl -or ($domainType -eq 4) -or $realmFederated
        if (-not $federationUrl -and $isFederated -and $null -ne $details -and $details.RealmAuthUrl) {
            $federationUrl = $details.RealmAuthUrl
        }

        $type = ''
        $category = 'Unknown'
        if ($finalCloud -eq 'GCCHigh') {
            $type = 'Likely a GCC High Account.'
            $category = 'GCCHigh'
        } elseif ($finalCloud -eq 'China') {
            $type = 'Likely an M365 China (VNET) Account.'
            $category = 'China'
        } else {
            switch ($code) {
                0 {
                    if ($isPersonalDomain) {
                        $type = 'Likely Personal Microsoft Account (MSA).'; $category = 'MSA'
                    } elseif ($isFederated) {
                        $type = 'Federated Entra ID (Work/School) Account.'; $category = 'Federated'
                    } else {
                        $type = 'Likely Entra Microsoft Account (Work/School).'; $category = 'Entra'
                    }
                }
                1 {
                    $type = 'Not a Microsoft Account.'; $category = 'NotMicrosoft'
                }
                5 {
                    if ($isPersonalDomain) {
                        $type = 'Likely Personal Microsoft Account (MSA).'; $category = 'MSA'
                    } else {
                        $type = 'Federated Entra ID (Work/School) Account.'; $category = 'Federated'
                    }
                }
                6 {
                    $type = 'Dual Identity (Work/School & Personal) Account.'; $category = 'Dual'
                }
                default {
                    $type = "Unknown Result Code: $code. Please try again or check logs."; $category = 'Unknown'
                    $result.Status = 'Inconclusive'
                }
            }
        }

        $result.AccountType = $type
        $result.Category    = $category
        if ($category -ne 'NotMicrosoft' -and $category -ne 'MSA') {
            $result.FederationRedirectUrl = $federationUrl
            if ($federationUrl) { $result.FederationProvider = Get-FederationProvider $federationUrl }
        }

        if ($category -eq 'NotMicrosoft') {
            $result.FoundInCloud = 'N/A'
            $result.CloudURL     = ''
            if ($null -ne $details -and ($details.TenantId -or $details.NamespaceType -eq 'Managed' -or $details.NamespaceType -eq 'Federated')) {
                $orgText = if ($details.Organization) { " ($($details.Organization))" } else { '' }
                $result.Status = 'Inconclusive'
                $result.Note   = "The address was not found, but the domain $domain belongs to a Microsoft tenant$orgText."
            }
        } else {
            $result.FoundInCloud = $finalCloud
            $result.CloudURL     = if ($type -match 'Personal') { 'https://login.live.com' } else { $CoreCloudBase[$finalCloud] }
        }
        if ($result.Throttled) {
            $result.Status = 'Throttled'
            Set-CoreValue 'LastError' ("{0}  {1}: throttled by Microsoft on all {2} attempts" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $address, $CoreMaxAttempts)
        } else {
            Set-CoreValue 'LastSuccess' ("{0}  {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $address)
        }
        if ($result.Status -eq 'Inconclusive') { Add-CoreCounter 'Inconclusive' }
        Set-StatusFields $result

        $cloudUrlText = if ($result.CloudURL) { $result.CloudURL } else { 'N/A' }
        $lines = @(
            "Username: $address",
            "FoundInCloud: $($result.FoundInCloud)",
            "Cloud URL: $cloudUrlText"
        )
        if ($result.FederationRedirectUrl) { $lines += "FederationRedirectUrl: $($result.FederationRedirectUrl)" }
        $lines += "Account Type: $type"
        $result.Summary = $lines -join "`r`n"

        try { $result.Raw = ConvertTo-Json -InputObject $raw -Depth 8 } catch { $result.Raw = '' }
        return $result
    }

    # How long a result may be reused, by outcome. Without a policy only stable answers are cached.
    function Get-CacheTtlHours {
        param([string]$Status, $Policy, [double]$Default = 24)
        if ($null -eq $Policy) {
            if ($Status -eq 'OK' -or $Status -eq 'Inconclusive') { return $Default }
            return 0
        }
        switch ($Status) {
            'OK'           { return [double]$Policy.StableHours }
            'Inconclusive' { return [double]$Policy.InconclusiveHours }
            'Throttled'    { return [double]$Policy.FailureMinutes / 60 }
            'Unavailable'  { return [double]$Policy.FailureMinutes / 60 }
            default        { return 0 }
        }
    }

    # MaxAgeHours > 0 means "refresh stale only": cached answers older than that are looked up again.
    function Get-AccountCheckCached {
        param([string]$Email, $Cache, $DomainCache = $null, [double]$TtlHours = 24, [switch]$SkipCache, $Policy = $null, [double]$MaxAgeHours = 0)
        $key = ([string]$Email).Trim().ToLowerInvariant()
        if ($null -ne $Cache -and $key -and -not $SkipCache) {
            $entry = $Cache[$key]
            if ($null -ne $entry) {
                $age = ([datetime]::UtcNow - $entry.Time).TotalHours
                $ttl = if ($entry.ContainsKey('TtlHours') -and $null -ne $entry['TtlHours']) { [double]$entry['TtlHours'] } else { $TtlHours }
                if ($age -lt $ttl -and ($MaxAgeHours -le 0 -or $age -lt $MaxAgeHours)) {
                    $copy = $entry.Result.PSObject.Copy()
                    $copy.Source    = 'Cache'
                    $copy.Timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
                    Add-CoreCounter 'CacheHits'
                    return $copy
                }
            }
        }
        $domainTtl = if ($null -ne $Policy) { [double]$Policy.StableHours } else { $TtlHours }
        $fresh = Get-AccountCheck -Email $Email -DomainCache $DomainCache -TtlHours $domainTtl
        $keepFor = Get-CacheTtlHours -Status $fresh.Status -Policy $Policy -Default $TtlHours
        if ($null -ne $Cache -and $key -and $keepFor -gt 0) {
            $Cache[$key] = @{ Time = [datetime]::UtcNow; Result = $fresh; TtlHours = $keepFor }
        }
        return $fresh
    }

    function ConvertTo-ApiResult {
        param($Result)
        return [ordered]@{
            username              = $Result.Username
            foundInCloud          = $Result.FoundInCloud
            cloudUrl              = $Result.CloudURL
            federationRedirectUrl = $Result.FederationRedirectUrl
            accountType           = $Result.AccountType
            category              = $Result.Category
            organization          = $Result.Organization
            tenantId              = $Result.TenantId
            tenantRegion          = $Result.TenantRegion
            domainType            = $Result.NamespaceType
            federationProvider    = $Result.FederationProvider
            status                = $Result.Status
            statusLabel           = $Result.StatusLabel
            nextAction            = $Result.NextAction
            changes               = [string]$Result.Changes
            throttled             = [bool]$Result.Throttled
            note                  = $Result.Note
            source                = $Result.Source
            checkedAt             = $Result.Timestamp
            output                = $Result.Summary
            error                 = $Result.Error
        }
    }
}

. $script:CoreLogic
$script:CoreText = $script:CoreLogic.ToString()
#endregion

#region Background scripts (run in their own runspaces)
# Batch worker. Several can share one queue of addresses (parallel lookups); all honour Pause and Stop.
$script:WorkerScript = {
    param($CoreText, $EmailQueue, $Queue, $Sync, $Cache, $DomainCache, $Stats, [int]$DelayMs, [double]$TtlHours, [bool]$SkipCache, $Policy = $null, [double]$MaxAgeHours = 0)
    . ([scriptblock]::Create($CoreText))
    $CoreStats = $Stats
    $email = $null
    while (-not $Sync.Cancel) {
        while ($Sync.Paused -and -not $Sync.Cancel) { Start-Sleep -Milliseconds 200 }
        if ($Sync.Cancel) { break }
        if (-not $EmailQueue.TryDequeue([ref]$email)) { break }
        $r = Get-AccountCheckCached -Email $email -Cache $Cache -DomainCache $DomainCache -TtlHours $TtlHours -SkipCache:$SkipCache -Policy $Policy -MaxAgeHours $MaxAgeHours
        $Queue.Enqueue($r)
        if ($DelayMs -gt 0 -and $r.Source -eq 'Live' -and -not $Sync.Cancel) { Start-Sleep -Milliseconds $DelayMs }
    }
}

$script:ApiScript = {
    param($CoreText, $Listener, $Shared, [datetime]$StartedUtc, [double]$TtlHours, [string]$Version, [string]$WebPage, [int]$Port)
    . ([scriptblock]::Create($CoreText))
    $CoreStats = $Shared.Stats

    function Get-QueryValue {
        param([string]$Query, [string]$Name)
        if ([string]::IsNullOrEmpty($Query)) { return $null }
        foreach ($pair in $Query.TrimStart('?').Split('&')) {
            if (-not $pair) { continue }
            $kv = $pair.Split([char[]]@('='), 2)
            if ([Uri]::UnescapeDataString($kv[0]) -ieq $Name) {
                if ($kv.Count -gt 1) { return [Uri]::UnescapeDataString($kv[1]) }
                return ''
            }
        }
        return $null
    }

    $endpointList = @(
        [ordered]@{ method = 'GET';  path = '/'; description = 'Web page for checking accounts in a browser' }
        [ordered]@{ method = 'GET';  path = '/api/check?email=user@domain.com'; description = 'Check one account' }
        [ordered]@{ method = 'POST'; path = '/api/batch'; description = 'Check up to 200 accounts. Body: ["a@x.com","b@y.com"] or {"emails":[...]}' }
        [ordered]@{ method = 'GET';  path = '/api/history?limit=200&search=text&type=&cloud=&status=&federation=&source=&days='; description = 'Recent checks from History, newest first, with counts. Optional filters: type (Entra, MSA, Federated, Dual, Sovereign, NotMicrosoft, Error), cloud (Commercial, GCCHigh, China, N/A), status (OK, Inconclusive, Throttled, Unavailable, Invalid), federation (yes, no), source (Live, Cache), days (0 = today)' }
        [ordered]@{ method = 'POST'; path = '/api/jobs'; description = 'Submit up to 5,000 addresses as a background job. Returns 202 with the job ID' }
        [ordered]@{ method = 'GET';  path = '/api/jobs'; description = 'List background jobs' }
        [ordered]@{ method = 'GET';  path = '/api/jobs/{id}'; description = 'Job status, progress and results' }
        [ordered]@{ method = 'DELETE'; path = '/api/jobs/{id}'; description = 'Cancel a job' }
        [ordered]@{ method = 'GET';  path = '/api/health'; description = 'Server health' }
        [ordered]@{ method = 'GET';  path = '/api/stats'; description = 'Request counters and lookup diagnostics' }
        [ordered]@{ method = 'GET';  path = '/api/openapi.json'; description = 'OpenAPI 3 description of this API' }
        [ordered]@{ method = 'GET';  path = '/api/endpoints'; description = 'This list' }
        [ordered]@{ method = 'GET';  path = '/health'; description = 'Liveness check (no token needed)' }
        [ordered]@{ method = 'GET';  path = '/metrics'; description = 'Prometheus metrics' }
    )
    $selfOrigins = @("http://localhost:$Port", "http://127.0.0.1:$Port")
    $rateWindows = @{}
    $invariant   = [System.Globalization.CultureInfo]::InvariantCulture

    function Read-EmailList {
        param([string]$Text)
        $parsed = $null
        try { $parsed = ConvertFrom-Json -InputObject $Text } catch { return $null }
        $list = New-Object System.Collections.Generic.List[string]
        if ($null -eq $parsed) { return $list }
        $source = $parsed
        if ($parsed -isnot [string] -and $parsed -isnot [array] -and $null -ne $parsed.PSObject.Properties['emails']) { $source = $parsed.emails }
        foreach ($item in @($source)) { if ($item -is [string] -and $item.Trim()) { $list.Add($item.Trim()) } }
        return ,$list
    }

    function Get-MetricsText {
        $snap = Get-CoreStatsSnapshot -Stats $Shared.Stats
        $lines = New-Object System.Collections.Generic.List[string]
        $add = {
            param([string]$Name, [string]$Type, [string]$Help, $Values)
            $lines.Add("# HELP $Name $Help")
            $lines.Add("# TYPE $Name $Type")
            foreach ($v in $Values) { $lines.Add(('{0}{1} {2}' -f $Name, $v[0], ([double]$v[1]).ToString('0.###', $invariant))) }
        }
        & $add 'mac_api_requests_total' 'counter' 'Requests handled by the local API.' @(,@('', [int]$Shared['RequestCount']))
        & $add 'mac_api_rejected_total' 'counter' 'Requests refused (other website, token or rate limit).' @(,@('', [int]$Shared['RejectedCount']))
        & $add 'mac_lookups_total' 'counter' 'Account lookups by source.' @(@('{source="live"}', $snap.LiveLookups), @('{source="cache"}', $snap.CacheHits))
        & $add 'mac_retries_total' 'counter' 'Requests sent again after throttling or a brief outage.' @(,@('', $snap.Retries))
        & $add 'mac_throttle_events_total' 'counter' 'Times Microsoft rate-limited a request.' @(,@('', $snap.ThrottleEvents))
        & $add 'mac_unavailable_total' 'counter' 'Lookups that failed because Microsoft could not be reached.' @(,@('', $snap.Unavailable))
        & $add 'mac_inconclusive_total' 'counter' 'Lookups with a non-definitive answer.' @(,@('', $snap.Inconclusive))
        & $add 'mac_cache_entries' 'gauge' 'Entries in the lookup cache.' @(,@('', $Shared.Cache.Count))
        & $add 'mac_uptime_seconds' 'gauge' 'Seconds since the API server started.' @(,@('', ([datetime]::UtcNow - $StartedUtc).TotalSeconds))
        $calls = @(); $failures = @(); $latency = @()
        foreach ($e in $snap.Endpoints) {
            $label = '{endpoint="' + ([string]$e.Endpoint).Replace('"', "'") + '"}'
            $calls += ,@($label, $e.Calls); $failures += ,@($label, $e.Failures); $latency += ,@($label, $e.AverageMs)
        }
        & $add 'mac_endpoint_calls_total' 'counter' 'Calls to each Microsoft endpoint, including retries.' $calls
        & $add 'mac_endpoint_failures_total' 'counter' 'Failed calls to each Microsoft endpoint.' $failures
        & $add 'mac_endpoint_latency_ms_average' 'gauge' 'Average response time of each Microsoft endpoint.' $latency
        $availability = if ($null -ne $snap.Recent.AvailabilityPct) { $snap.Recent.AvailabilityPct } else { 100 }
        & $add 'mac_recent_availability_percent' 'gauge' 'Share of Microsoft calls that succeeded in the last 5 minutes.' @(,@('', $availability))
        & $add 'mac_recent_throttle_rate_percent' 'gauge' 'Throttled lookups as a share of lookups in the last 5 minutes.' @(,@('', $snap.Recent.ThrottleRatePct))
        & $add 'mac_recent_latency_ms' 'gauge' 'Lookup response time in the last 5 minutes.' @(@('{quantile="0.5"}', $snap.Recent.MedianMs), @('{quantile="0.95"}', $snap.Recent.P95Ms))
        return (($lines -join "`n") + "`n")
    }

    function Get-OpenApiDocument {
        $ok      = [ordered]@{ description = 'OK' }
        $result  = [ordered]@{ type = 'object'; properties = [ordered]@{
            username = @{ type = 'string' }; foundInCloud = @{ type = 'string' }; cloudUrl = @{ type = 'string' }
            federationRedirectUrl = @{ type = 'string' }; accountType = @{ type = 'string' }
            category = [ordered]@{ type = 'string'; enum = @('Entra', 'MSA', 'Federated', 'Dual', 'GCCHigh', 'China', 'NotMicrosoft', 'Error', 'Unknown') }
            organization = @{ type = 'string' }; tenantId = @{ type = 'string' }; tenantRegion = @{ type = 'string' }; domainType = @{ type = 'string' }
            federationProvider = @{ type = 'string' }
            status = [ordered]@{ type = 'string'; enum = @('OK', 'Inconclusive', 'Throttled', 'Unavailable', 'Invalid') }
            statusLabel = @{ type = 'string' }; nextAction = @{ type = 'string' }; changes = @{ type = 'string' }
            throttled = @{ type = 'boolean' }; note = @{ type = 'string' }; source = [ordered]@{ type = 'string'; enum = @('Live', 'Cache') }
            checkedAt = @{ type = 'string' }; output = @{ type = 'string' }; error = @{ type = 'string' } } }
        $emails  = [ordered]@{ oneOf = @([ordered]@{ type = 'array'; items = @{ type = 'string' } }, [ordered]@{ type = 'object'; properties = [ordered]@{ emails = [ordered]@{ type = 'array'; items = @{ type = 'string' } } } }) }
        $errors  = [ordered]@{ '400' = @{ description = 'Bad request' }; '401' = @{ description = 'Missing or invalid token' }; '403' = @{ description = 'Request from another website' }; '429' = @{ description = 'Rate limit exceeded; see Retry-After' } }
        $withErrors = { param($responses) foreach ($code in $errors.Keys) { $responses[$code] = $errors[$code] }; return $responses }
        $q = { param($name, $description) [ordered]@{ name = $name; in = 'query'; required = $false; description = $description; schema = @{ type = 'string' } } }
        $paths = [ordered]@{
            '/api/check' = [ordered]@{ get = [ordered]@{ summary = 'Check one address'; parameters = @([ordered]@{ name = 'email'; in = 'query'; required = $true; schema = @{ type = 'string' } })
                responses = (& $withErrors ([ordered]@{ '200' = [ordered]@{ description = 'Result'; content = [ordered]@{ 'application/json' = [ordered]@{ schema = $result } } }; '502' = @{ description = 'Microsoft could not be reached' } })) } }
            '/api/batch' = [ordered]@{ post = [ordered]@{ summary = 'Check up to 200 addresses'; requestBody = [ordered]@{ required = $true; content = [ordered]@{ 'application/json' = [ordered]@{ schema = $emails } } }
                responses = (& $withErrors ([ordered]@{ '200' = $ok; '413' = @{ description = 'More than 200 addresses' } })) } }
            '/api/jobs' = [ordered]@{
                post = [ordered]@{ summary = 'Submit up to 5,000 addresses as a background job'; requestBody = [ordered]@{ required = $true; content = [ordered]@{ 'application/json' = [ordered]@{ schema = $emails } } }
                    responses = (& $withErrors ([ordered]@{ '202' = @{ description = 'Accepted; poll statusUrl' }; '413' = @{ description = 'More than 5,000 addresses' } })) }
                get = [ordered]@{ summary = 'List background jobs'; responses = (& $withErrors ([ordered]@{ '200' = $ok })) } }
            '/api/jobs/{id}' = [ordered]@{
                get = [ordered]@{ summary = 'Job status, progress and results'; parameters = @([ordered]@{ name = 'id'; in = 'path'; required = $true; schema = @{ type = 'string' } }); responses = (& $withErrors ([ordered]@{ '200' = $ok; '404' = @{ description = 'No such job' } })) }
                delete = [ordered]@{ summary = 'Cancel a job'; parameters = @([ordered]@{ name = 'id'; in = 'path'; required = $true; schema = @{ type = 'string' } }); responses = (& $withErrors ([ordered]@{ '200' = $ok; '404' = @{ description = 'No such job' } })) } }
            '/api/history' = [ordered]@{ get = [ordered]@{ summary = 'Recent checks, newest first, with counts'
                parameters = @((& $q 'limit' '1 to 5000, default 200'), (& $q 'search' 'Text to look for'), (& $q 'type' 'Entra, MSA, Federated, Dual, Sovereign, NotMicrosoft, Error'),
                               (& $q 'cloud' 'Commercial, GCCHigh, China, N/A'), (& $q 'status' 'OK, Inconclusive, Throttled, Unavailable, Invalid, Changed'),
                               (& $q 'federation' 'yes or no'), (& $q 'source' 'Live or Cache'), (& $q 'days' '0 = today, or a number of days back'))
                responses = (& $withErrors ([ordered]@{ '200' = $ok })) } }
            '/api/health'  = [ordered]@{ get = [ordered]@{ summary = 'Server status'; responses = (& $withErrors ([ordered]@{ '200' = $ok })) } }
            '/api/stats'   = [ordered]@{ get = [ordered]@{ summary = 'Request counters and lookup diagnostics'; responses = (& $withErrors ([ordered]@{ '200' = $ok })) } }
            '/health'      = [ordered]@{ get = [ordered]@{ summary = 'Liveness check; never needs a token'; security = @(); responses = [ordered]@{ '200' = $ok } } }
            '/metrics'     = [ordered]@{ get = [ordered]@{ summary = 'Prometheus metrics'; responses = (& $withErrors ([ordered]@{ '200' = [ordered]@{ description = 'Prometheus text format'; content = [ordered]@{ 'text/plain' = [ordered]@{ schema = @{ type = 'string' } } } } })) } }
        }
        return [ordered]@{
            openapi    = '3.0.3'
            info       = [ordered]@{ title = 'Microsoft Account Checker API'; version = $Version; description = 'Local API. Listens on this computer only and refuses requests from other websites. When a token is required, send it in the X-API-Key header.'; license = [ordered]@{ name = 'CC0-1.0' } }
            servers    = @([ordered]@{ url = "http://localhost:$Port" })
            components = [ordered]@{ securitySchemes = [ordered]@{ apiKey = [ordered]@{ type = 'apiKey'; in = 'header'; name = 'X-API-Key' } } }
            security   = @([ordered]@{ apiKey = @() })
            paths      = $paths
        }
    }
    $openApiJson = $null
    $csp = "default-src 'self'; script-src 'unsafe-inline'; style-src 'unsafe-inline'; img-src 'self' data:; connect-src 'self'; frame-ancestors 'none'; base-uri 'none'; form-action 'self'"

    while ($Listener.IsListening) {
        $context = $null
        try { $context = $Listener.GetContext() } catch { break }

        $timer       = [System.Diagnostics.Stopwatch]::StartNew()
        $request     = $context.Request
        $status      = 200
        $body        = $null
        $textBody    = $null
        $contentType = 'application/json; charset=utf-8'
        $isHtml      = $false
        $rejected    = $false
        $limited     = $false
        $retryAfter  = 0
        $path        = '/'
        $clientAddress = ''
        try { $clientAddress = $request.RemoteEndPoint.Address.ToString() } catch { }

        try {
            $path = $request.Url.AbsolutePath.TrimEnd('/').ToLowerInvariant()
            if (-not $path) { $path = '/' }
            $method = $request.HttpMethod.ToUpperInvariant()

            # Browser safety: refuse anything another website sends, so pages you visit cannot use this API.
            # Opening the web page itself by following a link is allowed; that site cannot read the page.
            $origin     = [string]$request.Headers['Origin']
            $fetchSite  = [string]$request.Headers['Sec-Fetch-Site']
            $fetchMode  = [string]$request.Headers['Sec-Fetch-Mode']
            $isPage     = ($method -eq 'GET' -and ($path -eq '/' -or $path -eq '/ui'))
            $crossSite  = ($fetchSite -and $fetchSite -ne 'same-origin' -and $fetchSite -ne 'none') -or ($origin -and ($selfOrigins -notcontains $origin))
            $navigation = $isPage -and ($fetchMode -eq 'navigate') -and -not $origin

            # Optional token for every /api request (the web page sends it automatically).
            $token = [string]$Shared['ApiToken']
            $presented = [string]$request.Headers['X-API-Key']
            if (-not $presented) {
                $authorization = [string]$request.Headers['Authorization']
                if ($authorization -match '^\s*Bearer\s+(\S+)\s*$') { $presented = $Matches[1] }
            }

            # Rate limit per client (the web page's own refreshes and /health are exempt).
            $limit = [int]$Shared['RateLimitPerMinute']
            if ($limit -gt 0 -and $path -ne '/health' -and $request.Url.Query -notmatch 'client=page' -and -not ($crossSite -and -not $navigation)) {
                $window = $rateWindows[$clientAddress]
                if ($null -eq $window) { $window = New-Object 'System.Collections.Generic.Queue[datetime]'; $rateWindows[$clientAddress] = $window }
                $cutoff = [datetime]::UtcNow.AddMinutes(-1)
                while ($window.Count -gt 0 -and $window.Peek() -lt $cutoff) { [void]$window.Dequeue() }
                if ($window.Count -ge $limit) {
                    $limited = $true
                    $retryAfter = [Math]::Max(1, [int][Math]::Ceiling(60 - ([datetime]::UtcNow - $window.Peek()).TotalSeconds))
                } else {
                    $window.Enqueue([datetime]::UtcNow)
                }
            }

            if ($crossSite -and -not $navigation) {
                $status = 403; $rejected = $true
                $body = [ordered]@{ error = 'Requests from other websites are not allowed.' }
            } elseif ($limited) {
                $status = 429; $rejected = $true
                $body = [ordered]@{ error = "Rate limit of $limit requests a minute exceeded. Try again in $retryAfter seconds."; retryAfterSeconds = $retryAfter }
            } elseif ($method -eq 'OPTIONS') {
                $status = 204
            } elseif ($isPage) {
                $pageToken   = if ($crossSite) { '' } else { $token }
                $textBody    = $WebPage.Replace('{{APP_THEME}}', [string]$Shared['Theme']).Replace('{{API_TOKEN}}', $pageToken)
                $contentType = 'text/html; charset=utf-8'
                $isHtml      = $true
            } elseif ($path -eq '/favicon.ico') {
                $status = 204
            } elseif ($path -eq '/health') {
                $body = [ordered]@{ status = 'ok'; version = $Version }
            } elseif ($token -and ($path.StartsWith('/api') -or $path -eq '/metrics') -and -not [string]::Equals($presented, $token, [System.StringComparison]::Ordinal)) {
                $status = 401; $rejected = $true
                $body = [ordered]@{ error = 'Missing or invalid API token. Send it in the X-API-Key header.' }
            } elseif ($path -eq '/api/check') {
                if ($method -ne 'GET') {
                    $status = 405; $body = [ordered]@{ error = 'Use GET /api/check?email=user@domain.com' }
                } else {
                    $email = Get-QueryValue -Query $request.Url.Query -Name 'email'
                    if ([string]::IsNullOrWhiteSpace($email)) {
                        $status = 400; $body = [ordered]@{ error = "Missing 'email' query parameter." }
                    } else {
                        $r = Get-AccountCheckCached -Email $email -Cache $Shared.Cache -DomainCache $Shared.DomainCache -TtlHours $TtlHours -Policy $Shared['CachePolicy']
                        if ($r.Status -ne 'Invalid') { $Shared.ApiResults.Enqueue($r) }
                        $status = switch ($r.Status) { 'Invalid' { 400 } 'Unavailable' { 502 } default { 200 } }
                        $body = ConvertTo-ApiResult $r
                    }
                }
            } elseif ($path -eq '/api/batch') {
                if ($method -ne 'POST') {
                    $status = 405; $body = [ordered]@{ error = 'Use POST /api/batch with a JSON array of email addresses.' }
                } else {
                    $reader = New-Object System.IO.StreamReader($request.InputStream, [System.Text.Encoding]::UTF8)
                    $text = $reader.ReadToEnd()
                    $reader.Close()
                    $parsed = $null
                    try { $parsed = ConvertFrom-Json -InputObject $text } catch { $parsed = $null }
                    $emails = New-Object System.Collections.Generic.List[string]
                    if ($null -ne $parsed) {
                        $source = $parsed
                        if ($parsed -isnot [string] -and $parsed -isnot [array] -and $null -ne $parsed.PSObject.Properties['emails']) {
                            $source = $parsed.emails
                        }
                        foreach ($item in @($source)) {
                            if ($item -is [string] -and $item.Trim()) { $emails.Add($item.Trim()) }
                        }
                    }
                    if ($emails.Count -eq 0) {
                        $status = 400; $body = [ordered]@{ error = 'Body must be a JSON array of email strings, or {"emails": [...]}.' }
                    } elseif ($emails.Count -gt 200) {
                        $status = 413; $body = [ordered]@{ error = 'At most 200 addresses per request.' }
                    } else {
                        $results = New-Object System.Collections.Generic.List[object]
                        foreach ($email in $emails) {
                            $r = Get-AccountCheckCached -Email $email -Cache $Shared.Cache -DomainCache $Shared.DomainCache -TtlHours $TtlHours -Policy $Shared['CachePolicy']
                            if ($r.Status -ne 'Invalid') { $Shared.ApiResults.Enqueue($r) }
                            $results.Add((ConvertTo-ApiResult $r))
                            if ($r.Source -eq 'Live') { Start-Sleep -Milliseconds 200 }
                        }
                        $body = [ordered]@{ count = $results.Count; results = $results.ToArray() }
                    }
                }
            } elseif ($path -eq '/api/history') {
                $limit = 200
                $limitText = Get-QueryValue -Query $request.Url.Query -Name 'limit'
                $parsedLimit = 0
                if ($limitText -and [int]::TryParse($limitText, [ref]$parsedLimit)) { $limit = [Math]::Max(1, [Math]::Min(5000, $parsedLimit)) }
                $search = ([string](Get-QueryValue -Query $request.Url.Query -Name 'search')).Trim()
                # Optional filters, the same as the desktop app's History tab.
                $filterType       = ([string](Get-QueryValue -Query $request.Url.Query -Name 'type')).Trim()
                $filterCloud      = ([string](Get-QueryValue -Query $request.Url.Query -Name 'cloud')).Trim()
                $filterStatus     = ([string](Get-QueryValue -Query $request.Url.Query -Name 'status')).Trim()
                $filterFederation = ([string](Get-QueryValue -Query $request.Url.Query -Name 'federation')).Trim()
                $filterSource     = ([string](Get-QueryValue -Query $request.Url.Query -Name 'source')).Trim()
                $since = $null
                $daysText = Get-QueryValue -Query $request.Url.Query -Name 'days'
                $days = 0
                if ($daysText -and [int]::TryParse($daysText, [ref]$days) -and $days -ge 0) { $since = (Get-Date).Date.AddDays(-$days) }
                $items = $Shared.History.ToArray()
                $counts = [ordered]@{ Total = 0; Entra = 0; MSA = 0; Federated = 0; Dual = 0; Sovereign = 0; NotMicrosoft = 0; Error = 0 }
                $matched = New-Object System.Collections.Generic.List[object]
                $matchedTotal = 0
                for ($i = $items.Count - 1; $i -ge 0; $i--) {
                    $h = $items[$i]
                    $counts['Total'] = $counts['Total'] + 1
                    $bucket = switch ([string]$h.Category) {
                        'Entra' { 'Entra' } 'MSA' { 'MSA' } 'Federated' { 'Federated' } 'Dual' { 'Dual' }
                        'GCCHigh' { 'Sovereign' } 'China' { 'Sovereign' } 'NotMicrosoft' { 'NotMicrosoft' } default { 'Error' }
                    }
                    $counts[$bucket] = $counts[$bucket] + 1
                    if ($filterType -and $bucket -ne $filterType) { continue }
                    if ($filterCloud -and [string]$h.FoundInCloud -ne $filterCloud) { continue }
                    if ($filterStatus -eq 'Changed') {
                        if (-not [string]$h.Changes) { continue }
                    } elseif ($filterStatus) {
                        $itemStatus = [string]$h.Status
                        if ($itemStatus -eq 'Error') { $itemStatus = 'Unavailable' }
                        if ($itemStatus -ne $filterStatus) { continue }
                    }
                    if ($filterFederation) {
                        $federated = ([string]$h.Category -eq 'Federated') -or [bool]$h.FederationProvider
                        if (($filterFederation -eq 'yes') -ne $federated) { continue }
                    }
                    if ($filterSource -and [string]$h.Source -ne $filterSource) { continue }
                    if ($null -ne $since) {
                        $checked = [datetime]::MinValue
                        $parsed = [datetime]::TryParseExact([string]$h.Timestamp, 'yyyy-MM-dd HH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$checked)
                        if (-not $parsed -or $checked -lt $since) { continue }
                    }
                    if ($search) {
                        $hit = $false
                        foreach ($value in @($h.Username, $h.AccountType, $h.FoundInCloud, $h.Organization, $h.Origin, $h.StatusLabel)) {
                            if (([string]$value).IndexOf($search, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { $hit = $true; break }
                        }
                        if (-not $hit) { continue }
                    }
                    $matchedTotal++
                    if ($matched.Count -lt $limit) {
                        $matched.Add([ordered]@{
                            checkedAt          = $h.Timestamp
                            username           = $h.Username
                            foundInCloud       = $h.FoundInCloud
                            cloudUrl           = $h.CloudURL
                            accountType        = $h.AccountType
                            category           = $h.Category
                            organization       = $h.Organization
                            tenantId           = $h.TenantId
                            tenantRegion       = $h.TenantRegion
                            federationProvider = $h.FederationProvider
                            status             = $h.Status
                            statusLabel        = $h.StatusLabel
                            changes            = [string]$h.Changes
                            source             = $h.Source
                            origin             = $h.Origin
                        })
                    }
                }
                $body = [ordered]@{ total = $items.Count; matched = $matchedTotal; count = $matched.Count; counts = $counts; items = $matched.ToArray() }
            } elseif ($path -eq '/metrics') {
                $textBody    = Get-MetricsText
                $contentType = 'text/plain; version=0.0.4; charset=utf-8'
            } elseif ($path -eq '/api/openapi.json') {
                if ($null -eq $openApiJson) { $openApiJson = ConvertTo-Json -InputObject (Get-OpenApiDocument) -Depth 20 }
                $textBody = $openApiJson
            } elseif ($path -eq '/api/jobs') {
                if ($method -eq 'POST') {
                    $reader = New-Object System.IO.StreamReader($request.InputStream, [System.Text.Encoding]::UTF8)
                    $emails = Read-EmailList -Text $reader.ReadToEnd()
                    $reader.Close()
                    if ($null -eq $emails -or $emails.Count -eq 0) {
                        $status = 400; $body = [ordered]@{ error = 'Body must be a JSON array of email strings, or {"emails": [...]}.' }
                    } elseif ($emails.Count -gt 5000) {
                        $status = 413; $body = [ordered]@{ error = 'At most 5,000 addresses per job.' }
                    } else {
                        $id = [guid]::NewGuid().ToString('N').Substring(0, 12)
                        $job = [hashtable]::Synchronized(@{
                            Id = $id; Status = 'queued'; Total = $emails.Count; Done = 0; Cancel = $false
                            Created = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'); Finished = ''; Emails = $emails.ToArray()
                            Results = [System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList))
                        })
                        $Shared.ApiJobs[$id] = $job
                        $Shared.ApiJobQueue.Enqueue($id)
                        # Keep the 100 newest jobs.
                        $finished = @($Shared.ApiJobs.Values | Where-Object { $_.Status -eq 'done' -or $_.Status -eq 'cancelled' } | Sort-Object { $_.Created })
                        while ($Shared.ApiJobs.Count -gt 100 -and $finished.Count -gt 0) { $Shared.ApiJobs.Remove($finished[0].Id); $finished = @($finished | Select-Object -Skip 1) }
                        $status = 202
                        $body = [ordered]@{ id = $id; status = 'queued'; total = $emails.Count; statusUrl = "/api/jobs/$id" }
                    }
                } elseif ($method -eq 'GET') {
                    $list = @($Shared.ApiJobs.Values | Sort-Object { $_.Created } -Descending | ForEach-Object {
                        [ordered]@{ id = $_.Id; status = $_.Status; total = $_.Total; done = $_.Done; created = $_.Created; finished = $_.Finished }
                    })
                    $body = [ordered]@{ count = $list.Count; jobs = $list }
                } else {
                    $status = 405; $body = [ordered]@{ error = 'Use POST to submit a job or GET to list jobs.' }
                }
            } elseif ($path -like '/api/jobs/*') {
                $id = $path.Substring(10)
                $job = $Shared.ApiJobs[$id]
                if ($null -eq $job) {
                    $status = 404; $body = [ordered]@{ error = "No job with ID $id." }
                } elseif ($method -eq 'DELETE') {
                    $job.Cancel = $true
                    if ($job.Status -eq 'queued') { $job.Status = 'cancelled'; $job.Finished = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss') }
                    $body = [ordered]@{ id = $id; status = $job.Status; cancelRequested = $true }
                } elseif ($method -eq 'GET') {
                    $body = [ordered]@{
                        id = $id; status = $job.Status; total = $job.Total; done = $job.Done
                        created = $job.Created; finished = $job.Finished; results = @($job.Results.ToArray())
                    }
                } else {
                    $status = 405; $body = [ordered]@{ error = 'Use GET for status or DELETE to cancel.' }
                }
            } elseif ($path -eq '/api/health') {
                $body = [ordered]@{
                    status        = 'ok'
                    version       = $Version
                    theme         = [string]$Shared['Theme']
                    tokenRequired = [bool]$token
                    uptimeSeconds = [int]([datetime]::UtcNow - $StartedUtc).TotalSeconds
                }
            } elseif ($path -eq '/api/stats') {
                $body = [ordered]@{
                    requests     = [int]$Shared['RequestCount']
                    rejected     = [int]$Shared['RejectedCount']
                    cacheEntries = $Shared.Cache.Count
                    startedUtc   = $StartedUtc.ToString('o')
                    lookups      = Get-CoreStatsSnapshot -Stats $Shared.Stats
                }
            } elseif ($path -eq '/api' -or $path -eq '/api/endpoints') {
                $body = [ordered]@{ name = 'Microsoft Account Checker API'; version = $Version; endpoints = $endpointList }
            } else {
                $status = 404; $body = [ordered]@{ error = 'Not found. See /api/endpoints.'; path = $request.Url.AbsolutePath }
            }
        } catch {
            $status = 500
            $body = [ordered]@{ error = $_.Exception.Message }
        }

        try {
            $response = $context.Response
            $payload = if ($null -ne $textBody) { $textBody } elseif ($null -ne $body) { ConvertTo-Json -InputObject $body -Depth 6 } else { '' }
            $bytes = [System.Text.Encoding]::UTF8.GetBytes($payload)
            $response.StatusCode = $status
            $response.AddHeader('Cache-Control', 'no-store')
            $response.AddHeader('X-Content-Type-Options', 'nosniff')
            $response.AddHeader('Referrer-Policy', 'no-referrer')
            if ($status -eq 429 -and $retryAfter -gt 0) { $response.AddHeader('Retry-After', [string]$retryAfter) }
            if ($isHtml) {
                $response.AddHeader('X-Frame-Options', 'DENY')
                $response.AddHeader('Content-Security-Policy', $csp)
            }
            $response.ContentType = $contentType
            $response.ContentLength64 = $bytes.Length
            if ($bytes.Length -gt 0) { $response.OutputStream.Write($bytes, 0, $bytes.Length) }
            $response.OutputStream.Close()
        } catch { }

        $Shared['RequestCount'] = [int]$Shared['RequestCount'] + 1
        if ($rejected) { $Shared['RejectedCount'] = [int]$Shared['RejectedCount'] + 1 }
        $Shared['LastRequest'] = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        # The web page's own background refreshes and favicon requests are not shown in the request log.
        if (-not $rejected -and ($path -eq '/favicon.ico' -or $request.Url.Query -match 'client=page')) { continue }
        $client = ''
        try { $client = $request.RemoteEndPoint.Address.ToString() } catch { }
        $Shared.ApiLog.Enqueue([pscustomobject][ordered]@{
            Time       = (Get-Date).ToString('HH:mm:ss')
            Method     = $request.HttpMethod
            Path       = $request.Url.PathAndQuery
            Status     = $status
            DurationMs = $timer.ElapsedMilliseconds
            Client     = $client
        })
    }
}

$script:ApiTestScript = {
    param([string]$Url, [string]$Token)
    $response = $null
    try {
        $request = [System.Net.WebRequest]::Create($Url)
        $request.Method  = 'GET'
        $request.Timeout = 120000
        if ($Token) { $request.Headers.Add('X-API-Key', $Token) }
        $response = $request.GetResponse()
    } catch {
        $ex = $_.Exception
        while ($null -ne $ex -and -not ($ex -is [System.Net.WebException])) { $ex = $ex.InnerException }
        if ($null -ne $ex -and $null -ne $ex.Response) { $response = $ex.Response }
        else { return "Request failed: $($_.Exception.Message)" }
    }
    try {
        $reader = New-Object System.IO.StreamReader($response.GetResponseStream())
        $text = $reader.ReadToEnd()
        $reader.Close()
        return "HTTP $([int]$response.StatusCode) $($response.StatusDescription)`r`n`r`n$text"
    } finally {
        $response.Close()
    }
}

# Runs jobs submitted to POST /api/jobs, one at a time, beside the API server.
$script:JobRunnerScript = {
    param($CoreText, $Shared, [double]$TtlHours, $Control)
    . ([scriptblock]::Create($CoreText))
    $CoreStats = $Shared.Stats
    $id = $null
    while (-not $Control.Stop) {
        if (-not $Shared.ApiJobQueue.TryDequeue([ref]$id)) { Start-Sleep -Milliseconds 300; continue }
        $job = $Shared.ApiJobs[$id]
        if ($null -eq $job) { continue }
        if ($job.Cancel) { $job.Status = 'cancelled'; $job.Finished = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'); continue }
        $job.Status = 'running'
        foreach ($email in $job.Emails) {
            if ($Control.Stop -or $job.Cancel) { break }
            $r = Get-AccountCheckCached -Email $email -Cache $Shared.Cache -DomainCache $Shared.DomainCache -TtlHours $TtlHours -Policy $Shared['CachePolicy']
            if ($r.Status -ne 'Invalid') { $Shared.ApiResults.Enqueue($r) }
            [void]$job.Results.Add((ConvertTo-ApiResult $r))
            $job.Done = $job.Done + 1
            $delay = [int]$Shared['BatchDelayMs']
            if ($delay -gt 0 -and $r.Source -eq 'Live') { Start-Sleep -Milliseconds $delay }
        }
        $job.Status   = if ($job.Cancel -or $Control.Stop) { 'cancelled' } else { 'done' }
        $job.Finished = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    }
}

# Update checker: reads the latest GitHub release.
$script:UpdateCheckScript = {
    param([string]$ApiUrl)
    try {
        $release = Invoke-RestMethod -Uri $ApiUrl -Headers @{ 'User-Agent' = 'MS-Account-Checker'; Accept = 'application/vnd.github+json' } -TimeoutSec 20 -ErrorAction Stop
        $assets = @($release.assets | ForEach-Object { [pscustomobject]@{ Name = [string]$_.name; Url = [string]$_.browser_download_url } })
        return [pscustomobject]@{ Ok = $true; Tag = [string]$release.tag_name; Url = [string]$release.html_url; Notes = [string]$release.body
                                  Published = [string]$release.published_at; Assets = $assets; Error = '' }
    } catch {
        $code = 0
        try { $code = [int]$_.Exception.Response.StatusCode } catch { }
        $message = if ($code -eq 404) { 'No release has been published yet.' } else { $_.Exception.Message }
        return [pscustomobject]@{ Ok = $false; Error = $message }
    }
}

# Downloads a release zip, checks it against the published SHA-256 checksum, and reports the
# Authenticode signature of the script inside. It never installs anything.
$script:UpdateDownloadScript = {
    param([string]$ZipUrl, [string]$ZipName, [string]$SumsUrl, [string]$Folder)
    $headers = @{ 'User-Agent' = 'MS-Account-Checker' }
    try {
        if (-not (Test-Path -LiteralPath $Folder)) { [void](New-Item -ItemType Directory -Path $Folder -Force) }
        $zipPath = Join-Path $Folder $ZipName
        Invoke-WebRequest -Uri $ZipUrl -OutFile $zipPath -Headers $headers -UseBasicParsing -TimeoutSec 300 -ErrorAction Stop
        $actual = (Get-FileHash -LiteralPath $zipPath -Algorithm SHA256).Hash.ToLowerInvariant()
        $expected = ''
        if ($SumsUrl) {
            $sumsPath = Join-Path ([System.IO.Path]::GetTempPath()) ('msac-sums-' + [guid]::NewGuid().ToString('N') + '.txt')
            Invoke-WebRequest -Uri $SumsUrl -OutFile $sumsPath -Headers $headers -UseBasicParsing -TimeoutSec 60 -ErrorAction Stop
            foreach ($line in (Get-Content -LiteralPath $sumsPath)) {
                if ($line -match '^\s*([0-9a-fA-F]{64})\s+\*?(.+?)\s*$' -and $Matches[2] -eq $ZipName) { $expected = $Matches[1].ToLowerInvariant() }
            }
            Remove-Item -LiteralPath $sumsPath -Force -ErrorAction SilentlyContinue
        }
        if ($expected -and $expected -ne $actual) {
            Remove-Item -LiteralPath $zipPath -Force -ErrorAction SilentlyContinue
            return [pscustomobject]@{ Ok = $false; Error = 'The download did not match the published SHA-256 checksum, so it was deleted. Try again, or download it from GitHub yourself.' }
        }
        $signature = 'Not signed'
        try {
            $temp = Join-Path ([System.IO.Path]::GetTempPath()) ('msac-' + [guid]::NewGuid().ToString('N'))
            Expand-Archive -LiteralPath $zipPath -DestinationPath $temp -Force
            $scriptFile = Get-ChildItem -LiteralPath $temp -Recurse -Filter 'MS_Account_Checker.ps1' | Select-Object -First 1
            if ($scriptFile) {
                $sig = Get-AuthenticodeSignature -FilePath $scriptFile.FullName
                $signature = switch ([string]$sig.Status) {
                    'Valid'     { "Signed by $($sig.SignerCertificate.Subject)" }
                    'NotSigned' { 'Not signed' }
                    default     { "Signature problem: $($sig.Status)" }
                }
            }
            Remove-Item -LiteralPath $temp -Recurse -Force -ErrorAction SilentlyContinue
        } catch { $signature = 'The signature could not be checked.' }
        return [pscustomobject]@{ Ok = $true; Path = $zipPath; Hash = $actual; Verified = [bool]$expected; Signature = $signature; Error = '' }
    } catch {
        return [pscustomobject]@{ Ok = $false; Error = $_.Exception.Message }
    }
}

# Shared state for the window, the batch workers and the API server (also used by the tests).
function New-SharedState {
    $stats = [hashtable]::Synchronized(@{
        LiveLookups = 0; CacheHits = 0; Retries = 0; ThrottleEvents = 0; Unavailable = 0; Inconclusive = 0
        LastSuccess = ''; LastError = ''; Since = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'); Endpoints = @{}
        Samples = New-Object 'System.Collections.Generic.List[object]'
        ThrottleTimes = New-Object 'System.Collections.Generic.List[datetime]'
    })
    return [hashtable]::Synchronized(@{
        Cache         = [hashtable]::Synchronized(@{})
        DomainCache   = [hashtable]::Synchronized(@{})
        ApiResults    = New-Object 'System.Collections.Concurrent.ConcurrentQueue[object]'
        ApiLog        = New-Object 'System.Collections.Concurrent.ConcurrentQueue[object]'
        History       = [System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList))
        Stats         = $stats
        RequestCount  = 0
        RejectedCount = 0
        LastRequest   = ''
        Theme         = 'light'
        ApiToken      = ''
        CachePolicy   = @{ StableHours = 24.0; InconclusiveHours = 24.0; FailureMinutes = 0.0 }
        RateLimitPerMinute = 120
        BatchDelayMs  = 200
        ApiJobs       = [hashtable]::Synchronized(@{})
        ApiJobQueue   = New-Object 'System.Collections.Concurrent.ConcurrentQueue[string]'
    })
}
#endregion

#region Pure helpers (shared by the window, the command-line mode, the module and the tests)
function Get-CategoryLabelText {
    param([string]$Category)
    $labels = @{
        Entra = 'Work/School (Entra ID)'; MSA = 'Personal (MSA)'; Federated = 'Federated (Work/School)'; Dual = 'Dual identity'
        GCCHigh = 'GCC High'; China = 'M365 China'; NotMicrosoft = 'Not a Microsoft account'; Error = 'Error'; Unknown = 'Unknown'
    }
    if ($labels.ContainsKey($Category)) { return $labels[$Category] }
    return $Category
}

function Test-IsFederated {
    param($Item)
    return ([string]$Item.Category -eq 'Federated') -or [bool]$Item.FederationProvider
}

# What changed for an address since its previous check: account type, cloud, tenant or federation.
# Only stable answers are compared, so a throttled or failed check never reports a change.
function Get-ResultChanges {
    param($Previous, $Current)
    if ($null -eq $Previous -or $null -eq $Current) { return '' }
    $stable = @('OK', 'Inconclusive')
    if ($stable -notcontains [string]$Previous.Status -or $stable -notcontains [string]$Current.Status) { return '' }
    $changes = New-Object System.Collections.Generic.List[string]
    $typeChanged = [string]$Previous.Category -ne [string]$Current.Category
    if ($typeChanged) {
        $changes.Add(('account type {0} -> {1}' -f (Get-CategoryLabelText ([string]$Previous.Category)), (Get-CategoryLabelText ([string]$Current.Category))))
    }
    if ([string]$Previous.FoundInCloud -ne [string]$Current.FoundInCloud) {
        $changes.Add(('cloud {0} -> {1}' -f $Previous.FoundInCloud, $Current.FoundInCloud))
    }
    $oldTenant = [string]$Previous.TenantId
    $newTenant = [string]$Current.TenantId
    if ($oldTenant -ne $newTenant -and ($oldTenant -or $newTenant)) {
        $from = if ($oldTenant) { $oldTenant } else { 'none' }
        $to   = if ($newTenant) { $newTenant } else { 'none' }
        $changes.Add("tenant $from -> $to")
    }
    if (-not $typeChanged) {
        $wasFederated = Test-IsFederated $Previous
        $isFederated  = Test-IsFederated $Current
        if ($wasFederated -ne $isFederated) {
            $changes.Add($(if ($isFederated) { 'now federated' } else { 'no longer federated' }))
        } elseif ($isFederated -and $Previous.FederationProvider -and $Current.FederationProvider -and [string]$Previous.FederationProvider -ne [string]$Current.FederationProvider) {
            $changes.Add(('federation provider {0} -> {1}' -f $Previous.FederationProvider, $Current.FederationProvider))
        }
    }
    return ($changes -join '; ')
}

# One row per domain: addresses, organization, tenant, clouds, federation providers and account-type mix.
function Get-DomainSummary {
    param($Items)
    $groups = [ordered]@{}
    foreach ($item in @($Items)) {
        $user = [string]$item.Username
        $at = $user.LastIndexOf('@')
        if ($at -lt 1) { continue }
        $domain = $user.Substring($at + 1).ToLowerInvariant()
        if (-not $groups.Contains($domain)) {
            $groups[$domain] = @{ Addresses = 0; Orgs = [ordered]@{}; Tenants = [ordered]@{}; Clouds = [ordered]@{}; Providers = [ordered]@{}; Types = [ordered]@{}; NotFound = 0; Attention = 0; Changed = 0 }
        }
        $g = $groups[$domain]
        $g['Addresses'] = $g['Addresses'] + 1
        foreach ($pair in @(@('Orgs', $item.Organization), @('Tenants', $item.TenantId), @('Clouds', $item.FoundInCloud), @('Providers', $item.FederationProvider))) {
            $value = [string]$pair[1]
            if ($value -and $value -ne 'N/A') { $g[$pair[0]][$value] = $true }
        }
        $label = Get-CategoryLabelText ([string]$item.Category)
        $g['Types'][$label] = [int]$g['Types'][$label] + 1
        if ([string]$item.Category -eq 'NotMicrosoft') { $g['NotFound'] = $g['NotFound'] + 1 }
        if ([string]$item.Status -ne 'OK') { $g['Attention'] = $g['Attention'] + 1 }
        if ([string]$item.Changes) { $g['Changed'] = $g['Changed'] + 1 }
    }
    $rows = foreach ($domain in @($groups.Keys)) {
        $g = $groups[$domain]
        $types = $g['Types']
        [pscustomobject][ordered]@{
            Domain             = $domain
            Addresses          = $g['Addresses']
            Organization       = @($g['Orgs'].Keys) -join ', '
            TenantId           = @($g['Tenants'].Keys) -join ', '
            Cloud              = @($g['Clouds'].Keys) -join ', '
            FederationProvider = @($g['Providers'].Keys) -join ', '
            AccountTypes       = (@($types.Keys | ForEach-Object { '{0} {1}' -f $_, $types[$_] }) -join ', ')
            NotFound           = $g['NotFound']
            NeedsAttention     = $g['Attention']
            Changed            = $g['Changed']
        }
    }
    return @($rows | Sort-Object -Property @{ Expression = 'Addresses'; Descending = $true }, @{ Expression = 'Domain'; Descending = $false })
}

function ConvertTo-RedactedEmail {
    param([string]$Email)
    if (-not $Email) { return '' }
    $at = $Email.LastIndexOf('@')
    if ($at -lt 1) { return '***' }
    return $Email.Substring(0, 1) + '***' + $Email.Substring($at)
}

# Reads addresses from a .csv or .txt file. For CSV files the email column is found by name or,
# failing that, by content; the other columns are kept so they can travel with the results.
function Import-EmailFile {
    param([string]$Path, [string]$EmailColumn = '', [string[]]$KeepColumns = $null)
    $finder = [regex]'[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}'
    $result = @{ Emails = New-Object System.Collections.Generic.List[string]; Columns = @(); EmailColumn = ''; KeepColumns = @(); Extras = @{} }
    if ([System.IO.Path]::GetExtension($Path) -ieq '.csv') {
        $rows = @(Import-Csv -LiteralPath $Path)
        if ($rows.Count -gt 0) {
            $columns = @($rows[0].PSObject.Properties.Name)
            $result.Columns = $columns
            $column = $EmailColumn
            if (-not $column -or $columns -notcontains $column) {
                $column = $columns | Where-Object { $_ -match '^(e-?mail|mail|email ?address|e-?mail ?address|userprincipalname|upn|username|user)$' } | Select-Object -First 1
            }
            if (-not $column) {
                $best = 0
                foreach ($name in $columns) {
                    $hits = @($rows | Select-Object -First 50 | Where-Object { $finder.IsMatch([string]$_.$name) }).Count
                    if ($hits -gt $best) { $best = $hits; $column = $name }
                }
            }
            if ($column) {
                $result.EmailColumn = $column
                $keep = if ($null -ne $KeepColumns) { @($KeepColumns | Where-Object { $columns -contains $_ -and $_ -ne $column }) } else { @($columns | Where-Object { $_ -ne $column }) }
                $result.KeepColumns = $keep
                foreach ($row in $rows) {
                    $value = ([string]$row.$column).Trim()
                    if (-not $value) { continue }
                    $result.Emails.Add($value)
                    if ($keep.Count -gt 0) {
                        $extra = [ordered]@{}
                        foreach ($name in $keep) { $extra[$name] = [string]$row.$name }
                        $result.Extras[$value.ToLowerInvariant()] = $extra
                    }
                }
                return $result
            }
        }
    }
    $text = [System.IO.File]::ReadAllText($Path)
    foreach ($match in $finder.Matches($text)) { $result.Emails.Add($match.Value) }
    return $result
}

function Get-CliExitCode {
    param($Results)
    $statuses = @($Results | ForEach-Object { [string]$_.status })
    if ($statuses -contains 'Unavailable') { return 2 }
    if ($statuses -contains 'Inconclusive' -or $statuses -contains 'Throttled' -or $statuses -contains 'Invalid') { return 1 }
    return 0
}

function Invoke-CliCheck {
    param([string[]]$Addresses, [string]$Path, [bool]$UseCache)
    $list = New-Object System.Collections.Generic.List[string]
    $extras = @{}
    $extraColumns = @()
    foreach ($value in @($Addresses)) {
        foreach ($token in ([string]$value -split '[\s,;]+')) { if ($token) { $list.Add($token.Trim()) } }
    }
    if ($Path) {
        if (-not (Test-Path -LiteralPath $Path)) { Write-Error "Input file not found: $Path"; return @{ Rows = @(); Code = 3 } }
        $import = Import-EmailFile -Path $Path
        foreach ($address in $import.Emails) { $list.Add($address) }
        $extras = $import.Extras
        $extraColumns = @($import.KeepColumns)
    }
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $unique = @($list | Where-Object { $_ -and $seen.Add($_) })
    if ($unique.Count -eq 0) { Write-Error 'No email addresses to check.'; return @{ Rows = @(); Code = 3 } }

    $shared = New-SharedState
    $CoreStats = $shared.Stats
    $cache = if ($UseCache) { $shared.Cache } else { $null }
    $rows = New-Object System.Collections.Generic.List[object]
    for ($i = 0; $i -lt $unique.Count; $i++) {
        $address = $unique[$i]
        if ($unique.Count -gt 1) { Write-Progress -Activity 'Checking Microsoft accounts' -Status $address -PercentComplete (100 * $i / $unique.Count) }
        $r = Get-AccountCheckCached -Email $address -Cache $cache -DomainCache $shared.DomainCache -Policy $shared['CachePolicy']
        $row = ConvertTo-ApiResult $r
        $extra = $extras[$address.ToLowerInvariant()]
        foreach ($name in $extraColumns) {
            $key = if ($row.Contains($name)) { "input_$name" } else { $name }
            $row[$key] = if ($null -ne $extra) { [string]$extra[$name] } else { '' }
        }
        $rows.Add([pscustomobject]$row)
        if ($r.Source -eq 'Live' -and $i -lt ($unique.Count - 1)) { Start-Sleep -Milliseconds 200 }
    }
    if ($unique.Count -gt 1) { Write-Progress -Activity 'Checking Microsoft accounts' -Completed }
    return @{ Rows = $rows.ToArray(); Code = (Get-CliExitCode -Results $rows) }
}
#endregion

#region Command-line mode (no window): -Email or -InputFile
if ($script:CliMode) {
    $cli = Invoke-CliCheck -Addresses $Email -Path $InputFile -UseCache (-not $NoCache)
    $rows = @($cli.Rows)
    if ($rows.Count -gt 0) {
        $effective = if ($OutFile -and $Format -eq 'Table') { if ($OutFile -like '*.json') { 'Json' } else { 'Csv' } } else { $Format }
        switch ($effective) {
            'Object' { $rows }
            'Json' {
                $json = ConvertTo-Json -InputObject $rows -Depth 5
                if ($OutFile) { Set-Content -LiteralPath $OutFile -Value $json -Encoding UTF8 } else { $json }
            }
            'Csv' {
                if ($OutFile) { $rows | Export-Csv -LiteralPath $OutFile -NoTypeInformation -Encoding UTF8 } else { $rows | ConvertTo-Csv -NoTypeInformation }
            }
            default {
                $rows | Select-Object username, foundInCloud, accountType, statusLabel, organization, tenantId | Format-Table -AutoSize | Out-Host
            }
        }
        if ($OutFile) { Write-Host "Saved $($rows.Count) results to $OutFile" }
    }
    if ($MyInvocation.InvocationName -eq '.') { $global:LASTEXITCODE = $cli.Code; return }
    exit $cli.Code
}
#endregion

#region Window layout (XAML)
$script:Xaml = @'
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Microsoft Account Checker"
    Width="1180" Height="780" MinWidth="960" MinHeight="620"
    WindowStartupLocation="Manual"
    FontFamily="Segoe UI" FontSize="13"
    Background="{DynamicResource Bg}" Foreground="{DynamicResource Text}"
    UseLayoutRounding="True" SnapsToDevicePixels="True"
    TextOptions.TextFormattingMode="Display">
  <Window.Resources>

    <Style x:Key="FocusRing">
      <Setter Property="Control.Template">
        <Setter.Value>
          <ControlTemplate>
            <Rectangle Margin="-3" RadiusX="8" RadiusY="8" Stroke="{DynamicResource Primary}" StrokeThickness="2" SnapsToDevicePixels="True"/>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="BtnBase" TargetType="Button">
      <Setter Property="Background" Value="{DynamicResource Secondary}"/>
      <Setter Property="Foreground" Value="{DynamicResource SecondaryText}"/>
      <Setter Property="BorderBrush" Value="Transparent"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Padding" Value="14,7"/>
      <Setter Property="MinHeight" Value="34"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="FocusVisualStyle" Value="{StaticResource FocusRing}"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="6" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" RecognizesAccessKey="True"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="bd" Property="Opacity" Value="0.88"/></Trigger>
              <Trigger Property="IsPressed" Value="True"><Setter TargetName="bd" Property="Opacity" Value="0.72"/></Trigger>
              <Trigger Property="IsEnabled" Value="False"><Setter TargetName="bd" Property="Opacity" Value="0.45"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="Button" BasedOn="{StaticResource BtnBase}"/>
    <Style x:Key="PrimaryButton" TargetType="Button" BasedOn="{StaticResource BtnBase}">
      <Setter Property="Background" Value="{DynamicResource Primary}"/>
      <Setter Property="Foreground" Value="{DynamicResource PrimaryText}"/>
    </Style>
    <Style x:Key="DangerButton" TargetType="Button" BasedOn="{StaticResource BtnBase}">
      <Setter Property="Background" Value="{DynamicResource Danger}"/>
      <Setter Property="Foreground" Value="{DynamicResource DangerText}"/>
    </Style>
    <Style x:Key="HeaderButton" TargetType="Button" BasedOn="{StaticResource BtnBase}">
      <Setter Property="Background" Value="{DynamicResource HeaderButtonBg}"/>
      <Setter Property="Foreground" Value="{DynamicResource HeaderText}"/>
      <Setter Property="BorderBrush" Value="{DynamicResource HeaderButtonBorder}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="MinHeight" Value="30"/>
      <Setter Property="Padding" Value="12,5"/>
    </Style>

    <Style TargetType="TextBox">
      <Setter Property="Background" Value="{DynamicResource InputBg}"/>
      <Setter Property="Foreground" Value="{DynamicResource Text}"/>
      <Setter Property="BorderBrush" Value="{DynamicResource Border}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding" Value="10,7"/>
      <Setter Property="VerticalScrollBarVisibility" Value="Disabled"/>
      <Setter Property="CaretBrush" Value="{DynamicResource Text}"/>
      <Setter Property="SelectionBrush" Value="{DynamicResource Primary}"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TextBox">
            <Border x:Name="bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="6">
              <ScrollViewer x:Name="PART_ContentHost" Margin="{TemplateBinding Padding}" Focusable="False"
                            VerticalAlignment="{TemplateBinding VerticalContentAlignment}"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsKeyboardFocusWithin" Value="True"><Setter TargetName="bd" Property="BorderBrush" Value="{DynamicResource Primary}"/></Trigger>
              <Trigger Property="IsEnabled" Value="False"><Setter TargetName="bd" Property="Opacity" Value="0.6"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="MonoBox" TargetType="TextBox" BasedOn="{StaticResource {x:Type TextBox}}">
      <Setter Property="FontFamily" Value="Consolas"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="IsReadOnly" Value="True"/>
      <Setter Property="AcceptsReturn" Value="True"/>
      <Setter Property="TextWrapping" Value="NoWrap"/>
      <Setter Property="VerticalScrollBarVisibility" Value="Auto"/>
      <Setter Property="HorizontalScrollBarVisibility" Value="Auto"/>
      <Setter Property="VerticalContentAlignment" Value="Stretch"/>
      <Setter Property="Background" Value="{DynamicResource GridBg}"/>
    </Style>
    <Style x:Key="ValueBox" TargetType="TextBox" BasedOn="{StaticResource {x:Type TextBox}}">
      <Setter Property="IsReadOnly" Value="True"/>
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Padding" Value="0"/>
      <Setter Property="Margin" Value="0,6,0,6"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="TextWrapping" Value="Wrap"/>
      <Setter Property="VerticalContentAlignment" Value="Top"/>
    </Style>
    <Style x:Key="ChangeText" TargetType="TextBlock">
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="TextTrimming" Value="CharacterEllipsis"/>
      <Setter Property="Foreground" Value="{DynamicResource CatDual}"/>
      <Setter Property="ToolTip" Value="{Binding Changes}"/>
    </Style>
    <Style x:Key="StatusText" TargetType="TextBlock">
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="TextTrimming" Value="CharacterEllipsis"/>
      <Setter Property="ToolTip" Value="{Binding NextAction}"/>
      <Style.Triggers>
        <DataTrigger Binding="{Binding Status}" Value="Throttled"><Setter Property="Foreground" Value="{DynamicResource CatError}"/></DataTrigger>
        <DataTrigger Binding="{Binding Status}" Value="Inconclusive"><Setter Property="Foreground" Value="{DynamicResource CatError}"/></DataTrigger>
        <DataTrigger Binding="{Binding Status}" Value="Unavailable"><Setter Property="Foreground" Value="{DynamicResource CatNotMicrosoft}"/></DataTrigger>
        <DataTrigger Binding="{Binding Status}" Value="Error"><Setter Property="Foreground" Value="{DynamicResource CatNotMicrosoft}"/></DataTrigger>
        <DataTrigger Binding="{Binding Status}" Value="Invalid"><Setter Property="Foreground" Value="{DynamicResource CatNotMicrosoft}"/></DataTrigger>
      </Style.Triggers>
    </Style>

    <Style x:Key="Switch" TargetType="CheckBox">
      <Setter Property="Foreground" Value="{DynamicResource HeaderText}"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="FocusVisualStyle" Value="{StaticResource FocusRing}"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="CheckBox">
            <StackPanel Orientation="Horizontal" Background="Transparent">
              <Grid Width="40" Height="22" VerticalAlignment="Center">
                <Border x:Name="track" CornerRadius="11" Background="{DynamicResource SwitchTrack}" BorderBrush="{DynamicResource SwitchTrackBorder}" BorderThickness="1"/>
                <Ellipse x:Name="thumb" Width="16" Height="16" Fill="{DynamicResource SwitchThumb}" HorizontalAlignment="Left" Margin="3,0,0,0"/>
              </Grid>
              <ContentPresenter Margin="8,0,0,0" VerticalAlignment="Center" RecognizesAccessKey="True"/>
            </StackPanel>
            <ControlTemplate.Triggers>
              <Trigger Property="IsChecked" Value="True">
                <Setter TargetName="track" Property="Background" Value="{DynamicResource SwitchOn}"/>
                <Setter TargetName="track" Property="BorderBrush" Value="{DynamicResource SwitchOn}"/>
                <Setter TargetName="thumb" Property="HorizontalAlignment" Value="Right"/>
                <Setter TargetName="thumb" Property="Margin" Value="0,0,3,0"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.5"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="BodySwitch" TargetType="CheckBox">
      <Setter Property="Foreground" Value="{DynamicResource Text}"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="FocusVisualStyle" Value="{StaticResource FocusRing}"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="CheckBox">
            <Grid Background="Transparent">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
              </Grid.ColumnDefinitions>
              <Grid Width="36" Height="20" VerticalAlignment="Top" Margin="0,1,0,0">
                <Border x:Name="track" CornerRadius="10" Background="{DynamicResource Secondary}" BorderBrush="{DynamicResource Border}" BorderThickness="1"/>
                <Ellipse x:Name="thumb" Width="14" Height="14" Fill="{DynamicResource Surface}" Stroke="{DynamicResource Border}"
                         HorizontalAlignment="Left" Margin="3,0,0,0"/>
              </Grid>
              <ContentPresenter Grid.Column="1" Margin="10,0,0,0" VerticalAlignment="Center" RecognizesAccessKey="True"/>
            </Grid>
            <ControlTemplate.Triggers>
              <Trigger Property="IsChecked" Value="True">
                <Setter TargetName="track" Property="Background" Value="{DynamicResource SwitchOn}"/>
                <Setter TargetName="track" Property="BorderBrush" Value="{DynamicResource SwitchOn}"/>
                <Setter TargetName="thumb" Property="Fill" Value="#FFFFFF"/>
                <Setter TargetName="thumb" Property="Stroke" Value="#FFFFFF"/>
                <Setter TargetName="thumb" Property="HorizontalAlignment" Value="Right"/>
                <Setter TargetName="thumb" Property="Margin" Value="0,0,3,0"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style TargetType="ToolTip">
      <Setter Property="Background" Value="{DynamicResource Surface}"/>
      <Setter Property="Foreground" Value="{DynamicResource Text}"/>
      <Setter Property="BorderBrush" Value="{DynamicResource Border}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding" Value="9,6"/>
      <Setter Property="MaxWidth" Value="380"/>
      <Setter Property="ContentTemplate">
        <Setter.Value>
          <DataTemplate><TextBlock Text="{Binding}" TextWrapping="Wrap"/></DataTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="IconButton" TargetType="Button" BasedOn="{StaticResource BtnBase}">
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="Foreground" Value="{DynamicResource Muted}"/>
      <Setter Property="FontFamily" Value="Segoe MDL2 Assets"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="FontWeight" Value="Normal"/>
      <Setter Property="Padding" Value="7,4"/>
      <Setter Property="MinHeight" Value="26"/>
      <Setter Property="Margin" Value="8,2,0,2"/>
      <Setter Property="VerticalAlignment" Value="Top"/>
      <Setter Property="Content" Value="&#xE8C8;"/>
    </Style>

    <Style x:Key="ComboToggle" TargetType="ToggleButton">
      <Setter Property="Focusable" Value="False"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ToggleButton">
            <Border x:Name="bd" Background="{DynamicResource InputBg}" BorderBrush="{DynamicResource Border}" BorderThickness="1" CornerRadius="6">
              <TextBlock FontFamily="Segoe MDL2 Assets" Text="&#xE70D;" FontSize="10" Foreground="{DynamicResource Muted}"
                         HorizontalAlignment="Right" VerticalAlignment="Center" Margin="0,0,10,0"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="bd" Property="BorderBrush" Value="{DynamicResource Primary}"/></Trigger>
              <Trigger Property="IsChecked" Value="True"><Setter TargetName="bd" Property="BorderBrush" Value="{DynamicResource Primary}"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="ComboBox">
      <Setter Property="Foreground" Value="{DynamicResource Text}"/>
      <Setter Property="MinHeight" Value="32"/>
      <Setter Property="MinWidth" Value="150"/>
      <Setter Property="Padding" Value="10,0,30,0"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="FocusVisualStyle" Value="{StaticResource FocusRing}"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ComboBox">
            <Grid>
              <ToggleButton Style="{StaticResource ComboToggle}" ClickMode="Press"
                            IsChecked="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}"/>
              <ContentPresenter IsHitTestVisible="False" Margin="{TemplateBinding Padding}" VerticalAlignment="Center"
                                Content="{TemplateBinding SelectionBoxItem}" ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}"/>
              <Popup x:Name="PART_Popup" Placement="Bottom" AllowsTransparency="True" Focusable="False" PopupAnimation="Fade"
                     IsOpen="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}">
                <Border Background="{DynamicResource Surface}" BorderBrush="{DynamicResource Border}" BorderThickness="1" CornerRadius="6"
                        Margin="0,4,0,0" Padding="4" MinWidth="{Binding ActualWidth, RelativeSource={RelativeSource TemplatedParent}}">
                  <ScrollViewer MaxHeight="320" VerticalScrollBarVisibility="Auto">
                    <ItemsPresenter/>
                  </ScrollViewer>
                </Border>
              </Popup>
            </Grid>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="ComboBoxItem">
      <Setter Property="Foreground" Value="{DynamicResource Text}"/>
      <Setter Property="Padding" Value="10,6"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ComboBoxItem">
            <Border x:Name="bd" Background="Transparent" CornerRadius="4" Padding="{TemplateBinding Padding}">
              <ContentPresenter/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsHighlighted" Value="True"><Setter TargetName="bd" Property="Background" Value="{DynamicResource Selection}"/></Trigger>
              <Trigger Property="IsSelected" Value="True"><Setter Property="FontWeight" Value="SemiBold"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="Card" TargetType="Border">
      <Setter Property="Background" Value="{DynamicResource Surface}"/>
      <Setter Property="BorderBrush" Value="{DynamicResource Border}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="CornerRadius" Value="10"/>
      <Setter Property="Padding" Value="18"/>
    </Style>

    <Style x:Key="H2" TargetType="TextBlock">
      <Setter Property="FontSize" Value="15"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
    </Style>
    <Style x:Key="Caption" TargetType="TextBlock">
      <Setter Property="Foreground" Value="{DynamicResource Muted}"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="TextWrapping" Value="Wrap"/>
    </Style>
    <Style x:Key="FieldLabel" TargetType="TextBlock">
      <Setter Property="Foreground" Value="{DynamicResource Muted}"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Margin" Value="0,0,0,6"/>
    </Style>
    <Style x:Key="FieldName" TargetType="TextBlock">
      <Setter Property="Foreground" Value="{DynamicResource Muted}"/>
      <Setter Property="Margin" Value="0,6,12,6"/>
    </Style>
    <Style x:Key="FieldValue" TargetType="TextBlock">
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Margin" Value="0,6,0,6"/>
      <Setter Property="TextWrapping" Value="Wrap"/>
    </Style>
    <Style x:Key="Glyph" TargetType="TextBlock">
      <Setter Property="FontFamily" Value="Segoe MDL2 Assets"/>
      <Setter Property="FontSize" Value="14"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="Margin" Value="0,0,8,0"/>
    </Style>
    <Style x:Key="StatValue" TargetType="TextBlock">
      <Setter Property="FontSize" Value="22"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
    </Style>
    <Style x:Key="CellText" TargetType="TextBlock">
      <Setter Property="TextTrimming" Value="CharacterEllipsis"/>
    </Style>
    <Style x:Key="TypeText" TargetType="TextBlock">
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="TextTrimming" Value="CharacterEllipsis"/>
      <Setter Property="ToolTip" Value="{Binding AccountType}"/>
      <Style.Triggers>
        <DataTrigger Binding="{Binding Category}" Value="Entra"><Setter Property="Foreground" Value="{DynamicResource CatEntra}"/></DataTrigger>
        <DataTrigger Binding="{Binding Category}" Value="MSA"><Setter Property="Foreground" Value="{DynamicResource CatMSA}"/></DataTrigger>
        <DataTrigger Binding="{Binding Category}" Value="Federated"><Setter Property="Foreground" Value="{DynamicResource CatFederated}"/></DataTrigger>
        <DataTrigger Binding="{Binding Category}" Value="Dual"><Setter Property="Foreground" Value="{DynamicResource CatDual}"/></DataTrigger>
        <DataTrigger Binding="{Binding Category}" Value="GCCHigh"><Setter Property="Foreground" Value="{DynamicResource CatGCCHigh}"/></DataTrigger>
        <DataTrigger Binding="{Binding Category}" Value="China"><Setter Property="Foreground" Value="{DynamicResource CatChina}"/></DataTrigger>
        <DataTrigger Binding="{Binding Category}" Value="NotMicrosoft"><Setter Property="Foreground" Value="{DynamicResource CatNotMicrosoft}"/></DataTrigger>
        <DataTrigger Binding="{Binding Category}" Value="Error"><Setter Property="Foreground" Value="{DynamicResource CatError}"/></DataTrigger>
        <DataTrigger Binding="{Binding Category}" Value="Unknown"><Setter Property="Foreground" Value="{DynamicResource CatUnknown}"/></DataTrigger>
      </Style.Triggers>
    </Style>

    <Style TargetType="TabControl">
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Padding" Value="0"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TabControl">
            <Grid>
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
              </Grid.RowDefinitions>
              <Border Grid.Row="0" Background="{DynamicResource Surface}" BorderBrush="{DynamicResource Border}" BorderThickness="0,0,0,1" Padding="16,0">
                <TabPanel IsItemsHost="True" Background="Transparent"/>
              </Border>
              <ContentPresenter Grid.Row="1" ContentSource="SelectedContent"/>
            </Grid>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="TabItem">
      <Setter Property="Foreground" Value="{DynamicResource Muted}"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="FocusVisualStyle" Value="{StaticResource FocusRing}"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TabItem">
            <Border x:Name="bd" Background="Transparent" Padding="14,11,14,9" BorderThickness="0,0,0,3" BorderBrush="Transparent">
              <ContentPresenter ContentSource="Header" TextElement.FontWeight="SemiBold"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter Property="Foreground" Value="{DynamicResource Text}"/></Trigger>
              <Trigger Property="IsSelected" Value="True">
                <Setter TargetName="bd" Property="BorderBrush" Value="{DynamicResource Primary}"/>
                <Setter Property="Foreground" Value="{DynamicResource Text}"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="GridHeader" TargetType="DataGridColumnHeader">
      <Setter Property="Background" Value="{DynamicResource GridBg}"/>
      <Setter Property="Foreground" Value="{DynamicResource Muted}"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="BorderBrush" Value="{DynamicResource Border}"/>
      <Setter Property="BorderThickness" Value="0,0,0,1"/>
      <Setter Property="Padding" Value="10,0"/>
      <Setter Property="Height" Value="34"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="DataGridColumnHeader">
            <Grid>
              <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}"
                      BorderThickness="{TemplateBinding BorderThickness}" Padding="{TemplateBinding Padding}">
                <ContentPresenter VerticalAlignment="Center"/>
              </Border>
              <Thumb x:Name="PART_RightHeaderGripper" HorizontalAlignment="Right" Width="6" Cursor="SizeWE" Opacity="0"/>
            </Grid>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="GridCell" TargetType="DataGridCell">
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="DataGridCell">
            <Border Background="{TemplateBinding Background}" Padding="10,0">
              <ContentPresenter VerticalAlignment="Center"/>
            </Border>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
      <Style.Triggers>
        <Trigger Property="IsSelected" Value="True">
          <Setter Property="Background" Value="{DynamicResource Selection}"/>
          <Setter Property="Foreground" Value="{DynamicResource Text}"/>
        </Trigger>
      </Style.Triggers>
    </Style>
    <Style TargetType="DataGrid">
      <Setter Property="Background" Value="{DynamicResource Surface}"/>
      <Setter Property="Foreground" Value="{DynamicResource Text}"/>
      <Setter Property="BorderBrush" Value="{DynamicResource Border}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="RowBackground" Value="{DynamicResource Surface}"/>
      <Setter Property="AlternatingRowBackground" Value="{DynamicResource GridAlt}"/>
      <Setter Property="HorizontalGridLinesBrush" Value="{DynamicResource Border}"/>
      <Setter Property="GridLinesVisibility" Value="Horizontal"/>
      <Setter Property="HeadersVisibility" Value="Column"/>
      <Setter Property="AutoGenerateColumns" Value="False"/>
      <Setter Property="IsReadOnly" Value="True"/>
      <Setter Property="CanUserAddRows" Value="False"/>
      <Setter Property="CanUserDeleteRows" Value="False"/>
      <Setter Property="SelectionMode" Value="Extended"/>
      <Setter Property="RowHeight" Value="32"/>
      <Setter Property="ClipboardCopyMode" Value="IncludeHeader"/>
      <Setter Property="ColumnHeaderStyle" Value="{StaticResource GridHeader}"/>
      <Setter Property="CellStyle" Value="{StaticResource GridCell}"/>
    </Style>

    <Style TargetType="ProgressBar">
      <Setter Property="Height" Value="6"/>
      <Setter Property="Foreground" Value="{DynamicResource Primary}"/>
      <Setter Property="Background" Value="{DynamicResource Secondary}"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ProgressBar">
            <Grid>
              <Border x:Name="PART_Track" CornerRadius="3" Background="{TemplateBinding Background}"/>
              <Border x:Name="PART_Indicator" CornerRadius="3" Background="{TemplateBinding Foreground}" HorizontalAlignment="Left"/>
            </Grid>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
  </Window.Resources>

  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <!-- Header with the three switches -->
    <Border Grid.Row="0" Background="{DynamicResource HeaderBg}" Padding="20,12">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
          <Border Width="38" Height="38" CornerRadius="9" Background="{DynamicResource HeaderButtonBg}" Margin="0,0,12,0">
            <TextBlock FontFamily="Segoe MDL2 Assets" Text="&#xE77B;" FontSize="18" Foreground="{DynamicResource HeaderText}"
                       HorizontalAlignment="Center" VerticalAlignment="Center"/>
          </Border>
          <StackPanel VerticalAlignment="Center">
            <TextBlock Text="Microsoft Account Checker" FontSize="17" FontWeight="SemiBold" Foreground="{DynamicResource HeaderText}"/>
            <TextBlock x:Name="HeaderSubtitle" Text="Commercial, GCC High and China clouds" FontSize="12" Foreground="{DynamicResource HeaderMuted}"/>
          </StackPanel>
        </StackPanel>
        <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
          <CheckBox x:Name="DarkToggle" Style="{StaticResource Switch}" Content="Dark mode" Margin="0,0,22,0"/>
          <CheckBox x:Name="ApiToggle" Style="{StaticResource Switch}" Content="API server" Margin="0,0,22,0"/>
          <CheckBox x:Name="LogToggle" Style="{StaticResource Switch}" Content="Logging" Margin="0,0,22,0"/>
          <Button x:Name="SettingsButton" Style="{StaticResource HeaderButton}" Content="Settings" Margin="0,0,8,0" ToolTip="Open settings (Ctrl+,)" AutomationProperties.HelpText="Open settings (Ctrl+,)"/>
          <Button x:Name="AboutButton" Style="{StaticResource HeaderButton}" Content="About" ToolTip="About this app (F1)" AutomationProperties.HelpText="About this app (F1)"/>
        </StackPanel>
      </Grid>
    </Border>

    <TabControl x:Name="MainTabs" Grid.Row="1">

      <!-- Single check -->
      <TabItem x:Name="TabSingle">
        <TabItem.Header>
          <StackPanel Orientation="Horizontal">
            <TextBlock Style="{StaticResource Glyph}" Text="&#xE721;"/>
            <TextBlock Text="Single check"/>
          </StackPanel>
        </TabItem.Header>
        <Grid Margin="20">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
          </Grid.RowDefinitions>

          <Border Style="{StaticResource Card}">
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
              </Grid.RowDefinitions>
              <TextBlock Grid.ColumnSpan="2" Style="{StaticResource FieldLabel}" Text="Email address"/>
              <Grid Grid.Row="1">
                <TextBox x:Name="EmailBox" MinHeight="40" FontSize="14" AutomationProperties.Name="Email address" ToolTip="The address to check. Press Enter to check it; Ctrl+L comes back here from anywhere." AutomationProperties.HelpText="The address to check. Press Enter to check it; Ctrl+L comes back here from anywhere."/>
                <TextBlock x:Name="EmailHint" Text="name@company.com" FontSize="14" Foreground="{DynamicResource Muted}"
                           Margin="11,0,0,0" VerticalAlignment="Center" IsHitTestVisible="False"/>
              </Grid>
              <StackPanel Grid.Row="1" Grid.Column="1" Orientation="Horizontal" Margin="12,0,0,0">
                <Button x:Name="CheckButton" Style="{StaticResource PrimaryButton}" Content="Check" MinWidth="110" MinHeight="40" ToolTip="Check the address (Enter)" AutomationProperties.HelpText="Check the address (Enter)"/>
                <Button x:Name="PasteButton" Content="Paste" MinWidth="80" MinHeight="40" Margin="8,0,0,0" ToolTip="Paste an address from the clipboard" AutomationProperties.HelpText="Paste an address from the clipboard"/>
                <Button x:Name="ClearButton" Content="Clear" MinWidth="80" MinHeight="40" Margin="8,0,0,0" ToolTip="Clear the address and the result" AutomationProperties.HelpText="Clear the address and the result"/>
              </StackPanel>
              <TextBlock x:Name="EmailError" Grid.Row="2" Grid.ColumnSpan="2" Margin="2,8,0,0" TextWrapping="Wrap"
                         Foreground="{DynamicResource CatNotMicrosoft}" Visibility="Collapsed"/>
            </Grid>
          </Border>

          <Border Grid.Row="1" Style="{StaticResource Card}" Margin="0,16,0,0">
            <Grid>
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
              </Grid.RowDefinitions>
              <DockPanel LastChildFill="False">
                <StackPanel DockPanel.Dock="Right" Orientation="Horizontal">
                  <Button x:Name="RecheckButton" Content="Check again live" Visibility="Collapsed"
                          ToolTip="This result came from the 24-hour cache. Ask Microsoft again." AutomationProperties.HelpText="This result came from the 24-hour cache. Ask Microsoft again."/>
                  <Button x:Name="RawToggleButton" Content="Show raw response" Margin="8,0,0,0" IsEnabled="False"
                          ToolTip="Show exactly what Microsoft's endpoints returned" AutomationProperties.HelpText="Show exactly what Microsoft's endpoints returned"/>
                  <Button x:Name="CopyButton" Style="{StaticResource PrimaryButton}" Content="Copy result" Margin="8,0,0,0" IsEnabled="False"
                          ToolTip="Copy the four-line result (Ctrl+Shift+C)" AutomationProperties.HelpText="Copy the four-line result (Ctrl+Shift+C)"/>
                </StackPanel>
                <TextBlock DockPanel.Dock="Left" Style="{StaticResource H2}" Text="Result"/>
                <Border x:Name="ResultBadge" DockPanel.Dock="Left" CornerRadius="12" Padding="10,3" Margin="12,0,0,0"
                        VerticalAlignment="Center" Visibility="Collapsed" ToolTip="Account type" AutomationProperties.HelpText="Account type">
                  <TextBlock x:Name="ResultBadgeText" FontSize="12" FontWeight="SemiBold"/>
                </Border>
                <Border x:Name="StatusBadge" DockPanel.Dock="Left" CornerRadius="12" Padding="10,3" Margin="8,0,0,0"
                        VerticalAlignment="Center" Visibility="Collapsed" ToolTip="How certain the answer is" AutomationProperties.HelpText="How certain the answer is">
                  <TextBlock x:Name="StatusBadgeText" FontSize="12" FontWeight="SemiBold"/>
                </Border>
              </DockPanel>

              <ScrollViewer Grid.Row="1" Margin="0,14,0,0" VerticalScrollBarVisibility="Auto">
                <StackPanel>
                  <TextBlock x:Name="ResultEmpty" Style="{StaticResource Caption}" FontSize="13"
                             Text="Enter an email address and select Check, or press Enter."/>

                  <Border x:Name="ResultCallout" CornerRadius="8" Padding="14,10" Margin="0,0,0,12" Visibility="Collapsed">
                    <Grid>
                      <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="*"/>
                      </Grid.ColumnDefinitions>
                      <TextBlock x:Name="ResultCalloutIcon" FontFamily="Segoe MDL2 Assets" Text="&#xE7BA;" FontSize="16" Margin="0,1,12,0" VerticalAlignment="Top"/>
                      <StackPanel Grid.Column="1">
                        <TextBlock x:Name="ResultCalloutTitle" FontWeight="SemiBold"/>
                        <TextBlock x:Name="ResultCalloutText" TextWrapping="Wrap" Margin="0,3,0,0"/>
                      </StackPanel>
                    </Grid>
                  </Border>

                  <Border x:Name="ChangeCallout" CornerRadius="8" Padding="14,10" Margin="0,0,0,12" Visibility="Collapsed" Background="{DynamicResource CatDualBg}">
                    <Grid>
                      <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="*"/>
                      </Grid.ColumnDefinitions>
                      <TextBlock FontFamily="Segoe MDL2 Assets" Text="&#xE81C;" FontSize="16" Margin="0,1,12,0" VerticalAlignment="Top" Foreground="{DynamicResource CatDual}"/>
                      <StackPanel Grid.Column="1">
                        <TextBlock x:Name="ChangeCalloutTitle" FontWeight="SemiBold" Foreground="{DynamicResource CatDual}"/>
                        <TextBlock x:Name="ChangeCalloutText" TextWrapping="Wrap" Margin="0,3,0,0"/>
                      </StackPanel>
                    </Grid>
                  </Border>

                  <Grid x:Name="ResultFields" Visibility="Collapsed">
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="170"/>
                      <ColumnDefinition Width="*"/>
                      <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <Grid.RowDefinitions>
                      <RowDefinition Height="Auto"/>
                      <RowDefinition Height="Auto"/>
                      <RowDefinition Height="Auto"/>
                      <RowDefinition Height="Auto"/>
                      <RowDefinition Height="Auto"/>
                      <RowDefinition Height="Auto"/>
                      <RowDefinition Height="Auto"/>
                      <RowDefinition Height="Auto"/>
                      <RowDefinition Height="Auto"/>
                      <RowDefinition Height="Auto"/>
                      <RowDefinition Height="Auto"/>
                      <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    <TextBlock Grid.Row="0" Style="{StaticResource FieldName}" Text="Username"/>
                    <TextBlock Grid.Row="0" Grid.Column="1" x:Name="ValUsername" Style="{StaticResource FieldValue}"/>
                    <TextBlock Grid.Row="1" Style="{StaticResource FieldName}" Text="FoundInCloud"
                               ToolTip="The Microsoft cloud where the account was found: Commercial, GCC High (US Government) or China. N/A when no account was found." AutomationProperties.HelpText="The Microsoft cloud where the account was found: Commercial, GCC High (US Government) or China. N/A when no account was found."/>
                    <TextBlock Grid.Row="1" Grid.Column="1" x:Name="ValCloud" Style="{StaticResource FieldValue}"/>
                    <TextBlock Grid.Row="2" Style="{StaticResource FieldName}" Text="Cloud URL" ToolTip="The sign-in address for that cloud." AutomationProperties.HelpText="The sign-in address for that cloud."/>
                    <TextBlock Grid.Row="2" Grid.Column="1" Style="{StaticResource FieldValue}">
                      <Hyperlink x:Name="CloudUrlLink" Foreground="{DynamicResource Link}"><Run x:Name="CloudUrlRun"/></Hyperlink>
                    </TextBlock>
                    <Button Grid.Row="2" Grid.Column="2" x:Name="CopyCloudUrlButton" Style="{StaticResource IconButton}" ToolTip="Copy the Cloud URL" AutomationProperties.HelpText="Copy the Cloud URL" AutomationProperties.Name="Copy the Cloud URL"/>
                    <TextBlock Grid.Row="3" x:Name="FedLabel" Style="{StaticResource FieldName}" Text="FederationRedirectUrl" Visibility="Collapsed"
                               ToolTip="Where Microsoft sends this user to sign in, because the organization uses its own identity provider." AutomationProperties.HelpText="Where Microsoft sends this user to sign in, because the organization uses its own identity provider."/>
                    <TextBlock Grid.Row="3" Grid.Column="1" x:Name="FedValue" Style="{StaticResource FieldValue}" Visibility="Collapsed">
                      <Hyperlink x:Name="FedUrlLink" Foreground="{DynamicResource Link}"><Run x:Name="FedUrlRun"/></Hyperlink>
                    </TextBlock>
                    <Button Grid.Row="3" Grid.Column="2" x:Name="CopyFedButton" Style="{StaticResource IconButton}" ToolTip="Copy the FederationRedirectUrl" AutomationProperties.HelpText="Copy the FederationRedirectUrl" AutomationProperties.Name="Copy the FederationRedirectUrl" Visibility="Collapsed"/>
                    <TextBlock Grid.Row="4" Style="{StaticResource FieldName}" Text="Account Type" ToolTip="What kind of Microsoft account this address belongs to." AutomationProperties.HelpText="What kind of Microsoft account this address belongs to."/>
                    <TextBlock Grid.Row="4" Grid.Column="1" x:Name="ValAccountType" Style="{StaticResource FieldValue}"/>

                    <Border Grid.Row="5" Grid.ColumnSpan="3" x:Name="DetailDivider" Height="1" Margin="0,10,0,6"
                            Background="{DynamicResource Border}" Visibility="Collapsed"/>
                    <TextBlock Grid.Row="6" x:Name="LblOrg" Style="{StaticResource FieldName}" Text="Organization" Visibility="Collapsed"
                               ToolTip="The organization's display name, from Microsoft's GetUserRealm endpoint." AutomationProperties.HelpText="The organization's display name, from Microsoft's GetUserRealm endpoint."/>
                    <TextBox Grid.Row="6" Grid.Column="1" x:Name="ValOrg" Style="{StaticResource ValueBox}" Visibility="Collapsed"/>
                    <Button Grid.Row="6" Grid.Column="2" x:Name="CopyOrgButton" Style="{StaticResource IconButton}" ToolTip="Copy the organization name" AutomationProperties.HelpText="Copy the organization name" AutomationProperties.Name="Copy the organization name" Visibility="Collapsed"/>
                    <TextBlock Grid.Row="7" x:Name="LblProvider" Style="{StaticResource FieldName}" Text="Federation provider" Visibility="Collapsed"
                               ToolTip="The identity provider recognized from the federation sign-in address, such as Okta, GoDaddy or AD FS." AutomationProperties.HelpText="The identity provider recognized from the federation sign-in address, such as Okta, GoDaddy or AD FS."/>
                    <TextBox Grid.Row="7" Grid.Column="1" x:Name="ValProvider" Style="{StaticResource ValueBox}" Visibility="Collapsed"/>
                    <TextBlock Grid.Row="8" x:Name="LblTenantId" Style="{StaticResource FieldName}" Text="Tenant ID" Visibility="Collapsed"
                               ToolTip="The unique ID (GUID) of the organization's Microsoft Entra tenant." AutomationProperties.HelpText="The unique ID (GUID) of the organization's Microsoft Entra tenant."/>
                    <TextBox Grid.Row="8" Grid.Column="1" x:Name="ValTenantId" Style="{StaticResource ValueBox}" FontFamily="Consolas" Visibility="Collapsed"/>
                    <Button Grid.Row="8" Grid.Column="2" x:Name="CopyTenantButton" Style="{StaticResource IconButton}" ToolTip="Copy the tenant ID" AutomationProperties.HelpText="Copy the tenant ID" AutomationProperties.Name="Copy the tenant ID" Visibility="Collapsed"/>
                    <TextBlock Grid.Row="9" x:Name="LblRegion" Style="{StaticResource FieldName}" Text="Tenant region" Visibility="Collapsed"
                               ToolTip="Where the tenant is located, from its OpenID discovery document. A sub-scope such as GCC is shown after the region." AutomationProperties.HelpText="Where the tenant is located, from its OpenID discovery document. A sub-scope such as GCC is shown after the region."/>
                    <TextBox Grid.Row="9" Grid.Column="1" x:Name="ValRegion" Style="{StaticResource ValueBox}" Visibility="Collapsed"/>
                    <TextBlock Grid.Row="10" x:Name="LblNamespace" Style="{StaticResource FieldName}" Text="Domain type" Visibility="Collapsed"
                               ToolTip="Managed: Microsoft handles sign-in. Federated: another identity provider handles sign-in." AutomationProperties.HelpText="Managed: Microsoft handles sign-in. Federated: another identity provider handles sign-in."/>
                    <TextBox Grid.Row="10" Grid.Column="1" x:Name="ValNamespace" Style="{StaticResource ValueBox}" Visibility="Collapsed"/>
                    <TextBlock Grid.Row="11" Grid.ColumnSpan="3" x:Name="ResultMeta" Style="{StaticResource Caption}" Margin="0,8,0,0"/>
                  </Grid>
                  <TextBlock x:Name="ResultHint" Style="{StaticResource Caption}" Margin="0,10,0,0" Visibility="Collapsed"/>
                  <TextBox x:Name="RawBox" Style="{StaticResource MonoBox}" Height="280" Margin="0,12,0,0" Visibility="Collapsed"/>
                </StackPanel>
              </ScrollViewer>
            </Grid>
          </Border>
        </Grid>
      </TabItem>

      <!-- Batch check -->
      <TabItem x:Name="TabBatch">
        <TabItem.Header>
          <StackPanel Orientation="Horizontal">
            <TextBlock Style="{StaticResource Glyph}" Text="&#xE8FD;"/>
            <TextBlock Text="Batch check"/>
          </StackPanel>
        </TabItem.Header>
        <Grid Margin="20">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="320"/>
            <ColumnDefinition Width="*"/>
          </Grid.ColumnDefinitions>

          <Border Style="{StaticResource Card}" Margin="0,0,16,0">
            <Grid>
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
              </Grid.RowDefinitions>
              <TextBlock Style="{StaticResource H2}" Text="Email list"/>
              <TextBlock Grid.Row="1" Style="{StaticResource Caption}" Margin="0,4,0,10"
                         Text="One address per line. Commas, semicolons and spaces also work. Import reads .txt and .csv files. Duplicates are skipped."/>
              <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,0,0,10">
                <Button x:Name="BatchImportButton" Content="Import file" ToolTip="Read addresses from a .txt or .csv file" AutomationProperties.HelpText="Read addresses from a .txt or .csv file"/>
                <Button x:Name="BatchPasteButton" Content="Paste" Margin="8,0,0,0" ToolTip="Add addresses from the clipboard" AutomationProperties.HelpText="Add addresses from the clipboard"/>
                <Button x:Name="BatchClearInputButton" Content="Clear" Margin="8,0,0,0" ToolTip="Empty the list" AutomationProperties.HelpText="Empty the list"/>
              </StackPanel>
              <TextBox x:Name="BatchInput" Grid.Row="3" AutomationProperties.Name="Email list" AcceptsReturn="True" TextWrapping="NoWrap"
                       VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"
                       VerticalContentAlignment="Stretch" FontFamily="Consolas" FontSize="12"
                       ToolTip="Press Ctrl+Enter to start the batch" AutomationProperties.HelpText="Press Ctrl+Enter to start the batch"/>
              <TextBlock x:Name="BatchCount" Grid.Row="4" Style="{StaticResource Caption}" Margin="0,8,0,0" Text="No addresses yet"/>
            </Grid>
          </Border>

          <Border Grid.Column="1" Style="{StaticResource Card}">
            <Grid>
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
              </Grid.RowDefinitions>
              <DockPanel LastChildFill="False">
                <TextBlock DockPanel.Dock="Left" Style="{StaticResource H2}" Text="Results"/>
                <TextBlock x:Name="BatchProgressText" DockPanel.Dock="Right" Style="{StaticResource Caption}" VerticalAlignment="Center"/>
              </DockPanel>
              <WrapPanel Grid.Row="1" Margin="0,10,0,0">
                <Button x:Name="BatchStartButton" Style="{StaticResource PrimaryButton}" Content="Start" MinWidth="90" Margin="0,0,8,8" ToolTip="Check every address in the list (Ctrl+Enter)" AutomationProperties.HelpText="Check every address in the list (Ctrl+Enter)"/>
                <Button x:Name="BatchPauseButton" Content="Pause" MinWidth="80" Margin="0,0,8,8" IsEnabled="False" ToolTip="Pause after the current lookup; select again to resume" AutomationProperties.HelpText="Pause after the current lookup; select again to resume"/>
                <Button x:Name="BatchStopButton" Style="{StaticResource DangerButton}" Content="Stop" MinWidth="80" Margin="0,0,8,8" IsEnabled="False" ToolTip="Stop after the current lookup" AutomationProperties.HelpText="Stop after the current lookup"/>
                <Button x:Name="BatchRetryButton" Content="Retry failed" Margin="0,0,16,8" IsEnabled="False"
                        ToolTip="Check again only the results that were Throttled or Temporarily unavailable" AutomationProperties.HelpText="Check again only the results that were Throttled or Temporarily unavailable"/>
                <Button x:Name="BatchExportCsvButton" Content="Export CSV" Margin="0,0,8,8" ToolTip="Save all results as CSV (Ctrl+E)" AutomationProperties.HelpText="Save all results as CSV (Ctrl+E)"/>
                <Button x:Name="BatchExportJsonButton" Content="Export JSON" Margin="0,0,8,8"/>
                <Button x:Name="BatchExportHtmlButton" Content="HTML report" Margin="0,0,8,8" ToolTip="Executive summary with totals, breakdowns, exceptions and all results" AutomationProperties.HelpText="Executive summary with totals, breakdowns, exceptions and all results"/>
                <Button x:Name="BatchExportPdfButton" Content="PDF report" Margin="0,0,8,8" ToolTip="The same report as a PDF (uses Microsoft Edge)" AutomationProperties.HelpText="The same report as a PDF (uses Microsoft Edge)"/>
                <Button x:Name="BatchClearResultsButton" Content="Clear results" Margin="0,0,8,8"/>
              </WrapPanel>
              <WrapPanel Grid.Row="2" Margin="0,2,0,8">
                <CheckBox x:Name="BatchSkipCacheCheck" Style="{StaticResource BodySwitch}" Margin="0,0,20,4"
                          ToolTip="Ask Microsoft for every address, even ones checked in the last 24 hours" AutomationProperties.HelpText="Ask Microsoft for every address, even ones checked in the last 24 hours">
                  <TextBlock Text="Skip cached results"/>
                </CheckBox>
                <CheckBox x:Name="RefreshStaleCheck" Style="{StaticResource BodySwitch}" Margin="0,0,20,4"
                          ToolTip="Reuse recent answers, but ask Microsoft again for any older than the Stale after setting (Settings tab)" AutomationProperties.HelpText="Reuse recent answers, but ask Microsoft again for any older than the Stale after setting (Settings tab)">
                  <TextBlock Text="Refresh stale only"/>
                </CheckBox>
                <CheckBox x:Name="GroupByDomainCheck" Style="{StaticResource BodySwitch}" Margin="0,0,20,4"
                          ToolTip="Show one row per domain with its organization, tenant, cloud and federation provider (Ctrl+G)" AutomationProperties.HelpText="Show one row per domain with its organization, tenant, cloud and federation provider (Ctrl+G)">
                  <TextBlock Text="Group by domain"/>
                </CheckBox>
                <CheckBox x:Name="ClipboardWatchCheck" Style="{StaticResource BodySwitch}" Margin="0,0,20,4"
                          ToolTip="While this is on, addresses you copy anywhere are added to the email list" AutomationProperties.HelpText="While this is on, addresses you copy anywhere are added to the email list">
                  <TextBlock Text="Collect from clipboard"/>
                </CheckBox>
                <TextBlock x:Name="BatchSettingsText" Style="{StaticResource Caption}" VerticalAlignment="Center" Margin="0,0,0,4"/>
              </WrapPanel>
              <ProgressBar x:Name="BatchProgress" Grid.Row="3" Margin="0,2,0,10" Minimum="0" Maximum="1" Value="0"/>
              <WrapPanel x:Name="BatchChips" Grid.Row="4"/>
              <DataGrid x:Name="BatchGrid" Grid.Row="5" Margin="0,6,0,0" AutomationProperties.Name="Batch results">
                <DataGrid.Columns>
                  <DataGridTextColumn Header="#" Binding="{Binding Index}" Width="50"/>
                  <DataGridTextColumn Header="Username" Binding="{Binding Username}" Width="2*" MinWidth="170" ElementStyle="{StaticResource CellText}"/>
                  <DataGridTextColumn Header="Cloud" Binding="{Binding FoundInCloud}" Width="90"/>
                  <DataGridTextColumn Header="Account type" Binding="{Binding AccountType}" Width="2.4*" MinWidth="210" ElementStyle="{StaticResource TypeText}"/>
                  <DataGridTextColumn Header="Status" Binding="{Binding StatusLabel}" Width="1*" MinWidth="110" ElementStyle="{StaticResource StatusText}"/>
                  <DataGridTextColumn Header="Changes" Binding="{Binding Changes}" Width="1.5*" MinWidth="120" ElementStyle="{StaticResource ChangeText}"/>
                  <DataGridTextColumn Header="Organization" Binding="{Binding Organization}" Width="1.4*" MinWidth="120" ElementStyle="{StaticResource CellText}"/>
                  <DataGridTextColumn Header="Federation provider" Binding="{Binding FederationProvider}" Width="1.4*" MinWidth="120" ElementStyle="{StaticResource CellText}"/>
                  <DataGridTextColumn Header="Tenant region" Binding="{Binding TenantRegion}" Width="1.2*" MinWidth="110" ElementStyle="{StaticResource CellText}"/>
                  <DataGridTextColumn Header="Tenant ID" Binding="{Binding TenantId}" Width="1.8*" MinWidth="140" ElementStyle="{StaticResource CellText}"/>
                  <DataGridTextColumn Header="Source" Binding="{Binding Source}" Width="65"/>
                  <DataGridTextColumn Header="Time" Binding="{Binding Timestamp}" Width="135"/>
                </DataGrid.Columns>
              </DataGrid>
              <DataGrid x:Name="DomainGrid" Grid.Row="5" Margin="0,6,0,0" Visibility="Collapsed" AutomationProperties.Name="Results by domain">
                <DataGrid.Columns>
                  <DataGridTextColumn Header="Domain" Binding="{Binding Domain}" Width="1.6*" MinWidth="150" ElementStyle="{StaticResource CellText}"/>
                  <DataGridTextColumn Header="Addresses" Binding="{Binding Addresses}" Width="85"/>
                  <DataGridTextColumn Header="Organization" Binding="{Binding Organization}" Width="1.4*" MinWidth="120" ElementStyle="{StaticResource CellText}"/>
                  <DataGridTextColumn Header="Tenant ID" Binding="{Binding TenantId}" Width="1.8*" MinWidth="140" ElementStyle="{StaticResource CellText}"/>
                  <DataGridTextColumn Header="Cloud" Binding="{Binding Cloud}" Width="100"/>
                  <DataGridTextColumn Header="Federation provider" Binding="{Binding FederationProvider}" Width="1.4*" MinWidth="120" ElementStyle="{StaticResource CellText}"/>
                  <DataGridTextColumn Header="Account types" Binding="{Binding AccountTypes}" Width="2*" MinWidth="160" ElementStyle="{StaticResource CellText}"/>
                  <DataGridTextColumn Header="Not found" Binding="{Binding NotFound}" Width="85"/>
                  <DataGridTextColumn Header="Need attention" Binding="{Binding NeedsAttention}" Width="110"/>
                  <DataGridTextColumn Header="Changed" Binding="{Binding Changed}" Width="80"/>
                </DataGrid.Columns>
              </DataGrid>
            </Grid>
          </Border>
        </Grid>
      </TabItem>

      <!-- History -->
      <TabItem x:Name="TabHistory">
        <TabItem.Header>
          <StackPanel Orientation="Horizontal">
            <TextBlock Style="{StaticResource Glyph}" Text="&#xE81C;"/>
            <TextBlock Text="History"/>
          </StackPanel>
        </TabItem.Header>
        <Grid Margin="20">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
          </Grid.RowDefinitions>
          <UniformGrid Columns="8">
            <Border Style="{StaticResource Card}" Padding="14,12" Margin="0,0,10,0">
              <StackPanel>
                <Border Height="3" Width="28" CornerRadius="2" HorizontalAlignment="Left" Margin="0,0,0,8" Background="{DynamicResource Primary}"/>
                <TextBlock x:Name="StatTotal" Style="{StaticResource StatValue}" Text="0"/>
                <TextBlock Style="{StaticResource Caption}" Text="Total checks"/>
              </StackPanel>
            </Border>
            <Border Style="{StaticResource Card}" Padding="14,12" Margin="0,0,10,0">
              <StackPanel>
                <Border Height="3" Width="28" CornerRadius="2" HorizontalAlignment="Left" Margin="0,0,0,8" Background="{DynamicResource CatEntra}"/>
                <TextBlock x:Name="StatEntra" Style="{StaticResource StatValue}" Text="0"/>
                <TextBlock Style="{StaticResource Caption}" Text="Work/School"/>
              </StackPanel>
            </Border>
            <Border Style="{StaticResource Card}" Padding="14,12" Margin="0,0,10,0">
              <StackPanel>
                <Border Height="3" Width="28" CornerRadius="2" HorizontalAlignment="Left" Margin="0,0,0,8" Background="{DynamicResource CatMSA}"/>
                <TextBlock x:Name="StatMSA" Style="{StaticResource StatValue}" Text="0"/>
                <TextBlock Style="{StaticResource Caption}" Text="Personal"/>
              </StackPanel>
            </Border>
            <Border Style="{StaticResource Card}" Padding="14,12" Margin="0,0,10,0">
              <StackPanel>
                <Border Height="3" Width="28" CornerRadius="2" HorizontalAlignment="Left" Margin="0,0,0,8" Background="{DynamicResource CatFederated}"/>
                <TextBlock x:Name="StatFederated" Style="{StaticResource StatValue}" Text="0"/>
                <TextBlock Style="{StaticResource Caption}" Text="Federated"/>
              </StackPanel>
            </Border>
            <Border Style="{StaticResource Card}" Padding="14,12" Margin="0,0,10,0">
              <StackPanel>
                <Border Height="3" Width="28" CornerRadius="2" HorizontalAlignment="Left" Margin="0,0,0,8" Background="{DynamicResource CatDual}"/>
                <TextBlock x:Name="StatDual" Style="{StaticResource StatValue}" Text="0"/>
                <TextBlock Style="{StaticResource Caption}" Text="Dual identity"/>
              </StackPanel>
            </Border>
            <Border Style="{StaticResource Card}" Padding="14,12" Margin="0,0,10,0">
              <StackPanel>
                <Border Height="3" Width="28" CornerRadius="2" HorizontalAlignment="Left" Margin="0,0,0,8" Background="{DynamicResource CatGCCHigh}"/>
                <TextBlock x:Name="StatSovereign" Style="{StaticResource StatValue}" Text="0"/>
                <TextBlock Style="{StaticResource Caption}" Text="GCC High / China"/>
              </StackPanel>
            </Border>
            <Border Style="{StaticResource Card}" Padding="14,12" Margin="0,0,10,0">
              <StackPanel>
                <Border Height="3" Width="28" CornerRadius="2" HorizontalAlignment="Left" Margin="0,0,0,8" Background="{DynamicResource CatNotMicrosoft}"/>
                <TextBlock x:Name="StatNotMicrosoft" Style="{StaticResource StatValue}" Text="0"/>
                <TextBlock Style="{StaticResource Caption}" Text="Not Microsoft"/>
              </StackPanel>
            </Border>
            <Border Style="{StaticResource Card}" Padding="14,12">
              <StackPanel>
                <Border Height="3" Width="28" CornerRadius="2" HorizontalAlignment="Left" Margin="0,0,0,8" Background="{DynamicResource CatError}"/>
                <TextBlock x:Name="StatError" Style="{StaticResource StatValue}" Text="0"/>
                <TextBlock Style="{StaticResource Caption}" Text="Errors / unknown"/>
              </StackPanel>
            </Border>
          </UniformGrid>

          <Border Grid.Row="1" Style="{StaticResource Card}" Margin="0,16,0,0">
            <Grid>
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
              </Grid.RowDefinitions>
              <Grid>
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <Grid>
                  <TextBox x:Name="HistorySearch" MinHeight="36" AutomationProperties.Name="Search history" ToolTip="Search history (Ctrl+F)" AutomationProperties.HelpText="Search history (Ctrl+F)"/>
                  <TextBlock x:Name="HistorySearchHint" Text="Search by username, account type, cloud, organization or origin" Foreground="{DynamicResource Muted}"
                             Margin="11,0,0,0" VerticalAlignment="Center" IsHitTestVisible="False"/>
                </Grid>
                <StackPanel Grid.Column="1" Orientation="Horizontal" Margin="12,0,0,0">
                  <Button x:Name="HistoryExportButton" Content="Export CSV" ToolTip="Save the rows shown, after filters and search (Ctrl+E)" AutomationProperties.HelpText="Save the rows shown, after filters and search (Ctrl+E)"/>
                  <Button x:Name="HistoryExportHtmlButton" Content="HTML report" Margin="8,0,0,0" ToolTip="Executive summary of the rows shown" AutomationProperties.HelpText="Executive summary of the rows shown"/>
                  <Button x:Name="HistoryExportPdfButton" Content="PDF report" Margin="8,0,0,0" ToolTip="The same report as a PDF (uses Microsoft Edge)" AutomationProperties.HelpText="The same report as a PDF (uses Microsoft Edge)"/>
                  <Button x:Name="LogOpenButton" Content="Open log" Margin="8,0,0,0" ToolTip="Open today's log file" AutomationProperties.HelpText="Open today's log file"/>
                  <Button x:Name="LogExportButton" Content="Export log" Margin="8,0,0,0" ToolTip="Save today's log file as CSV" AutomationProperties.HelpText="Save today's log file as CSV"/>
                </StackPanel>
              </Grid>
              <WrapPanel Grid.Row="1" Margin="0,10,0,0">
                <ComboBox x:Name="FilterType" Margin="0,0,8,8" ToolTip="Filter by account type" AutomationProperties.HelpText="Filter by account type" AutomationProperties.Name="Filter by account type">
                  <ComboBoxItem Content="All account types" Tag="All" IsSelected="True"/>
                  <ComboBoxItem Content="Work/School" Tag="Entra"/>
                  <ComboBoxItem Content="Personal" Tag="MSA"/>
                  <ComboBoxItem Content="Federated" Tag="Federated"/>
                  <ComboBoxItem Content="Dual identity" Tag="Dual"/>
                  <ComboBoxItem Content="GCC High / China" Tag="Sovereign"/>
                  <ComboBoxItem Content="Not Microsoft" Tag="NotMicrosoft"/>
                  <ComboBoxItem Content="Errors / unknown" Tag="Error"/>
                </ComboBox>
                <ComboBox x:Name="FilterCloud" Margin="0,0,8,8" ToolTip="Filter by cloud" AutomationProperties.HelpText="Filter by cloud" AutomationProperties.Name="Filter by cloud">
                  <ComboBoxItem Content="All clouds" Tag="All" IsSelected="True"/>
                  <ComboBoxItem Content="Commercial" Tag="Commercial"/>
                  <ComboBoxItem Content="GCC High" Tag="GCCHigh"/>
                  <ComboBoxItem Content="China" Tag="China"/>
                  <ComboBoxItem Content="Not found (N/A)" Tag="N/A"/>
                </ComboBox>
                <ComboBox x:Name="FilterStatus" Margin="0,0,8,8" ToolTip="Filter by how certain the answer is" AutomationProperties.HelpText="Filter by how certain the answer is" AutomationProperties.Name="Filter by how certain the answer is">
                  <ComboBoxItem Content="All statuses" Tag="All" IsSelected="True"/>
                  <ComboBoxItem Content="Definitive" Tag="OK"/>
                  <ComboBoxItem Content="Inconclusive" Tag="Inconclusive"/>
                  <ComboBoxItem Content="Throttled" Tag="Throttled"/>
                  <ComboBoxItem Content="Temporarily unavailable" Tag="Unavailable"/>
                  <ComboBoxItem Content="Invalid address" Tag="Invalid"/>
                  <ComboBoxItem Content="Changed since previous check" Tag="Changed"/>
                </ComboBox>
                <ComboBox x:Name="FilterFederation" Margin="0,0,8,8" ToolTip="Filter by federation" AutomationProperties.HelpText="Filter by federation" AutomationProperties.Name="Filter by federation">
                  <ComboBoxItem Content="Federated or not" Tag="All" IsSelected="True"/>
                  <ComboBoxItem Content="Federated only" Tag="Yes"/>
                  <ComboBoxItem Content="Not federated" Tag="No"/>
                </ComboBox>
                <ComboBox x:Name="FilterSource" Margin="0,0,8,8" ToolTip="Filter by where the answer came from" AutomationProperties.HelpText="Filter by where the answer came from" AutomationProperties.Name="Filter by where the answer came from">
                  <ComboBoxItem Content="Live and cache" Tag="All" IsSelected="True"/>
                  <ComboBoxItem Content="Live only" Tag="Live"/>
                  <ComboBoxItem Content="Cache only" Tag="Cache"/>
                </ComboBox>
                <ComboBox x:Name="FilterDate" Margin="0,0,8,8" ToolTip="Filter by date" AutomationProperties.HelpText="Filter by date" AutomationProperties.Name="Filter by date">
                  <ComboBoxItem Content="All dates" Tag="All" IsSelected="True"/>
                  <ComboBoxItem Content="Today" Tag="0"/>
                  <ComboBoxItem Content="Last 7 days" Tag="7"/>
                  <ComboBoxItem Content="Last 30 days" Tag="30"/>
                  <ComboBoxItem Content="Last 90 days" Tag="90"/>
                </ComboBox>
                <Button x:Name="FilterResetButton" Content="Reset filters" Margin="0,0,8,8"/>
              </WrapPanel>
              <DockPanel Grid.Row="2" LastChildFill="False" Margin="0,0,0,2">
                <TextBlock x:Name="HistoryShownText" DockPanel.Dock="Right" Style="{StaticResource Caption}" VerticalAlignment="Center"/>
                <TextBlock DockPanel.Dock="Left" Style="{StaticResource Caption}" VerticalAlignment="Center">
                  <Run Text="Email addresses and tenant details are stored only on this computer."/>
                  <Hyperlink x:Name="HistoryPrivacyLink" Foreground="{DynamicResource Link}">Privacy settings</Hyperlink>
                </TextBlock>
              </DockPanel>
              <DataGrid x:Name="HistoryGrid" Grid.Row="3" Margin="0,10,0,0" AutomationProperties.Name="History">
                <DataGrid.Columns>
                  <DataGridTextColumn Header="Time" Binding="{Binding Timestamp}" Width="140"/>
                  <DataGridTextColumn Header="Username" Binding="{Binding Username}" Width="2*" MinWidth="170" ElementStyle="{StaticResource CellText}"/>
                  <DataGridTextColumn Header="Cloud" Binding="{Binding FoundInCloud}" Width="90"/>
                  <DataGridTextColumn Header="Account type" Binding="{Binding AccountType}" Width="2.4*" MinWidth="210" ElementStyle="{StaticResource TypeText}"/>
                  <DataGridTextColumn Header="Status" Binding="{Binding StatusLabel}" Width="1*" MinWidth="110" ElementStyle="{StaticResource StatusText}"/>
                  <DataGridTextColumn Header="Changes" Binding="{Binding Changes}" Width="1.5*" MinWidth="120" ElementStyle="{StaticResource ChangeText}"/>
                  <DataGridTextColumn Header="Organization" Binding="{Binding Organization}" Width="1.4*" MinWidth="120" ElementStyle="{StaticResource CellText}"/>
                  <DataGridTextColumn Header="Federation provider" Binding="{Binding FederationProvider}" Width="1.3*" MinWidth="120" ElementStyle="{StaticResource CellText}"/>
                  <DataGridTextColumn Header="Source" Binding="{Binding Source}" Width="65"/>
                  <DataGridTextColumn Header="Origin" Binding="{Binding Origin}" Width="65"/>
                </DataGrid.Columns>
              </DataGrid>
            </Grid>
          </Border>
        </Grid>
      </TabItem>

      <!-- API server -->
      <TabItem x:Name="TabApi">
        <TabItem.Header>
          <StackPanel Orientation="Horizontal">
            <TextBlock Style="{StaticResource Glyph}" Text="&#xE774;"/>
            <TextBlock Text="API server"/>
          </StackPanel>
        </TabItem.Header>
        <Grid Margin="20">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="380"/>
            <ColumnDefinition Width="*"/>
          </Grid.ColumnDefinitions>

          <Border Style="{StaticResource Card}" Margin="0,0,16,0" Padding="0">
            <ScrollViewer VerticalScrollBarVisibility="Auto">
              <StackPanel Margin="18">
                <TextBlock Style="{StaticResource H2}" Text="Local REST API"/>
                <TextBlock Style="{StaticResource Caption}" Margin="0,4,0,14"
                           Text="Runs the same lookup over HTTP on this computer only, with a web page for checking accounts in a browser. Results go to History and share the 24-hour cache. Start and stop it with the API server switch at the top."/>
                <Border CornerRadius="8" Background="{DynamicResource GridBg}" BorderBrush="{DynamicResource Border}" BorderThickness="1" Padding="12,10">
                  <StackPanel Orientation="Horizontal">
                    <Ellipse x:Name="ApiDot" Width="10" Height="10" Margin="0,0,10,0" VerticalAlignment="Center"/>
                    <TextBlock x:Name="ApiStatusText" FontWeight="SemiBold" Text="Stopped" VerticalAlignment="Center"/>
                  </StackPanel>
                </Border>
                <TextBlock Style="{StaticResource FieldLabel}" Margin="0,16,0,6" Text="Port"/>
                <TextBox x:Name="ApiPortBox" Width="120" MinHeight="36" HorizontalAlignment="Left"/>
                <TextBlock Style="{StaticResource FieldLabel}" Margin="0,16,0,6" Text="Base URL"/>
                <TextBlock>
                  <Hyperlink x:Name="ApiUrlLink" Foreground="{DynamicResource Link}"><Run x:Name="ApiUrlRun"/></Hyperlink>
                </TextBlock>
                <Button x:Name="ApiOpenWebButton" Style="{StaticResource PrimaryButton}" Content="Open web page" HorizontalAlignment="Left" Margin="0,10,0,0" IsEnabled="False"/>
                <CheckBox x:Name="ApiOpenBrowserCheck" Style="{StaticResource BodySwitch}" Margin="0,16,0,0">
                  <TextBlock TextWrapping="Wrap" Text="Open the web page when the API server starts"/>
                </CheckBox>
                <CheckBox x:Name="ApiTrayCheck" Style="{StaticResource BodySwitch}" Margin="0,10,0,0">
                  <TextBlock TextWrapping="Wrap" Text="While the API server runs, closing or minimizing the window keeps the app in the system tray"/>
                </CheckBox>
                <TextBlock Style="{StaticResource FieldLabel}" Margin="0,16,0,6" Text="Endpoints"/>
                <TextBox x:Name="ApiEndpointsBox" Style="{StaticResource MonoBox}" Height="170"/>
                <TextBlock Style="{StaticResource FieldLabel}" Margin="0,16,0,6" Text="Send a test request"/>
                <Grid>
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                  </Grid.ColumnDefinitions>
                  <TextBox x:Name="ApiTestEmail" MinHeight="36"/>
                  <Button x:Name="ApiTestButton" Grid.Column="1" Style="{StaticResource PrimaryButton}" Content="Send" MinWidth="80" Margin="8,0,0,0"/>
                </Grid>
              </StackPanel>
            </ScrollViewer>
          </Border>

          <Border Grid.Column="1" Style="{StaticResource Card}">
            <Grid>
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
              </Grid.RowDefinitions>
              <DockPanel LastChildFill="False">
                <TextBlock DockPanel.Dock="Left" Style="{StaticResource H2}" Text="Request log"/>
                <Button x:Name="ApiClearLogButton" DockPanel.Dock="Right" Content="Clear log"/>
              </DockPanel>
              <DataGrid x:Name="ApiLogGrid" Grid.Row="1" Margin="0,12,0,0" AutomationProperties.Name="API request log">
                <DataGrid.Columns>
                  <DataGridTextColumn Header="Time" Binding="{Binding Time}" Width="80"/>
                  <DataGridTextColumn Header="Method" Binding="{Binding Method}" Width="70"/>
                  <DataGridTextColumn Header="Path" Binding="{Binding Path}" Width="*" ElementStyle="{StaticResource CellText}"/>
                  <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="60"/>
                  <DataGridTextColumn Header="ms" Binding="{Binding DurationMs}" Width="70"/>
                  <DataGridTextColumn Header="Client" Binding="{Binding Client}" Width="110"/>
                </DataGrid.Columns>
              </DataGrid>
              <TextBlock Grid.Row="2" Style="{StaticResource FieldLabel}" Margin="0,14,0,6" Text="Last test response"/>
              <TextBox x:Name="ApiTestResponse" Grid.Row="3" Style="{StaticResource MonoBox}" Height="170"/>
            </Grid>
          </Border>
        </Grid>
      </TabItem>

      <!-- Diagnostics -->
      <TabItem x:Name="TabDiagnostics">
        <TabItem.Header>
          <StackPanel Orientation="Horizontal">
            <TextBlock Style="{StaticResource Glyph}" Text="&#xE9D9;"/>
            <TextBlock Text="Diagnostics"/>
          </StackPanel>
        </TabItem.Header>
        <ScrollViewer VerticalScrollBarVisibility="Auto">
          <Grid Margin="20">
            <Grid.RowDefinitions>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <UniformGrid Columns="4">
              <Border Style="{StaticResource Card}" Padding="14,12" Margin="0,0,10,10" ToolTip="Lookups sent to Microsoft (not answered from the cache)" AutomationProperties.HelpText="Lookups sent to Microsoft (not answered from the cache)">
                <StackPanel>
                  <Border Height="3" Width="28" CornerRadius="2" HorizontalAlignment="Left" Margin="0,0,0,8" Background="{DynamicResource Primary}"/>
                  <TextBlock x:Name="DiagLive" Style="{StaticResource StatValue}" Text="0"/>
                  <TextBlock Style="{StaticResource Caption}" Text="Live lookups"/>
                </StackPanel>
              </Border>
              <Border Style="{StaticResource Card}" Padding="14,12" Margin="0,0,10,10" ToolTip="Answers taken from the cache" AutomationProperties.HelpText="Answers taken from the cache">
                <StackPanel>
                  <Border Height="3" Width="28" CornerRadius="2" HorizontalAlignment="Left" Margin="0,0,0,8" Background="{DynamicResource CatEntra}"/>
                  <TextBlock x:Name="DiagCacheHits" Style="{StaticResource StatValue}" Text="0"/>
                  <TextBlock Style="{StaticResource Caption}" Text="Cache hits"/>
                </StackPanel>
              </Border>
              <Border Style="{StaticResource Card}" Padding="14,12" Margin="0,0,10,10">
                <StackPanel>
                  <Border Height="3" Width="28" CornerRadius="2" HorizontalAlignment="Left" Margin="0,0,0,8" Background="{DynamicResource CatMSA}"/>
                  <TextBlock x:Name="DiagHitRate" Style="{StaticResource StatValue}" Text="-"/>
                  <TextBlock Style="{StaticResource Caption}" Text="Cache hit rate"/>
                </StackPanel>
              </Border>
              <Border Style="{StaticResource Card}" Padding="14,12" Margin="0,0,0,10" ToolTip="Requests sent again after Microsoft throttled or was briefly unavailable" AutomationProperties.HelpText="Requests sent again after Microsoft throttled or was briefly unavailable">
                <StackPanel>
                  <Border Height="3" Width="28" CornerRadius="2" HorizontalAlignment="Left" Margin="0,0,0,8" Background="{DynamicResource CatDual}"/>
                  <TextBlock x:Name="DiagRetries" Style="{StaticResource StatValue}" Text="0"/>
                  <TextBlock Style="{StaticResource Caption}" Text="Retries"/>
                </StackPanel>
              </Border>
              <Border Style="{StaticResource Card}" Padding="14,12" Margin="0,0,10,10" ToolTip="Times Microsoft rate-limited a request (HTTP 429 or ThrottleStatus)" AutomationProperties.HelpText="Times Microsoft rate-limited a request (HTTP 429 or ThrottleStatus)">
                <StackPanel>
                  <Border Height="3" Width="28" CornerRadius="2" HorizontalAlignment="Left" Margin="0,0,0,8" Background="{DynamicResource CatError}"/>
                  <TextBlock x:Name="DiagThrottle" Style="{StaticResource StatValue}" Text="0"/>
                  <TextBlock Style="{StaticResource Caption}" Text="Throttling events"/>
                </StackPanel>
              </Border>
              <Border Style="{StaticResource Card}" Padding="14,12" Margin="0,0,10,10" ToolTip="Lookups that failed because Microsoft could not be reached" AutomationProperties.HelpText="Lookups that failed because Microsoft could not be reached">
                <StackPanel>
                  <Border Height="3" Width="28" CornerRadius="2" HorizontalAlignment="Left" Margin="0,0,0,8" Background="{DynamicResource CatNotMicrosoft}"/>
                  <TextBlock x:Name="DiagUnavailable" Style="{StaticResource StatValue}" Text="0"/>
                  <TextBlock Style="{StaticResource Caption}" Text="Unavailable"/>
                </StackPanel>
              </Border>
              <Border Style="{StaticResource Card}" Padding="14,12" Margin="0,0,10,10" ToolTip="Answers that were not definitive" AutomationProperties.HelpText="Answers that were not definitive">
                <StackPanel>
                  <Border Height="3" Width="28" CornerRadius="2" HorizontalAlignment="Left" Margin="0,0,0,8" Background="{DynamicResource CatFederated}"/>
                  <TextBlock x:Name="DiagInconclusive" Style="{StaticResource StatValue}" Text="0"/>
                  <TextBlock Style="{StaticResource Caption}" Text="Inconclusive"/>
                </StackPanel>
              </Border>
              <Border Style="{StaticResource Card}" Padding="14,12" Margin="0,0,0,10" ToolTip="Requests handled by the local API server" AutomationProperties.HelpText="Requests handled by the local API server">
                <StackPanel>
                  <Border Height="3" Width="28" CornerRadius="2" HorizontalAlignment="Left" Margin="0,0,0,8" Background="{DynamicResource CatGCCHigh}"/>
                  <TextBlock x:Name="DiagApiRequests" Style="{StaticResource StatValue}" Text="0"/>
                  <TextBlock Style="{StaticResource Caption}" Text="API requests"/>
                </StackPanel>
              </Border>
            </UniformGrid>

            <Border Grid.Row="1" Style="{StaticResource Card}" Margin="0,6,0,16">
              <Grid>
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="380"/>
                  <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <StackPanel>
                  <TextBlock Style="{StaticResource H2}" Text="Microsoft endpoint health, last 5 minutes"/>
                  <Grid Margin="0,12,0,0">
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="170"/>
                      <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <Grid.RowDefinitions>
                      <RowDefinition Height="Auto"/>
                      <RowDefinition Height="Auto"/>
                      <RowDefinition Height="Auto"/>
                      <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    <TextBlock Grid.Row="0" Style="{StaticResource FieldName}" Text="Availability" ToolTip="Share of calls to Microsoft that got an answer" AutomationProperties.HelpText="Share of calls to Microsoft that got an answer"/>
                    <TextBlock Grid.Row="0" Grid.Column="1" x:Name="DiagRecentAvailability" Style="{StaticResource FieldValue}" Text="-"/>
                    <TextBlock Grid.Row="1" Style="{StaticResource FieldName}" Text="Median response" ToolTip="Half of the lookups were answered faster than this" AutomationProperties.HelpText="Half of the lookups were answered faster than this"/>
                    <TextBlock Grid.Row="1" Grid.Column="1" x:Name="DiagRecentP50" Style="{StaticResource FieldValue}" Text="-"/>
                    <TextBlock Grid.Row="2" Style="{StaticResource FieldName}" Text="95th percentile" ToolTip="95 out of 100 lookups were answered faster than this" AutomationProperties.HelpText="95 out of 100 lookups were answered faster than this"/>
                    <TextBlock Grid.Row="2" Grid.Column="1" x:Name="DiagRecentP95" Style="{StaticResource FieldValue}" Text="-"/>
                    <TextBlock Grid.Row="3" Style="{StaticResource FieldName}" Text="Throttling rate" ToolTip="Throttling events as a share of lookups" AutomationProperties.HelpText="Throttling events as a share of lookups"/>
                    <TextBlock Grid.Row="3" Grid.Column="1" x:Name="DiagRecentThrottle" Style="{StaticResource FieldValue}" Text="-"/>
                  </Grid>
                  <Border CornerRadius="8" Background="{DynamicResource Secondary}" Padding="12,10" Margin="0,12,0,0">
                    <StackPanel>
                      <TextBlock Text="Suggested safe batch settings" FontWeight="SemiBold" Foreground="{DynamicResource SecondaryText}"/>
                      <TextBlock x:Name="DiagSuggestion" TextWrapping="Wrap" Margin="0,4,0,0" Foreground="{DynamicResource SecondaryText}"/>
                      <Button x:Name="DiagApplyButton" Style="{StaticResource PrimaryButton}" Content="Apply to batch settings" HorizontalAlignment="Left" Margin="0,10,0,0"/>
                    </StackPanel>
                  </Border>
                </StackPanel>
                <Grid Grid.Column="1" Margin="24,0,0,0">
                  <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="Auto"/>
                  </Grid.RowDefinitions>
                  <TextBlock Style="{StaticResource H2}" Text="Lookup response time, last 30 minutes"/>
                  <Border Grid.Row="1" Margin="0,10,0,0" MinHeight="170" CornerRadius="6" Background="{DynamicResource GridBg}" BorderBrush="{DynamicResource Border}" BorderThickness="1"
                          AutomationProperties.Name="Lookup response time chart">
                    <Grid>
                      <Canvas x:Name="DiagChart" ClipToBounds="True" Margin="10"/>
                      <TextBlock x:Name="DiagChartEmpty" Style="{StaticResource Caption}" Text="No lookups in the last 30 minutes." HorizontalAlignment="Center" VerticalAlignment="Center"/>
                    </Grid>
                  </Border>
                  <TextBlock x:Name="DiagChartCaption" Grid.Row="2" Style="{StaticResource Caption}" Margin="0,6,0,0"/>
                </Grid>
              </Grid>
            </Border>

            <Grid Grid.Row="2">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="340"/>
              </Grid.ColumnDefinitions>
              <Border Style="{StaticResource Card}">
                <Grid>
                  <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                  </Grid.RowDefinitions>
                  <TextBlock Style="{StaticResource H2}" Text="Microsoft endpoints, this session"/>
                  <TextBlock Grid.Row="1" Style="{StaticResource Caption}" Margin="0,4,0,10"
                             Text="Every call to Microsoft, including retries. Availability is the share of calls that got an answer."/>
                  <DataGrid x:Name="DiagEndpointGrid" Grid.Row="2" MinHeight="160" AutomationProperties.Name="Microsoft endpoints">
                    <DataGrid.Columns>
                      <DataGridTextColumn Header="Endpoint" Binding="{Binding Endpoint}" Width="2*" MinWidth="200" ElementStyle="{StaticResource CellText}"/>
                      <DataGridTextColumn Header="Calls" Binding="{Binding Calls}" Width="70"/>
                      <DataGridTextColumn Header="Failures" Binding="{Binding Failures}" Width="80"/>
                      <DataGridTextColumn Header="Average ms" Binding="{Binding AverageMs}" Width="95"/>
                      <DataGridTextColumn Header="Last ms" Binding="{Binding LastMs}" Width="80"/>
                      <DataGridTextColumn Header="Availability" Binding="{Binding Availability}" Width="100"/>
                    </DataGrid.Columns>
                  </DataGrid>
                </Grid>
              </Border>
              <Border Grid.Column="1" Style="{StaticResource Card}" Margin="16,0,0,0">
                <StackPanel>
                  <TextBlock Style="{StaticResource H2}" Text="Status"/>
                  <TextBlock Style="{StaticResource FieldLabel}" Margin="0,14,0,4" Text="Last successful lookup"/>
                  <TextBlock x:Name="DiagLastSuccess" TextWrapping="Wrap" Text="None yet"/>
                  <TextBlock Style="{StaticResource FieldLabel}" Margin="0,14,0,4" Text="Last problem"/>
                  <TextBlock x:Name="DiagLastError" TextWrapping="Wrap" Text="None"/>
                  <TextBlock Style="{StaticResource FieldLabel}" Margin="0,14,0,4" Text="API server"/>
                  <TextBlock x:Name="DiagApiState" TextWrapping="Wrap" Text="Off"/>
                  <TextBlock Style="{StaticResource FieldLabel}" Margin="0,14,0,4" Text="API requests refused"
                             ToolTip="Requests refused because they came from another website, had no valid token or went over the rate limit" AutomationProperties.HelpText="Requests refused because they came from another website, had no valid token or went over the rate limit"/>
                  <TextBlock x:Name="DiagApiRejected" Text="0"/>
                  <TextBlock Style="{StaticResource FieldLabel}" Margin="0,14,0,4" Text="Counting since"/>
                  <TextBlock x:Name="DiagSince"/>
                  <Button x:Name="DiagResetButton" Content="Reset counters" HorizontalAlignment="Left" Margin="0,18,0,0"/>
                  <TextBlock Style="{StaticResource Caption}" Margin="0,10,0,0" Text="Counters cover this session only and are not saved."/>
                </StackPanel>
              </Border>
            </Grid>
          </Grid>
        </ScrollViewer>
      </TabItem>

      <!-- Settings -->
      <TabItem x:Name="TabSettings">
        <TabItem.Header>
          <StackPanel Orientation="Horizontal">
            <TextBlock Style="{StaticResource Glyph}" Text="&#xE713;"/>
            <TextBlock Text="Settings"/>
          </StackPanel>
        </TabItem.Header>
        <ScrollViewer VerticalScrollBarVisibility="Auto">
          <Grid Margin="20">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="16"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <StackPanel>
              <Border Style="{StaticResource Card}">
                <StackPanel>
                  <TextBlock Style="{StaticResource H2}" Text="Privacy and stored data"/>
                  <Border CornerRadius="8" Background="{DynamicResource Secondary}" Padding="12,10" Margin="0,12,0,0">
                    <TextBlock TextWrapping="Wrap" Foreground="{DynamicResource SecondaryText}"
                               Text="Email addresses, account types and tenant details (organization, tenant ID and region) are stored only on this computer, in the folder below. Nothing is sent anywhere except Microsoft's sign-in endpoints."/>
                  </Border>
                  <TextBlock Style="{StaticResource FieldLabel}" Margin="0,14,0,6" Text="Data folder"/>
                  <DockPanel>
                    <Button x:Name="SettingsOpenFolderButton" DockPanel.Dock="Right" Content="Open folder" Margin="12,0,0,0"/>
                    <TextBlock x:Name="SettingsDataFolder" TextWrapping="Wrap" VerticalAlignment="Center"/>
                  </DockPanel>
                  <CheckBox x:Name="EncryptStoredCheck" Style="{StaticResource BodySwitch}" Margin="0,16,0,0"
                            ToolTip="Uses Windows data protection. Only your Windows account on this computer can read the files." AutomationProperties.HelpText="Uses Windows data protection. Only your Windows account on this computer can read the files.">
                    <TextBlock TextWrapping="Wrap" Text="Encrypt history and the cache on disk"/>
                  </CheckBox>
                  <CheckBox x:Name="PrivateSessionCheck" Style="{StaticResource BodySwitch}" Margin="0,10,0,0"
                            ToolTip="Checks made while this is on are kept only until the app closes. Turns off again when the app restarts." AutomationProperties.HelpText="Checks made while this is on are kept only until the app closes. Turns off again when the app restarts.">
                    <TextBlock TextWrapping="Wrap" Text="Private session: do not save history, cache or logs for checks made now"/>
                  </CheckBox>
                  <CheckBox x:Name="RedactExportsCheck" Style="{StaticResource BodySwitch}" Margin="0,10,0,0"
                            ToolTip="Exports show j***@contoso.com instead of the full address, and imported columns are blanked" AutomationProperties.HelpText="Exports show j***@contoso.com instead of the full address, and imported columns are blanked">
                    <TextBlock TextWrapping="Wrap" Text="Redact email addresses and imported columns in exports"/>
                  </CheckBox>
                  <TextBlock Style="{StaticResource FieldLabel}" Margin="0,16,0,6" Text="Keep history for"/>
                  <StackPanel Orientation="Horizontal">
                    <TextBox x:Name="HistoryRetentionBox" Width="80" MinHeight="34" AutomationProperties.Name="Keep history for, days"/>
                    <TextBlock Style="{StaticResource Caption}" Margin="10,0,0,0" VerticalAlignment="Center" Text="days (0 keeps everything, up to 5,000 checks)"/>
                  </StackPanel>
                  <TextBlock Style="{StaticResource FieldLabel}" Margin="0,14,0,6" Text="Delete log files older than"/>
                  <StackPanel Orientation="Horizontal">
                    <TextBox x:Name="LogRetentionBox" Width="80" MinHeight="34" AutomationProperties.Name="Delete log files older than, days"/>
                    <TextBlock Style="{StaticResource Caption}" Margin="10,0,0,0" VerticalAlignment="Center" Text="days (0 keeps all logs)"/>
                  </StackPanel>
                  <CheckBox x:Name="ClearOnExitCheck" Style="{StaticResource BodySwitch}" Margin="0,16,0,0">
                    <TextBlock TextWrapping="Wrap" Text="Clear history and the lookup cache every time the app closes"/>
                  </CheckBox>
                  <WrapPanel Margin="0,16,0,0">
                    <Button x:Name="HistoryClearButton" Style="{StaticResource DangerButton}" Content="Clear history" Margin="0,0,8,8"/>
                    <Button x:Name="CacheClearButton" Content="Clear cache" Margin="0,0,8,8"/>
                    <Button x:Name="LogDeleteButton" Content="Delete all logs" Margin="0,0,8,8"/>
                    <Button x:Name="ExportDeleteButton" Content="Delete exported files" Margin="0,0,8,8"
                            ToolTip="Deletes every CSV, JSON, HTML and PDF file this app has exported" AutomationProperties.HelpText="Deletes every CSV, JSON, HTML and PDF file this app has exported"/>
                  </WrapPanel>
                </StackPanel>
              </Border>

              <Border Style="{StaticResource Card}" Margin="0,16,0,0">
                <StackPanel>
                  <TextBlock Style="{StaticResource H2}" Text="Cache"/>
                  <TextBlock Style="{StaticResource Caption}" Margin="0,4,0,0"
                             Text="How long an answer is reused before Microsoft is asked again, by outcome. 0 means never reuse."/>
                  <Grid Margin="0,12,0,0">
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="210"/>
                      <ColumnDefinition Width="Auto"/>
                      <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <Grid.RowDefinitions>
                      <RowDefinition Height="Auto"/>
                      <RowDefinition Height="Auto"/>
                      <RowDefinition Height="Auto"/>
                      <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    <TextBlock Grid.Row="0" Style="{StaticResource FieldName}" VerticalAlignment="Center" Text="Definitive answers"/>
                    <TextBox Grid.Row="0" Grid.Column="1" x:Name="StableCacheBox" Width="80" MinHeight="34" Margin="0,4" AutomationProperties.Name="Cache definitive answers, hours"/>
                    <TextBlock Grid.Row="0" Grid.Column="2" Style="{StaticResource Caption}" Margin="10,0,0,0" VerticalAlignment="Center" Text="hours (default 24)"/>
                    <TextBlock Grid.Row="1" Style="{StaticResource FieldName}" VerticalAlignment="Center" Text="Inconclusive answers"/>
                    <TextBox Grid.Row="1" Grid.Column="1" x:Name="InconclusiveCacheBox" Width="80" MinHeight="34" Margin="0,4" AutomationProperties.Name="Cache inconclusive answers, hours"/>
                    <TextBlock Grid.Row="1" Grid.Column="2" Style="{StaticResource Caption}" Margin="10,0,0,0" VerticalAlignment="Center" Text="hours (default 24)"/>
                    <TextBlock Grid.Row="2" Style="{StaticResource FieldName}" VerticalAlignment="Center" Text="Throttled or unavailable"/>
                    <TextBox Grid.Row="2" Grid.Column="1" x:Name="FailureCacheBox" Width="80" MinHeight="34" Margin="0,4" AutomationProperties.Name="Cache throttled or unavailable results, minutes"/>
                    <TextBlock Grid.Row="2" Grid.Column="2" Style="{StaticResource Caption}" Margin="10,0,0,0" VerticalAlignment="Center" Text="minutes (default 0, up to 60)"/>
                    <TextBlock Grid.Row="3" Style="{StaticResource FieldName}" VerticalAlignment="Center" Text="Stale after" ToolTip="Used by the Refresh stale only option on the Batch check tab" AutomationProperties.HelpText="Used by the Refresh stale only option on the Batch check tab"/>
                    <TextBox Grid.Row="3" Grid.Column="1" x:Name="StaleAfterBox" Width="80" MinHeight="34" Margin="0,4" AutomationProperties.Name="Answers are stale after, hours"/>
                    <TextBlock Grid.Row="3" Grid.Column="2" Style="{StaticResource Caption}" Margin="10,0,0,0" VerticalAlignment="Center" Text="hours, for Refresh stale only (default 12)"/>
                  </Grid>
                </StackPanel>
              </Border>

              <Border Style="{StaticResource Card}" Margin="0,16,0,0">
                <StackPanel>
                  <TextBlock Style="{StaticResource H2}" Text="Batch lookups"/>
                  <TextBlock Style="{StaticResource FieldLabel}" Margin="0,14,0,6" Text="Delay between lookups"/>
                  <StackPanel Orientation="Horizontal">
                    <TextBox x:Name="BatchDelayBox" Width="80" MinHeight="34" AutomationProperties.Name="Delay between lookups, milliseconds"/>
                    <TextBlock Style="{StaticResource Caption}" Margin="10,0,0,0" VerticalAlignment="Center" Text="milliseconds, 0 to 10,000"/>
                  </StackPanel>
                  <TextBlock Style="{StaticResource FieldLabel}" Margin="0,14,0,6" Text="Parallel lookups"/>
                  <StackPanel Orientation="Horizontal">
                    <TextBox x:Name="BatchConcurrencyBox" Width="80" MinHeight="34" AutomationProperties.Name="Parallel lookups"/>
                    <TextBlock Style="{StaticResource Caption}" Margin="10,0,0,0" VerticalAlignment="Center" Text="1 to 4 at a time"/>
                  </StackPanel>
                  <TextBlock Style="{StaticResource Caption}" Margin="0,12,0,0"
                             Text="More parallel lookups or a shorter delay finish sooner, but make Microsoft throttling more likely. The Diagnostics tab suggests safe values from recent throttling."/>
                </StackPanel>
              </Border>

              <Border Style="{StaticResource Card}" Margin="0,16,0,0">
                <StackPanel>
                  <TextBlock Style="{StaticResource H2}" Text="Watched folder"/>
                  <TextBlock Style="{StaticResource Caption}" Margin="0,4,0,0"
                             Text="New .csv and .txt files dropped in this folder are checked automatically. Results are saved to a results subfolder, and the input file is moved to processed."/>
                  <TextBlock Style="{StaticResource FieldLabel}" Margin="0,14,0,6" Text="Folder"/>
                  <Grid>
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="*"/>
                      <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <TextBox x:Name="WatchFolderBox" IsReadOnly="True" MinHeight="34" AutomationProperties.Name="Watched folder"/>
                    <Button x:Name="WatchFolderBrowseButton" Grid.Column="1" Content="Choose" Margin="8,0,0,0"/>
                  </Grid>
                  <CheckBox x:Name="WatchFolderCheck" Style="{StaticResource BodySwitch}" Margin="0,14,0,0">
                    <TextBlock TextWrapping="Wrap" Text="Process new files in this folder automatically"/>
                  </CheckBox>
                  <TextBlock x:Name="WatchStatusText" Style="{StaticResource Caption}" Margin="0,8,0,0"/>
                </StackPanel>
              </Border>
            </StackPanel>

            <StackPanel Grid.Column="2">
              <Border Style="{StaticResource Card}">
                <StackPanel>
                  <TextBlock Style="{StaticResource H2}" Text="API security and limits"/>
                  <TextBlock Style="{StaticResource Caption}" Margin="0,6,0,0"
                             Text="The API only listens on this computer and always refuses requests sent by other websites, so a web page you visit cannot use it. A token also stops other programs on this computer from using it."/>
                  <CheckBox x:Name="ApiTokenCheck" Style="{StaticResource BodySwitch}" Margin="0,14,0,0">
                    <TextBlock TextWrapping="Wrap" Text="Require an API token"/>
                  </CheckBox>
                  <TextBlock Style="{StaticResource FieldLabel}" Margin="0,14,0,6" Text="Token"/>
                  <Grid>
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="*"/>
                      <ColumnDefinition Width="Auto"/>
                      <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <TextBox x:Name="ApiTokenBox" IsReadOnly="True" FontFamily="Consolas" MinHeight="34" AutomationProperties.Name="API token"/>
                    <Button x:Name="ApiTokenCopyButton" Grid.Column="1" Content="Copy" Margin="8,0,0,0"/>
                    <Button x:Name="ApiTokenNewButton" Grid.Column="2" Content="New token" Margin="8,0,0,0"
                            ToolTip="Create a new token. Programs using the old one stop working." AutomationProperties.HelpText="Create a new token. Programs using the old one stop working."/>
                  </Grid>
                  <TextBlock Style="{StaticResource Caption}" Margin="0,10,0,0"
                             Text="Programs send it in the X-API-Key header, or as Authorization: Bearer followed by the token. The web page adds it automatically when you open it from this computer."/>
                  <TextBlock Style="{StaticResource FieldLabel}" Margin="0,16,0,6" Text="Rate limit"/>
                  <StackPanel Orientation="Horizontal">
                    <TextBox x:Name="RateLimitBox" Width="80" MinHeight="34" AutomationProperties.Name="Rate limit, requests a minute"/>
                    <TextBlock Style="{StaticResource Caption}" Margin="10,0,0,0" VerticalAlignment="Center" Text="requests a minute per program (0 for no limit)"/>
                  </StackPanel>
                  <TextBlock Style="{StaticResource Caption}" Margin="0,10,0,0"
                             Text="OpenAPI description: /api/openapi.json. Monitoring: /health and /metrics (Prometheus)."/>
                </StackPanel>
              </Border>

              <Border Style="{StaticResource Card}" Margin="0,16,0,0">
                <StackPanel>
                  <TextBlock Style="{StaticResource H2}" Text="Updates"/>
                  <CheckBox x:Name="UpdateOnStartCheck" Style="{StaticResource BodySwitch}" Margin="0,12,0,0"
                            ToolTip="Asks GitHub for the latest release once at start-up. Nothing else is sent." AutomationProperties.HelpText="Asks GitHub for the latest release once at start-up. Nothing else is sent.">
                    <TextBlock TextWrapping="Wrap" Text="Check for updates when the app starts"/>
                  </CheckBox>
                  <WrapPanel Margin="0,14,0,0">
                    <Button x:Name="UpdateCheckButton" Content="Check for updates" Margin="0,0,8,8"/>
                    <Button x:Name="UpdateDownloadButton" Style="{StaticResource PrimaryButton}" Content="Download and verify" Margin="0,0,8,8" Visibility="Collapsed"
                            ToolTip="Downloads the release zip to your Downloads folder, checks its SHA-256 checksum and reports its signature. Nothing is installed." AutomationProperties.HelpText="Downloads the release zip to your Downloads folder, checks its SHA-256 checksum and reports its signature. Nothing is installed."/>
                    <Button x:Name="ChangelogButton" Content="What's new" Margin="0,0,8,8"/>
                  </WrapPanel>
                  <TextBlock x:Name="UpdateStatusText" TextWrapping="Wrap" Margin="0,4,0,0" Text="Not checked yet."/>
                </StackPanel>
              </Border>

              <Border Style="{StaticResource Card}" Margin="0,16,0,0">
                <StackPanel>
                  <TextBlock Style="{StaticResource H2}" Text="Display and accessibility"/>
                  <CheckBox x:Name="CompactTablesCheck" Style="{StaticResource BodySwitch}" Margin="0,12,0,0">
                    <TextBlock TextWrapping="Wrap" Text="Compact tables (smaller rows, more results on screen)"/>
                  </CheckBox>
                  <TextBlock Style="{StaticResource Caption}" Margin="0,10,0,0"
                             Text="Windows high contrast themes are followed automatically. Every colour pair in the light and dark themes meets the WCAG 4.5:1 contrast minimum. Right-click a row in Batch check or History for copy options."/>
                </StackPanel>
              </Border>

              <Border Style="{StaticResource Card}" Margin="0,16,0,0">
                <StackPanel>
                  <TextBlock Style="{StaticResource H2}" Text="Keyboard shortcuts"/>
                  <Grid Margin="0,12,0,0">
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="150"/>
                      <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <Grid.RowDefinitions>
                      <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                      <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                      <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    <TextBlock Grid.Row="0" Style="{StaticResource FieldValue}" FontFamily="Consolas" Text="Ctrl+1 ... Ctrl+6"/>
                    <TextBlock Grid.Row="0" Grid.Column="1" Style="{StaticResource FieldName}" Text="Switch tabs"/>
                    <TextBlock Grid.Row="1" Style="{StaticResource FieldValue}" FontFamily="Consolas" Text="Ctrl+,"/>
                    <TextBlock Grid.Row="1" Grid.Column="1" Style="{StaticResource FieldName}" Text="Settings"/>
                    <TextBlock Grid.Row="2" Style="{StaticResource FieldValue}" FontFamily="Consolas" Text="Ctrl+L"/>
                    <TextBlock Grid.Row="2" Grid.Column="1" Style="{StaticResource FieldName}" Text="Go to the email box"/>
                    <TextBlock Grid.Row="3" Style="{StaticResource FieldValue}" FontFamily="Consolas" Text="Enter"/>
                    <TextBlock Grid.Row="3" Grid.Column="1" Style="{StaticResource FieldName}" Text="Check the address"/>
                    <TextBlock Grid.Row="4" Style="{StaticResource FieldValue}" FontFamily="Consolas" Text="Ctrl+Shift+C"/>
                    <TextBlock Grid.Row="4" Grid.Column="1" Style="{StaticResource FieldName}" Text="Copy the result"/>
                    <TextBlock Grid.Row="5" Style="{StaticResource FieldValue}" FontFamily="Consolas" Text="Ctrl+Enter"/>
                    <TextBlock Grid.Row="5" Grid.Column="1" Style="{StaticResource FieldName}" Text="Start the batch"/>
                    <TextBlock Grid.Row="6" Style="{StaticResource FieldValue}" FontFamily="Consolas" Text="Ctrl+P"/>
                    <TextBlock Grid.Row="6" Grid.Column="1" Style="{StaticResource FieldName}" Text="Pause or resume the batch"/>
                    <TextBlock Grid.Row="7" Style="{StaticResource FieldValue}" FontFamily="Consolas" Text="Ctrl+Shift+S"/>
                    <TextBlock Grid.Row="7" Grid.Column="1" Style="{StaticResource FieldName}" Text="Stop the batch"/>
                    <TextBlock Grid.Row="8" Style="{StaticResource FieldValue}" FontFamily="Consolas" Text="Ctrl+R"/>
                    <TextBlock Grid.Row="8" Grid.Column="1" Style="{StaticResource FieldName}" Text="Retry failed lookups"/>
                    <TextBlock Grid.Row="9" Style="{StaticResource FieldValue}" FontFamily="Consolas" Text="Ctrl+G"/>
                    <TextBlock Grid.Row="9" Grid.Column="1" Style="{StaticResource FieldName}" Text="Group batch results by domain"/>
                    <TextBlock Grid.Row="10" Style="{StaticResource FieldValue}" FontFamily="Consolas" Text="Ctrl+F"/>
                    <TextBlock Grid.Row="10" Grid.Column="1" Style="{StaticResource FieldName}" Text="Search history"/>
                    <TextBlock Grid.Row="11" Style="{StaticResource FieldValue}" FontFamily="Consolas" Text="Ctrl+E"/>
                    <TextBlock Grid.Row="11" Grid.Column="1" Style="{StaticResource FieldName}" Text="Export history or batch results"/>
                    <TextBlock Grid.Row="12" Style="{StaticResource FieldValue}" FontFamily="Consolas" Text="Ctrl+D"/>
                    <TextBlock Grid.Row="12" Grid.Column="1" Style="{StaticResource FieldName}" Text="Dark mode on or off"/>
                    <TextBlock Grid.Row="13" Style="{StaticResource FieldValue}" FontFamily="Consolas" Text="F1"/>
                    <TextBlock Grid.Row="13" Grid.Column="1" Style="{StaticResource FieldName}" Text="About"/>
                    <TextBlock Grid.Row="14" Style="{StaticResource FieldValue}" FontFamily="Consolas" Text="Esc"/>
                    <TextBlock Grid.Row="14" Grid.Column="1" Style="{StaticResource FieldName}" Text="Close About, What's new or the import dialog"/>
                  </Grid>
                </StackPanel>
              </Border>
            </StackPanel>
          </Grid>
        </ScrollViewer>
      </TabItem>
    </TabControl>

    <!-- Status bar -->
    <Border Grid.Row="2" Background="{DynamicResource Surface}" BorderBrush="{DynamicResource Border}" BorderThickness="0,1,0,0" Padding="16,7">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <TextBlock x:Name="StatusText" Text="Ready" AutomationProperties.LiveSetting="Polite" AutomationProperties.Name="Status" TextTrimming="CharacterEllipsis" VerticalAlignment="Center"/>
        <StackPanel Grid.Column="1" Orientation="Horizontal">
          <TextBlock x:Name="StatusPrivate" Foreground="{DynamicResource CatError}" FontWeight="SemiBold" Margin="16,0,0,0" Visibility="Collapsed" Text="Private session"/>
          <TextBlock x:Name="StatusApi" Foreground="{DynamicResource Muted}" Margin="16,0,0,0"/>
          <TextBlock x:Name="StatusCache" Foreground="{DynamicResource Muted}" Margin="16,0,0,0"/>
          <TextBlock x:Name="StatusVersion" Foreground="{DynamicResource Muted}" Margin="16,0,0,0"/>
        </StackPanel>
      </Grid>
    </Border>

    <!-- CSV import: column mapping -->
    <Grid x:Name="ImportOverlay" Grid.RowSpan="3" Visibility="Collapsed" Background="#99000000">
      <Border Width="580" Margin="24" HorizontalAlignment="Center" VerticalAlignment="Center" Padding="24"
              Background="{DynamicResource Surface}" BorderBrush="{DynamicResource Border}" BorderThickness="1" CornerRadius="14">
        <StackPanel>
          <TextBlock Style="{StaticResource H2}" FontSize="17" Text="Import a CSV file"/>
          <TextBlock x:Name="ImportFileText" Style="{StaticResource Caption}" Margin="0,4,0,16"/>
          <TextBlock Style="{StaticResource FieldLabel}" Text="Column with the email addresses"/>
          <ComboBox x:Name="ImportEmailColumn" HorizontalAlignment="Stretch" AutomationProperties.Name="Column with the email addresses"/>
          <TextBlock Style="{StaticResource FieldLabel}" Margin="0,16,0,6" Text="Columns to keep with the results and in exports"/>
          <Border BorderBrush="{DynamicResource Border}" BorderThickness="1" CornerRadius="6" MaxHeight="220">
            <ScrollViewer VerticalScrollBarVisibility="Auto">
              <WrapPanel x:Name="ImportColumnsPanel" Margin="10,8" AutomationProperties.Name="Columns to keep"/>
            </ScrollViewer>
          </Border>
          <TextBlock x:Name="ImportPreviewText" Style="{StaticResource Caption}" Margin="0,10,0,0"/>
          <DockPanel Margin="0,20,0,0" LastChildFill="False">
            <Button x:Name="ImportTemplateButton" DockPanel.Dock="Left" Content="Save a template"
                    ToolTip="Save an example CSV with Email, EmployeeId, DisplayName and Department columns" AutomationProperties.HelpText="Save an example CSV with Email, EmployeeId, DisplayName and Department columns"/>
            <Button x:Name="ImportConfirmButton" DockPanel.Dock="Right" Style="{StaticResource PrimaryButton}" Content="Import" MinWidth="90"/>
            <Button x:Name="ImportCancelButton" DockPanel.Dock="Right" Content="Cancel" Margin="0,0,8,0"/>
          </DockPanel>
        </StackPanel>
      </Border>
    </Grid>

    <!-- What's new / release notes -->
    <Grid x:Name="ChangelogOverlay" Grid.RowSpan="3" Visibility="Collapsed" Background="#99000000">
      <Border Width="760" Margin="24" HorizontalAlignment="Center" VerticalAlignment="Center" Padding="24"
              Background="{DynamicResource Surface}" BorderBrush="{DynamicResource Border}" BorderThickness="1" CornerRadius="14">
        <Grid>
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>
          <TextBlock x:Name="ChangelogTitle" Style="{StaticResource H2}" FontSize="17" Text="What's new"/>
          <TextBox x:Name="ChangelogText" Grid.Row="1" Style="{StaticResource MonoBox}" TextWrapping="Wrap" Margin="0,12,0,0" MinHeight="320" MaxHeight="560"
                   AutomationProperties.Name="Changelog"/>
          <DockPanel Grid.Row="2" Margin="0,16,0,0" LastChildFill="False">
            <Button x:Name="ChangelogWebButton" DockPanel.Dock="Left" Content="Open on GitHub"/>
            <Button x:Name="ChangelogCloseButton" DockPanel.Dock="Right" Style="{StaticResource PrimaryButton}" Content="Close" MinWidth="90"/>
          </DockPanel>
        </Grid>
      </Border>
    </Grid>

    <!-- About -->
    <Grid x:Name="AboutOverlay" Grid.RowSpan="3" Visibility="Collapsed" Background="#99000000">
      <Border Width="580" Margin="24" HorizontalAlignment="Center" VerticalAlignment="Center"
              Background="{DynamicResource Surface}" BorderBrush="{DynamicResource Border}" BorderThickness="1" CornerRadius="14">
        <Border.Effect>
          <DropShadowEffect BlurRadius="32" ShadowDepth="0" Opacity="0.35"/>
        </Border.Effect>
        <Grid>
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>

          <Border Background="{DynamicResource HeaderBg}" CornerRadius="13,13,0,0" Padding="24,22">
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <Border Width="54" Height="54" CornerRadius="12" Background="{DynamicResource HeaderButtonBg}">
                <TextBlock FontFamily="Segoe MDL2 Assets" Text="&#xE77B;" FontSize="26" Foreground="{DynamicResource HeaderText}"
                           HorizontalAlignment="Center" VerticalAlignment="Center"/>
              </Border>
              <StackPanel Grid.Column="1" Margin="16,0,0,0" VerticalAlignment="Center">
                <TextBlock Text="Microsoft Account Checker" FontSize="20" FontWeight="SemiBold" Foreground="{DynamicResource HeaderText}"/>
                <TextBlock x:Name="AboutVersion" FontSize="13" Foreground="{DynamicResource HeaderMuted}" Margin="0,2,0,0"/>
              </StackPanel>
              <Button x:Name="AboutCloseX" Grid.Column="2" Style="{StaticResource HeaderButton}" VerticalAlignment="Top"
                      FontFamily="Segoe MDL2 Assets" Content="&#xE711;" Padding="9,6" ToolTip="Close" AutomationProperties.HelpText="Close"/>
            </Grid>
          </Border>

          <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
            <StackPanel Margin="24,20,24,10">
              <TextBlock TextWrapping="Wrap" LineHeight="20"
                         Text="Identifies the Microsoft account type behind an email address across the Commercial, GCC High and China clouds, with tenant and federation details."/>

              <TextBlock Style="{StaticResource H2}" Margin="0,20,0,10" Text="What is new in this version"/>
              <Grid Margin="0,0,0,8">
                <Grid.ColumnDefinitions><ColumnDefinition Width="18"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                <Ellipse Width="6" Height="6" Fill="{DynamicResource Primary}" VerticalAlignment="Top" Margin="0,7,0,0"/>
                <TextBlock Grid.Column="1" TextWrapping="Wrap" Text="Federated accounts are confirmed with a second lookup (GetUserRealm) instead of a single response field."/>
              </Grid>
              <Grid Margin="0,0,0,8">
                <Grid.ColumnDefinitions><ColumnDefinition Width="18"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                <Ellipse Width="6" Height="6" Fill="{DynamicResource Primary}" VerticalAlignment="Top" Margin="0,7,0,0"/>
                <TextBlock Grid.Column="1" TextWrapping="Wrap" Text="Throttling by Microsoft is detected and retried automatically. Results that stay throttled are flagged."/>
              </Grid>
              <Grid Margin="0,0,0,8">
                <Grid.ColumnDefinitions><ColumnDefinition Width="18"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                <Ellipse Width="6" Height="6" Fill="{DynamicResource Primary}" VerticalAlignment="Top" Margin="0,7,0,0"/>
                <TextBlock Grid.Column="1" TextWrapping="Wrap" Text="Organization name, tenant ID, tenant region and federation provider are shown for work and school domains."/>
              </Grid>
              <Grid Margin="0,0,0,8">
                <Grid.ColumnDefinitions><ColumnDefinition Width="18"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                <Ellipse Width="6" Height="6" Fill="{DynamicResource Primary}" VerticalAlignment="Top" Margin="0,7,0,0"/>
                <TextBlock Grid.Column="1" TextWrapping="Wrap" Text="New interface with live dark mode, batch checks, history and exports, plus an optional local API with a browser page and a system tray mode."/>
              </Grid>

              <Border Height="1" Background="{DynamicResource Border}" Margin="0,14,0,14"/>
              <Grid>
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="120"/>
                  <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <TextBlock Grid.Row="0" Style="{StaticResource FieldName}" Text="Author"/>
                <TextBlock Grid.Row="0" Grid.Column="1" Style="{StaticResource FieldValue}">
                  <Hyperlink x:Name="AboutAuthorLink" Foreground="{DynamicResource Link}"><Run x:Name="AboutAuthorRun"/></Hyperlink>
                </TextBlock>
                <TextBlock Grid.Row="1" Style="{StaticResource FieldName}" Text="Repository"/>
                <TextBlock Grid.Row="1" Grid.Column="1" Style="{StaticResource FieldValue}">
                  <Hyperlink x:Name="AboutRepoLink" Foreground="{DynamicResource Link}"><Run x:Name="AboutRepoRun"/></Hyperlink>
                </TextBlock>
                <TextBlock Grid.Row="2" Style="{StaticResource FieldName}" Text="License"/>
                <TextBlock Grid.Row="2" Grid.Column="1" x:Name="AboutLicense" Style="{StaticResource FieldValue}"/>
                <TextBlock Grid.Row="3" Style="{StaticResource FieldName}" Text="PowerShell"/>
                <TextBlock Grid.Row="3" Grid.Column="1" x:Name="AboutPsVersion" Style="{StaticResource FieldValue}"/>
                <TextBlock Grid.Row="4" Style="{StaticResource FieldName}" Text="Data folder"/>
                <TextBlock Grid.Row="4" Grid.Column="1" x:Name="AboutDataFolder" Style="{StaticResource FieldValue}"/>
                <TextBlock Grid.Row="5" Style="{StaticResource FieldName}" Text="Signature"/>
                <TextBlock Grid.Row="5" Grid.Column="1" x:Name="AboutSignature" Style="{StaticResource FieldValue}"/>
              </Grid>
            </StackPanel>
          </ScrollViewer>

          <Border Grid.Row="2" BorderBrush="{DynamicResource Border}" BorderThickness="0,1,0,0" Padding="24,14">
            <DockPanel LastChildFill="False">
              <Button x:Name="AboutRepoButton" DockPanel.Dock="Left" Style="{StaticResource PrimaryButton}" Content="View on GitHub"/>
              <Button x:Name="AboutIssueButton" DockPanel.Dock="Left" Content="Report an issue" Margin="8,0,0,0"/>
              <Button x:Name="AboutFolderButton" DockPanel.Dock="Left" Content="Open data folder" Margin="8,0,0,0"/>
              <Button x:Name="AboutChangelogButton" DockPanel.Dock="Left" Content="What's new" Margin="8,0,0,0"/>
              <Button x:Name="AboutCloseButton" DockPanel.Dock="Right" Content="Close" MinWidth="80"/>
            </DockPanel>
          </Border>
        </Grid>
      </Border>
    </Grid>
  </Grid>
</Window>
'@
#endregion

#region Web page served by the API at http://localhost:<port>/
$script:WebPageTemplate = @'
<!DOCTYPE html>
<html lang="en" data-app-theme="{{APP_THEME}}">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Microsoft Account Checker</title>
<style>
:root{--bg:#F4F7FC;--surface:#FFFFFF;--text:#1F2937;--muted:#5D6C83;--border:#D8E1EF;--primary:#3659C9;--primary-text:#FFFFFF;--secondary:#E8EEFF;--secondary-text:#243E9E;--input:#F8FAFD;--grid:#F8FAFD;--link:#3659C9;--h1:#3659C9;--h2:#4668D8;--entra:#15803D;--msa:#1D4ED8;--fed:#475569;--dual:#7E22CE;--gcc:#0E7490;--china:#A21CAF;--notms:#C62828;--err:#B45309;--unk:#5D6C83;color-scheme:light}
:root[data-theme="dark"]{--bg:#0F172A;--surface:#111B33;--text:#DCE7FF;--muted:#94A3B8;--border:#304264;--primary:#4C6FE0;--primary-text:#FFFFFF;--secondary:#243457;--secondary-text:#DCE7FF;--input:#16213D;--grid:#1B2745;--link:#8AB4FF;--h1:#14224A;--h2:#2A3F8F;--entra:#4ADE80;--msa:#60A5FA;--fed:#CBD5E1;--dual:#C084FC;--gcc:#22D3EE;--china:#E879F9;--notms:#F87171;--err:#FBBF24;--unk:#94A3B8;color-scheme:dark}
*{box-sizing:border-box}
[hidden]{display:none!important}
body{margin:0;font-family:"Segoe UI",system-ui,-apple-system,sans-serif;font-size:14px;line-height:1.45;background:var(--bg);color:var(--text)}
a{color:var(--link)}
.top{display:flex;align-items:center;justify-content:space-between;gap:16px;padding:14px 28px;background:linear-gradient(90deg,var(--h1),var(--h2));color:#FFFFFF}
.brand{display:flex;align-items:center;gap:12px}
.logo{width:38px;height:38px;border-radius:9px;background:rgba(255,255,255,.18);display:grid;place-items:center;font-weight:700;font-size:20px}
h1{font-size:17px;font-weight:600;margin:0}
.sub{font-size:12px;color:#F1F5FF}
.right{display:flex;align-items:center;gap:14px;flex-wrap:wrap}
.status{display:flex;align-items:center;gap:8px;font-size:13px}
.dot{width:9px;height:9px;border-radius:50%;background:#94A3B8}
.dot.ok{background:#22C55E}
.dot.bad{background:#F87171}
main{max-width:1280px;margin:0 auto;padding:20px 24px 8px}
.tabs{display:flex;gap:4px;border-bottom:1px solid var(--border);margin-bottom:16px;overflow-x:auto;scrollbar-width:thin}
.card{background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:18px;margin-bottom:16px}
label{display:block;font-size:12px;font-weight:600;color:var(--muted);margin-bottom:6px}
.row{display:flex;gap:8px;align-items:center}
.row.space{justify-content:space-between;flex-wrap:wrap;margin-top:10px}
.row input{flex:1}
input,textarea{width:100%;font:inherit;color:var(--text);background:var(--input);border:1px solid var(--border);border-radius:6px;padding:9px 11px}
input:focus,textarea:focus{outline:none;border-color:var(--primary);box-shadow:0 0 0 3px color-mix(in srgb,var(--primary) 22%,transparent)}
textarea{font-family:Consolas,monospace;font-size:13px;resize:vertical}
button{font:inherit;font-weight:600;cursor:pointer;border:0;border-radius:6px;padding:8px 14px;background:var(--secondary);color:var(--secondary-text)}
button:hover{opacity:.9}
button:disabled{opacity:.45;cursor:default}
button:focus-visible{outline:2px solid var(--primary);outline-offset:2px}
button.primary{background:var(--primary);color:var(--primary-text)}
button.danger{background:#D93025;color:#FFFFFF}
.seg{display:flex;gap:2px;padding:2px;border-radius:8px;background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.3)}
.seg button{background:none;color:#FFFFFF;padding:5px 11px;border-radius:6px;font-size:12px}
.seg button.on{background:#FFFFFF;color:#243E9E}
.stats{display:grid;grid-template-columns:repeat(8,minmax(0,1fr));gap:10px;margin-bottom:16px}
.stat{background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:12px 14px;min-width:0}
.stat i{display:block;width:28px;height:3px;border-radius:2px;background:var(--c);margin-bottom:8px}
.stat b{display:block;font-size:22px;font-weight:600}
.stat span{display:block;color:var(--muted);font-size:12px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.toolbar{display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-bottom:12px}
.toolbar input{flex:1 1 260px;min-width:0}
.toolbar button{flex:none}
.filters{display:grid;grid-template-columns:repeat(7,minmax(0,1fr));gap:8px;margin-bottom:12px}
.filters select,.filters button{width:100%;min-width:0}
select{font:inherit;color:var(--text);background:var(--input);border:1px solid var(--border);border-radius:6px;padding:7px 10px;cursor:pointer;text-overflow:ellipsis}
select:focus{outline:none;border-color:var(--primary);box-shadow:0 0 0 3px color-mix(in srgb,var(--primary) 22%,transparent)}
select.on{border-color:var(--primary);font-weight:600}
.tab{background:none;border-radius:0;border-bottom:3px solid transparent;color:var(--muted);padding:10px 14px;white-space:nowrap;flex:none}
.tab.active{color:var(--text);border-bottom-color:var(--primary)}
.panel{display:none}
.panel.active{display:block}
.error{color:var(--notms);margin-top:8px}
.result-head{display:flex;align-items:center;gap:8px 12px;flex-wrap:wrap}
.result-head h2{font-size:15px;margin:0}
.spacer{flex:1}
.badge{color:var(--c);background:color-mix(in srgb,var(--c) 15%,transparent);border-radius:12px;padding:3px 10px;font-size:12px;font-weight:600}
pre{font-family:Consolas,monospace;font-size:13px;background:var(--grid);border:1px solid var(--border);border-radius:8px;padding:12px 14px;margin:14px 0 0;white-space:pre-wrap;word-break:break-all}
.details{display:grid;grid-template-columns:170px 1fr;gap:6px 12px;margin:14px 0 0}
dt{color:var(--muted)}
dd{margin:0;font-weight:600;word-break:break-all}
.note{margin:12px 0 0;color:var(--muted)}
.note.warn{color:var(--err)}
.callout{margin:14px 0 0;padding:10px 14px;border-radius:8px;border-left:4px solid var(--c);background:color-mix(in srgb,var(--c) 12%,transparent)}
.callout b{display:block;color:var(--c);margin-bottom:2px}
.meta{color:var(--muted);font-size:12px;margin:10px 0 0}
.progress{height:6px;border-radius:3px;background:var(--secondary);margin-top:14px;overflow:hidden}
#bar{height:100%;width:0;background:var(--primary);transition:width .2s}
.table-wrap{overflow-x:auto}
table{border-collapse:collapse;width:100%;font-size:13px}
th{text-align:left;color:var(--muted);background:var(--grid);padding:8px 10px;border-bottom:1px solid var(--border);white-space:nowrap}
td{padding:7px 10px;border-bottom:1px solid var(--border);white-space:nowrap}
td.type{white-space:normal;min-width:170px}
td.st{color:var(--c);font-weight:600}
td.wrap{white-space:normal;min-width:200px}
td.type{color:var(--c);font-weight:600}
td.warn{color:var(--err);font-weight:600}
footer{max-width:1280px;margin:0 auto;padding:4px 24px 24px;color:var(--muted);font-size:12px}
@media (max-width:1180px){.filters{grid-template-columns:repeat(4,minmax(0,1fr))}}
@media (max-width:1100px){.stats{grid-template-columns:repeat(4,minmax(0,1fr))}}
@media (max-width:760px){.filters{grid-template-columns:repeat(2,minmax(0,1fr))}.filters button{grid-column:1 / -1}}
@media (max-width:640px){
  .top{flex-direction:column;align-items:flex-start;padding:12px 16px}
  main{padding:16px 12px 8px}
  footer{padding:4px 12px 20px}
  .card{padding:14px}
  .stats{grid-template-columns:repeat(2,minmax(0,1fr))}
  .details{grid-template-columns:1fr}
  .row{flex-wrap:wrap}
  .row input{flex:1 1 100%}
  .toolbar button{flex:1 1 auto}
}
@media (prefers-reduced-motion:reduce){#bar{transition:none}}
</style>
</head>
<body>
<header class="top">
  <div class="brand">
    <div class="logo">@</div>
    <div>
      <h1>Microsoft Account Checker</h1>
      <div class="sub">Commercial, GCC High and China clouds - version {{VERSION}}</div>
    </div>
  </div>
  <div class="right">
    <span class="status"><span id="dot" class="dot"></span><span id="statusText" role="status" aria-live="polite">Connecting...</span></span>
    <div class="seg" role="group" aria-label="Theme">
      <button type="button" data-mode="app" title="Use the same theme as the desktop app">Match app</button>
      <button type="button" data-mode="light">Light</button>
      <button type="button" data-mode="dark">Dark</button>
    </div>
  </div>
</header>

<main>
  <nav class="tabs">
    <button class="tab active" type="button" data-tab="single">Single check</button>
    <button class="tab" type="button" data-tab="batch">Batch check</button>
    <button class="tab" type="button" data-tab="history">History</button>
  </nav>

  <section id="single" class="panel active">
    <div class="card">
      <label for="email">Email address</label>
      <form id="singleForm" class="row" novalidate>
        <input id="email" type="email" placeholder="name@company.com" autocomplete="off" spellcheck="false">
        <button id="checkBtn" class="primary" type="submit">Check</button>
      </form>
      <div id="singleError" class="error" role="alert" hidden></div>
    </div>
    <div id="resultCard" class="card" hidden>
      <div class="result-head">
        <h2>Result</h2>
        <span id="badge" class="badge"></span>
        <span id="statusBadge" class="badge"></span>
        <span class="spacer"></span>
        <button id="copyBtn" type="button">Copy result</button>
      </div>
      <div id="callout" class="callout" hidden><b id="calloutTitle"></b><span id="calloutText"></span></div>
      <pre id="summary"></pre>
      <dl id="details" class="details"></dl>
      <p id="note" class="note" hidden></p>
      <p id="meta" class="meta"></p>
    </div>
  </section>

  <section id="batch" class="panel">
    <div class="card">
      <label for="list">Email list</label>
      <textarea id="list" rows="8" placeholder="One address per line. Commas, semicolons and spaces also work." spellcheck="false"></textarea>
      <div class="row space">
        <span id="listCount" class="meta">No addresses yet</span>
        <div class="row">
          <button id="clearBtn" type="button">Clear</button>
          <button id="exportBtn" type="button" disabled>Export CSV</button>
          <button id="runBtn" class="primary" type="button">Check all</button>
        </div>
      </div>
      <div id="progress" class="progress" hidden><div id="bar"></div></div>
      <div id="progressText" class="meta" role="status" aria-live="polite"></div>
    </div>
    <div id="tableCard" class="card" hidden>
      <div class="table-wrap">
        <table>
          <thead><tr><th>#</th><th>Username</th><th>Cloud</th><th>Account type</th><th>Status</th><th>Organization</th><th>Federation provider</th></tr></thead>
          <tbody id="rows"></tbody>
        </table>
      </div>
    </div>
  </section>

  <section id="history" class="panel">
    <div id="stats" class="stats"></div>
    <div class="card">
      <div class="toolbar">
        <input id="hsearch" type="search" placeholder="Search by username, cloud, account type, organization or origin" spellcheck="false">
        <button id="hrefresh" type="button">Refresh</button>
        <button id="hexport" type="button" disabled>Export CSV</button>
      </div>
      <div class="filters">
        <select id="ftype" aria-label="Account type">
          <option value="">All account types</option>
          <option value="Entra">Work/School</option>
          <option value="MSA">Personal</option>
          <option value="Federated">Federated</option>
          <option value="Dual">Dual identity</option>
          <option value="Sovereign">GCC High / China</option>
          <option value="NotMicrosoft">Not Microsoft</option>
          <option value="Error">Errors / unknown</option>
        </select>
        <select id="fcloud" aria-label="Cloud">
          <option value="">All clouds</option>
          <option value="Commercial">Commercial</option>
          <option value="GCCHigh">GCC High</option>
          <option value="China">China</option>
          <option value="N/A">Not found (N/A)</option>
        </select>
        <select id="fstatus" aria-label="Status">
          <option value="">All statuses</option>
          <option value="OK">Definitive</option>
          <option value="Inconclusive">Inconclusive</option>
          <option value="Throttled">Throttled</option>
          <option value="Unavailable">Temporarily unavailable</option>
          <option value="Invalid">Invalid address</option>
          <option value="Changed">Changed since previous check</option>
        </select>
        <select id="ffed" aria-label="Federation">
          <option value="">Federated or not</option>
          <option value="yes">Federated only</option>
          <option value="no">Not federated</option>
        </select>
        <select id="fsource" aria-label="Source">
          <option value="">Live and cache</option>
          <option value="Live">Live only</option>
          <option value="Cache">Cache only</option>
        </select>
        <select id="fdays" aria-label="Date">
          <option value="">All dates</option>
          <option value="0">Today</option>
          <option value="7">Last 7 days</option>
          <option value="30">Last 30 days</option>
          <option value="90">Last 90 days</option>
        </select>
        <button id="freset" type="button">Reset filters</button>
      </div>
      <div class="table-wrap">
        <table>
          <thead><tr><th>Time</th><th>Username</th><th>Cloud</th><th>Account type</th><th>Status</th><th>Changes</th><th>Organization</th><th>Source</th><th>Origin</th></tr></thead>
          <tbody id="hrows"></tbody>
        </table>
      </div>
      <p id="hmeta" class="meta" role="status" aria-live="polite"></p>
    </div>
  </section>
</main>

<footer>
  Checks made here also appear in the desktop app's History. History is cleared from the desktop app.
  <a href="/api/endpoints">API endpoints</a> -
  <a href="https://github.com/imumeshk/MS-Account-Checker" target="_blank" rel="noopener">GitHub</a>
</footer>

<script>
(function () {
  'use strict';
  // API token, filled in by the server when the page is opened from this computer.
  var TOKEN = '{{API_TOKEN}}';
  if (TOKEN.indexOf('{{') === 0) { TOKEN = ''; }
  var $ = function (id) { return document.getElementById(id); };
  var EMAIL = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/;
  var CATEGORIES = {
    Entra: ['--entra', 'Work/School (Entra ID)'], MSA: ['--msa', 'Personal (MSA)'],
    Federated: ['--fed', 'Federated (Work/School)'], Dual: ['--dual', 'Dual identity'],
    GCCHigh: ['--gcc', 'GCC High'], China: ['--china', 'M365 China'],
    NotMicrosoft: ['--notms', 'Not a Microsoft account'], Error: ['--err', 'Error'], Unknown: ['--unk', 'Unknown']
  };
  function category(c) { return CATEGORIES[c] || CATEGORIES.Unknown; }
  function paint(el, c) { el.style.setProperty('--c', 'var(' + category(c)[0] + ')'); }
  function text(v) { return (v === undefined || v === null) ? '' : String(v); }

  // Theme: "Match app" follows the desktop app's Dark mode switch; Light and Dark override it in this browser.
  var root = document.documentElement;
  var appTheme = root.getAttribute('data-app-theme') === 'dark' ? 'dark' : 'light';
  var mode = 'app';
  try { mode = localStorage.getItem('mac-theme-mode') || 'app'; } catch (e) { }
  var modeButtons = document.querySelectorAll('.seg button');
  function applyTheme() {
    root.setAttribute('data-theme', mode === 'app' ? appTheme : mode);
    modeButtons.forEach(function (b) { b.classList.toggle('on', b.getAttribute('data-mode') === mode); b.setAttribute('aria-pressed', b.getAttribute('data-mode') === mode ? 'true' : 'false'); });
  }
  modeButtons.forEach(function (b) {
    b.addEventListener('click', function () {
      mode = b.getAttribute('data-mode');
      try { localStorage.setItem('mac-theme-mode', mode); } catch (e) { }
      applyTheme();
    });
  });
  applyTheme();

  // Server status
  function ping() {
    getJson('/api/health?client=page', { cache: 'no-store' }).then(function (res) {
      if (res.status === 401) { $('dot').className = 'dot bad'; $('statusText').textContent = 'Token required'; return; }
      var j = res.data;
      $('dot').className = 'dot ok'; $('statusText').textContent = 'Connected';
      if (j && (j.theme === 'dark' || j.theme === 'light') && j.theme !== appTheme) { appTheme = j.theme; applyTheme(); }
    }).catch(function () {
      $('dot').className = 'dot bad'; $('statusText').textContent = 'Server not reachable';
    });
  }
  ping();
  setInterval(ping, 5000);

  // Tabs
  var tabs = document.querySelectorAll('.tab');
  tabs.forEach(function (tab) {
    tab.addEventListener('click', function () {
      tabs.forEach(function (t) { t.classList.toggle('active', t === tab); });
      document.querySelectorAll('.panel').forEach(function (p) { p.classList.toggle('active', p.id === tab.getAttribute('data-tab')); });
      if (tab.getAttribute('data-tab') === 'history') { loadHistory(); }
    });
  });

  var STATUS_KEYS = { OK: '--entra', Inconclusive: '--err', Throttled: '--err', Unavailable: '--notms', Error: '--notms', Invalid: '--notms' };
  function statusColor(el, status) { el.style.setProperty('--c', 'var(' + (STATUS_KEYS[status] || '--unk') + ')'); }
  function getJson(url, options) {
    options = options || {};
    var headers = options.headers || {};
    if (TOKEN) { headers['X-API-Key'] = TOKEN; }
    options.headers = headers;
    return fetch(url, options).then(function (r) {
      return r.text().then(function (body) {
        var data = null;
        try { data = body ? JSON.parse(body) : null; } catch (e) { data = null; }
        return { ok: r.ok, status: r.status, data: data };
      });
    });
  }

  function copyText(value, button) {
    var done = function () {
      var label = button.textContent; button.textContent = 'Copied';
      setTimeout(function () { button.textContent = label; }, 1200);
    };
    var fallback = function () {
      var area = document.createElement('textarea');
      area.value = value; document.body.appendChild(area); area.select();
      try { document.execCommand('copy'); done(); } catch (e) { }
      document.body.removeChild(area);
    };
    if (navigator.clipboard && navigator.clipboard.writeText) { navigator.clipboard.writeText(value).then(done, fallback); }
    else { fallback(); }
  }

  // Single check
  var last = null;
  var DETAILS = [['Organization', 'organization'], ['Federation provider', 'federationProvider'],
    ['Tenant ID', 'tenantId'], ['Tenant region', 'tenantRegion'], ['Domain type', 'domainType']];

  function showError(message) { var e = $('singleError'); e.textContent = message; e.hidden = false; }

  function render(d) {
    last = d;
    $('resultCard').hidden = false;
    var badge = $('badge');
    badge.textContent = category(d.category)[1];
    paint(badge, d.category);
    $('summary').textContent = text(d.output);
    var list = $('details');
    list.textContent = '';
    DETAILS.forEach(function (pair) {
      if (!d[pair[1]]) { return; }
      var dt = document.createElement('dt'); dt.textContent = pair[0];
      var dd = document.createElement('dd'); dd.textContent = text(d[pair[1]]);
      list.appendChild(dt); list.appendChild(dd);
    });
    list.hidden = list.children.length === 0;
    var sb = $('statusBadge');
    sb.textContent = d.statusLabel || d.status || '';
    statusColor(sb, d.status);
    var callout = $('callout');
    if (d.status && d.status !== 'OK') {
      statusColor(callout, d.status);
      $('calloutTitle').textContent = d.statusLabel || d.status;
      $('calloutText').textContent = [d.note, d.nextAction].filter(function (x) { return x; }).join(' ');
      callout.hidden = false;
    } else { callout.hidden = true; }
    var notes = [];
    if (d.category === 'NotMicrosoft' && d.status === 'OK') { notes.push('If you believe this result is incorrect, please confirm with your administrator that the email address is a valid Microsoft account and is set as the primary alias.'); }
    if ((d.status === 'Unavailable' || d.status === 'Error') && d.error) { notes.push('Details: ' + d.error); }
    var note = $('note');
    note.textContent = notes.join(' ');
    note.hidden = notes.length === 0;
    note.className = d.status === 'OK' ? 'note' : 'note warn';
    $('meta').textContent = 'Checked ' + text(d.checkedAt) + ', ' + (d.source === 'Cache' ? 'from the 24-hour cache.' : 'live from Microsoft.');
  }

  $('singleForm').addEventListener('submit', function (ev) {
    ev.preventDefault();
    var email = $('email').value.trim();
    if (!email) { showError('Please enter an email address.'); return; }
    if (!EMAIL.test(email)) { showError('Invalid email format. Please check and try again.'); return; }
    $('singleError').hidden = true;
    var button = $('checkBtn');
    button.disabled = true; button.textContent = 'Checking...';
    getJson('/api/check?email=' + encodeURIComponent(email)).then(function (res) {
      if (res.data && res.data.username !== undefined) { render(res.data); }
      else if (res.status === 401) { showError('The API requires a token. Open this page from the desktop app (API server tab, Open web page) so it is included automatically.'); }
      else { showError((res.data && res.data.error) || ('The server answered HTTP ' + res.status + '.')); }
    }).catch(function () {
      showError('The API server could not be reached. Make sure it is still running in the desktop app.');
    }).then(function () {
      button.disabled = false; button.textContent = 'Check';
    });
  });
  $('copyBtn').addEventListener('click', function () { if (last) { copyText(text(last.output), this); } });

  // Batch check
  function tokens(value) {
    var seen = {}, list = [];
    value.split(/[\s,;]+/).forEach(function (t) {
      t = t.replace(/^["'<>()\[\]]+|["'<>()\[\]]+$/g, '');
      if (!t) { return; }
      var key = t.toLowerCase();
      if (seen[key]) { return; }
      seen[key] = true; list.push(t);
    });
    return list;
  }
  function updateCount() {
    var list = tokens($('list').value);
    var valid = list.filter(function (t) { return EMAIL.test(t); }).length;
    var invalid = list.length - valid;
    $('listCount').textContent = list.length ? (valid + ' valid' + (invalid ? ', ' + invalid + ' invalid' : '')) : 'No addresses yet';
  }
  $('list').addEventListener('input', updateCount);
  $('clearBtn').addEventListener('click', function () { $('list').value = ''; updateCount(); });

  var results = [];
  var running = false;
  function progress(done, total) {
    $('bar').style.width = (total ? Math.round(done * 100 / total) : 0) + '%';
    $('progressText').textContent = done + ' of ' + total + ' checked';
  }
  function addRow(d) {
    results.push(d);
    var row = document.createElement('tr');
    [results.length, d.username, d.foundInCloud, d.accountType, d.statusLabel || d.status, d.organization, d.federationProvider].forEach(function (value, index) {
      var cell = document.createElement('td');
      cell.textContent = text(value);
      if (index === 3) { cell.className = 'type'; paint(cell, d.category); }
      if (index === 4 && d.status && d.status !== 'OK') { cell.className = 'st'; statusColor(cell, d.status); if (d.nextAction) { cell.title = d.nextAction; } }
      row.appendChild(cell);
    });
    $('rows').appendChild(row);
  }

  $('runBtn').addEventListener('click', function () {
    var button = this;
    if (running) { running = false; button.disabled = true; button.textContent = 'Stopping...'; return; }
    var list = tokens($('list').value);
    if (!list.length) { $('progressText').textContent = 'Add email addresses to the list first.'; return; }
    results = [];
    $('rows').textContent = '';
    $('tableCard').hidden = false;
    $('progress').hidden = false;
    $('exportBtn').disabled = true;
    running = true;
    button.textContent = 'Stop';
    button.classList.remove('primary'); button.classList.add('danger');
    var total = list.length, done = 0, size = 20, failed = false;
    progress(0, total);

    function finish() {
      running = false;
      button.disabled = false; button.textContent = 'Check all';
      button.classList.add('primary'); button.classList.remove('danger');
      $('exportBtn').disabled = results.length === 0;
      if (failed) { return; }
      $('progressText').textContent = done >= total ? ('Finished: ' + done + ' of ' + total + ' checked.') : ('Stopped after ' + done + ' of ' + total + '.');
    }
    function next(start) {
      if (!running || start >= total) { finish(); return; }
      var chunk = list.slice(start, start + size);
      getJson('/api/batch', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(chunk) }).then(function (res) {
        if (res.data && res.data.results) { res.data.results.forEach(addRow); }
        else {
          var reason = (res.data && res.data.error) || ('HTTP ' + res.status);
          chunk.forEach(function (e) { addRow({ username: e, accountType: reason, category: 'Error', status: 'Unavailable', statusLabel: res.status === 401 ? 'Token required' : 'Temporarily unavailable' }); });
        }
        done += chunk.length;
        progress(done, total);
        next(start + size);
      }).catch(function () {
        failed = true;
        $('progressText').textContent = 'The API server could not be reached. ' + done + ' of ' + total + ' were checked.';
        finish();
      });
    }
    next(0);
  });

  function downloadCsv(rows, columns, fileName) {
    var escape = function (v) { v = text(v); return /[",\r\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v; };
    var lines = [columns.join(',')].concat(rows.map(function (r) { return columns.map(function (c) { return escape(r[c]); }).join(','); }));
    var blob = new Blob(['\ufeff' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8' });
    var link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    setTimeout(function () { URL.revokeObjectURL(link.href); link.remove(); }, 500);
  }
  $('exportBtn').addEventListener('click', function () {
    downloadCsv(results, ['username', 'foundInCloud', 'cloudUrl', 'federationRedirectUrl', 'accountType', 'organization', 'tenantId',
      'tenantRegion', 'domainType', 'federationProvider', 'status', 'statusLabel', 'nextAction', 'note', 'source', 'checkedAt'], 'account_check_results.csv');
  });

  // History (read from the desktop app)
  var STATS = [['Total', 'Total checks', '--primary'], ['Entra', 'Work/School', '--entra'], ['MSA', 'Personal', '--msa'],
    ['Federated', 'Federated', '--fed'], ['Dual', 'Dual identity', '--dual'], ['Sovereign', 'GCC High / China', '--gcc'],
    ['NotMicrosoft', 'Not Microsoft', '--notms'], ['Error', 'Errors / unknown', '--err']];
  var historyItems = [];
  var historyTimer = null;
  function renderStats(counts) {
    var box = $('stats');
    box.textContent = '';
    STATS.forEach(function (s) {
      var card = document.createElement('div'); card.className = 'stat'; card.style.setProperty('--c', 'var(' + s[2] + ')');
      var bar = document.createElement('i');
      var value = document.createElement('b'); value.textContent = text(counts ? counts[s[0]] : 0);
      var label = document.createElement('span'); label.textContent = s[1];
      card.appendChild(bar); card.appendChild(value); card.appendChild(label);
      box.appendChild(card);
    });
  }
  function loadHistory() {
    var search = $('hsearch').value.trim();
    var query = '/api/history?limit=500&client=page&search=' + encodeURIComponent(search);
    var filtered = false;
    FILTERS.forEach(function (f) {
      var box = $(f[0]);
      box.classList.toggle('on', box.value !== '');
      if (box.value !== '') { query += '&' + f[1] + '=' + encodeURIComponent(box.value); filtered = true; }
    });
    getJson(query, { cache: 'no-store' }).then(function (res) {
      if (res.status === 401) { $('hmeta').textContent = 'The API requires a token. Open this page from the desktop app so it is included automatically.'; return; }
      if (!res.data || !res.data.items) { $('hmeta').textContent = 'History could not be loaded (HTTP ' + res.status + ').'; return; }
      var d = res.data;
      historyItems = d.items;
      renderStats(d.counts);
      var body = $('hrows');
      body.textContent = '';
      d.items.forEach(function (h) {
        var row = document.createElement('tr');
        [h.checkedAt, h.username, h.foundInCloud, h.accountType, h.statusLabel || h.status, h.changes, h.organization, h.source, h.origin].forEach(function (value, index) {
          var cell = document.createElement('td');
          cell.textContent = text(value);
          if (index === 3) { cell.className = 'type'; paint(cell, h.category); }
          if (index === 4 && h.status && h.status !== 'OK') { cell.className = 'st'; statusColor(cell, h.status); }
          if (index === 5 && value) { cell.className = 'st wrap'; cell.style.setProperty('--c', 'var(--dual)'); cell.title = value; }
          row.appendChild(cell);
        });
        body.appendChild(row);
      });
      $('hexport').disabled = d.items.length === 0;
      var narrowed = search || filtered;
      var shown = d.count < d.matched ? ('Showing the newest ' + d.count + ' of ' + d.matched + ' matching checks. Export CSV saves the rows shown.')
        : (narrowed ? (d.matched + ' of ' + d.total + (d.total === 1 ? ' check matches.' : ' checks match.')) : (d.total + (d.total === 1 ? ' check in history.' : ' checks in history.')));
      $('hmeta').textContent = d.total === 0 ? 'No checks in history yet.' : shown;
    }).catch(function () { $('hmeta').textContent = 'The API server could not be reached.'; });
  }
  var FILTERS = [['ftype', 'type'], ['fcloud', 'cloud'], ['fstatus', 'status'], ['ffed', 'federation'], ['fsource', 'source'], ['fdays', 'days']];
  FILTERS.forEach(function (f) { $(f[0]).addEventListener('change', loadHistory); });
  $('freset').addEventListener('click', function () {
    FILTERS.forEach(function (f) { $(f[0]).value = ''; });
    $('hsearch').value = '';
    loadHistory();
  });
  $('hrefresh').addEventListener('click', loadHistory);
  $('hsearch').addEventListener('input', function () { clearTimeout(historyTimer); historyTimer = setTimeout(loadHistory, 250); });
  $('hexport').addEventListener('click', function () {
    downloadCsv(historyItems, ['checkedAt', 'username', 'foundInCloud', 'cloudUrl', 'accountType', 'organization', 'tenantId',
      'tenantRegion', 'federationProvider', 'status', 'statusLabel', 'changes', 'source', 'origin'], 'account_check_history.csv');
  });
  setInterval(function () { if ($('history').classList.contains('active') && document.visibilityState === 'visible') { loadHistory(); } }, 10000);
  renderStats(null);

  $('email').focus();
})();
</script>
</body>
</html>

'@
#endregion

#region Theme
function Get-ThemePalette {
    param([bool]$Dark)
    if ($Dark) {
        return [ordered]@{
            Bg = '#0F172A'; Surface = '#111B33'; Text = '#DCE7FF'; Muted = '#94A3B8'; Border = '#304264'
            Primary = '#4C6FE0'; PrimaryText = '#FFFFFF'; Secondary = '#243457'; SecondaryText = '#DCE7FF'
            Danger = '#D93025'; DangerText = '#FFFFFF'; InputBg = '#16213D'; GridBg = '#1B2745'; GridAlt = '#14203B'
            Selection = '#2A3F75'; Link = '#8AB4FF'; HeaderText = '#FFFFFF'; HeaderMuted = '#AFC2EA'; SwitchOn = '#22C55E'
            HeaderStart = '#14224A'; HeaderEnd = '#2A3F8F'
            HeaderButtonBg = '#26FFFFFF'; HeaderButtonBorder = '#40FFFFFF'; SwitchTrack = '#33FFFFFF'; SwitchTrackBorder = '#66FFFFFF'; SwitchThumb = '#FFFFFF'
            CatEntra = '#4ADE80'; CatMSA = '#60A5FA'; CatFederated = '#CBD5E1'; CatDual = '#C084FC'
            CatGCCHigh = '#22D3EE'; CatChina = '#E879F9'; CatNotMicrosoft = '#F87171'; CatError = '#FBBF24'; CatUnknown = '#94A3B8'
        }
    }
    return [ordered]@{
        Bg = '#F4F7FC'; Surface = '#FFFFFF'; Text = '#1F2937'; Muted = '#5D6C83'; Border = '#D8E1EF'
        Primary = '#3659C9'; PrimaryText = '#FFFFFF'; Secondary = '#E8EEFF'; SecondaryText = '#243E9E'
        Danger = '#D93025'; DangerText = '#FFFFFF'; InputBg = '#F8FAFD'; GridBg = '#F8FAFD'; GridAlt = '#F4F7FC'
        Selection = '#DCE5FF'; Link = '#3659C9'; HeaderText = '#FFFFFF'; HeaderMuted = '#F1F5FF'; SwitchOn = '#22C55E'
        HeaderStart = '#3659C9'; HeaderEnd = '#4668D8'
        HeaderButtonBg = '#26FFFFFF'; HeaderButtonBorder = '#40FFFFFF'; SwitchTrack = '#33FFFFFF'; SwitchTrackBorder = '#66FFFFFF'; SwitchThumb = '#FFFFFF'
        CatEntra = '#15803D'; CatMSA = '#1D4ED8'; CatFederated = '#475569'; CatDual = '#7E22CE'
        CatGCCHigh = '#0E7490'; CatChina = '#A21CAF'; CatNotMicrosoft = '#C62828'; CatError = '#B45309'; CatUnknown = '#5D6C83'
    }
}

function New-FrozenBrush {
    param($Color)
    $brush = [System.Windows.Media.SolidColorBrush]::new($Color)
    $brush.Freeze()
    return $brush
}

function Set-Theme {
    param([bool]$Dark)
    $highContrast = Test-HighContrast
    $palette   = if ($highContrast) { Get-HighContrastPalette } else { Get-ThemePalette -Dark $Dark }
    $resources = $script:Window.Resources
    foreach ($key in @($palette.Keys)) {
        if ($key -eq 'HeaderStart' -or $key -eq 'HeaderEnd') { continue }
        $color = [System.Windows.Media.ColorConverter]::ConvertFromString($palette[$key])
        $resources[$key] = New-FrozenBrush $color
        if ($key -like 'Cat*') {
            $alpha = if ($highContrast) { 0 } elseif ($Dark) { 56 } else { 31 }
            $tint  = [System.Windows.Media.Color]::FromArgb([byte]$alpha, $color.R, $color.G, $color.B)
            $resources["$($key)Bg"] = New-FrozenBrush $tint
        }
    }
    $start = [System.Windows.Media.ColorConverter]::ConvertFromString($palette['HeaderStart'])
    $end   = [System.Windows.Media.ColorConverter]::ConvertFromString($palette['HeaderEnd'])
    $gradient = [System.Windows.Media.LinearGradientBrush]::new($start, $end, 0.0)
    $gradient.Freeze()
    $resources['HeaderBg'] = $gradient
    $script:Settings.DarkMode = $Dark
    $effectiveDark = $Dark -and -not $highContrast
    if ($null -ne $script:Shared) { $script:Shared['Theme'] = if ($effectiveDark) { 'dark' } else { 'light' } }
    Set-TitleBarTheme -Dark $effectiveDark
    if ($script:UI -and $script:UI.DarkToggle) {
        $script:UI.DarkToggle.IsEnabled = -not $highContrast
        $script:UI.DarkToggle.ToolTip = if ($highContrast) { 'Windows high contrast is on, so the app uses your high contrast colours.' } else { 'Switch between light and dark (Ctrl+D)' }
    }
}

function Set-TitleBarTheme {
    param([bool]$Dark)
    if (-not $script:Hwnd -or -not ('MsAccCheckerNative3' -as [type])) { return }
    $value = [int]$Dark
    try {
        $hr = [MsAccCheckerNative3]::DwmSetWindowAttribute($script:Hwnd, 20, [ref]$value, 4)
        if ($hr -ne 0) { [void][MsAccCheckerNative3]::DwmSetWindowAttribute($script:Hwnd, 19, [ref]$value, 4) }
    } catch { }
}
#endregion

#region Settings, history, cache, log files
function Read-Settings {
    $settings = @{
        EnableLogging = $false; DarkMode = $false; ApiPort = 8080; OpenBrowserOnApiStart = $true; TrayOnClose = $true
        ApiRequireToken = $false; ApiToken = ''; HistoryRetentionDays = 0; LogRetentionDays = 0; ClearOnExit = $false
        BatchDelayMs = 200; BatchConcurrency = 1; BatchSkipCache = $false
        StableCacheHours = 24; InconclusiveCacheHours = 24; FailureCacheMinutes = 0; StaleAfterHours = 12; ApiRateLimitPerMinute = 120
        EncryptStoredData = $false; RedactExports = $false; WatchFolder = ''; WatchFolderEnabled = $false; CheckUpdatesOnStart = $false; CompactTables = $false
        WindowLeft = $null; WindowTop = $null; WindowWidth = $null; WindowHeight = $null; WindowMaximized = $false
    }
    if (-not (Test-Path -LiteralPath $script:ConfigFile)) { return $settings }
    try {
        $text = Get-Content -LiteralPath $script:ConfigFile -Raw
        if ([string]::IsNullOrWhiteSpace($text)) { return $settings }
        $data = ConvertFrom-StringData -StringData $text
        foreach ($name in @('EnableLogging', 'DarkMode', 'WindowMaximized', 'OpenBrowserOnApiStart', 'TrayOnClose', 'ApiRequireToken', 'ClearOnExit', 'BatchSkipCache',
                            'EncryptStoredData', 'RedactExports', 'WatchFolderEnabled', 'CheckUpdatesOnStart', 'CompactTables')) {
            $flag = $false
            if ($data.ContainsKey($name) -and [bool]::TryParse([string]$data[$name], [ref]$flag)) { $settings[$name] = $flag }
        }
        $ranges = @{
            ApiPort = @(1, 65535); HistoryRetentionDays = @(0, 3650); LogRetentionDays = @(0, 3650); BatchDelayMs = @(0, 10000); BatchConcurrency = @(1, 4)
            StableCacheHours = @(0, 720); InconclusiveCacheHours = @(0, 720); FailureCacheMinutes = @(0, 60); StaleAfterHours = @(1, 720); ApiRateLimitPerMinute = @(0, 6000)
        }
        foreach ($name in @($ranges.Keys)) {
            $number = 0
            if ($data.ContainsKey($name) -and [int]::TryParse([string]$data[$name], [ref]$number)) {
                $settings[$name] = [Math]::Max($ranges[$name][0], [Math]::Min($ranges[$name][1], $number))
            }
        }
        if ($data.ContainsKey('ApiToken') -and [string]$data['ApiToken'] -match '^[0-9a-f]{16,128}$') { $settings.ApiToken = [string]$data['ApiToken'] }
        if ($data.ContainsKey('WatchFolder')) { $settings.WatchFolder = [string]$data['WatchFolder'] }
        foreach ($name in @('WindowLeft', 'WindowTop', 'WindowWidth', 'WindowHeight')) {
            $number = 0.0
            if ($data.ContainsKey($name) -and [double]::TryParse([string]$data[$name], [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$number)) {
                $settings[$name] = $number
            }
        }
    } catch {
        Write-Warning "Could not read '$($script:ConfigFile)'. Default settings are used. $($_.Exception.Message)"
    }
    return $settings
}

function Save-Settings {
    $inv = [System.Globalization.CultureInfo]::InvariantCulture
    $s = $script:Settings
    $lines = @()
    foreach ($name in @('EnableLogging', 'DarkMode', 'ApiPort', 'OpenBrowserOnApiStart', 'TrayOnClose', 'ApiRequireToken', 'ApiToken',
                        'HistoryRetentionDays', 'LogRetentionDays', 'ClearOnExit', 'BatchDelayMs', 'BatchConcurrency', 'BatchSkipCache',
                        'StableCacheHours', 'InconclusiveCacheHours', 'FailureCacheMinutes', 'StaleAfterHours', 'ApiRateLimitPerMinute',
                        'EncryptStoredData', 'RedactExports', 'WatchFolderEnabled', 'CheckUpdatesOnStart', 'CompactTables', 'WindowMaximized')) {
        $lines += "$name=$($s[$name])"
    }
    # ConvertFrom-StringData treats a backslash as an escape character, so paths are written with doubled backslashes.
    $lines += 'WatchFolder=' + ([string]$s.WatchFolder).Replace('\', '\\')
    foreach ($name in @('WindowLeft', 'WindowTop', 'WindowWidth', 'WindowHeight')) {
        if ($null -ne $s[$name]) { $lines += ('{0}={1}' -f $name, ([double]$s[$name]).ToString('0.##', $inv)) }
    }
    try {
        Set-Content -LiteralPath $script:ConfigFile -Value $lines -Encoding UTF8
    } catch {
        Set-Status "Could not save settings: $($_.Exception.Message)" 'Error'
    }
}

function New-Counts {
    return @{ Total = 0; Entra = 0; MSA = 0; Federated = 0; Dual = 0; Sovereign = 0; NotMicrosoft = 0; Error = 0; Throttled = 0; Inconclusive = 0; Unavailable = 0; Changed = 0 }
}

function Get-Bucket {
    param([string]$Category)
    switch ($Category) {
        'Entra'        { return 'Entra' }
        'MSA'          { return 'MSA' }
        'Federated'    { return 'Federated' }
        'Dual'         { return 'Dual' }
        'GCCHigh'      { return 'Sovereign' }
        'China'        { return 'Sovereign' }
        'NotMicrosoft' { return 'NotMicrosoft' }
        default        { return 'Error' }
    }
}

function Add-ToCounts {
    param([hashtable]$Counts, [string]$Category, [string]$Status = '', [int]$Delta = 1)
    $Counts['Total'] = $Counts['Total'] + $Delta
    $bucket = Get-Bucket $Category
    $Counts[$bucket] = $Counts[$bucket] + $Delta
    if ($Status -eq 'Error') { $Status = 'Unavailable' }
    if ($Status -eq 'Throttled' -or $Status -eq 'Inconclusive' -or $Status -eq 'Unavailable') { $Counts[$Status] = $Counts[$Status] + $Delta }
}

function Read-History {
    $script:HistoryRows.Clear()
    $script:Shared.History.Clear()
    $script:HistoryCounts = New-Counts
    $script:LatestByUser = @{}
    if (-not (Test-Path -LiteralPath $script:HistoryFile)) { return }
    try {
        $text = Read-DataFile -Path $script:HistoryFile
        if ([string]::IsNullOrWhiteSpace($text)) { return }
        $data = ConvertFrom-Json -InputObject $text
        $loaded = 0
        foreach ($entry in $data) {
            if ($null -eq $entry) { continue }
            if ($loaded -ge $script:HistoryLimit) { break }
            if ($null -eq $entry.PSObject.Properties['StatusLabel']) {
                $advice = Get-StatusAdvice -Status ([string]$entry.Status)
                $entry | Add-Member -NotePropertyName 'StatusLabel' -NotePropertyValue $advice.Label -Force
                $entry | Add-Member -NotePropertyName 'NextAction' -NotePropertyValue $advice.Action -Force
            }
            $script:HistoryRows.Add($entry)
            Add-ToCounts -Counts $script:HistoryCounts -Category ([string]$entry.Category) -Status ([string]$entry.Status)
            if ([string]$entry.Changes) { $script:HistoryCounts['Changed'] = $script:HistoryCounts['Changed'] + 1 }
            $key = ([string]$entry.Username).ToLowerInvariant()
            if (([string]$entry.Status -eq 'OK' -or [string]$entry.Status -eq 'Inconclusive') -and -not $script:LatestByUser.ContainsKey($key)) { $script:LatestByUser[$key] = $entry }
            $loaded++
        }
        for ($i = $script:HistoryRows.Count - 1; $i -ge 0; $i--) { [void]$script:Shared.History.Add($script:HistoryRows[$i]) }
    } catch {
        Set-Status 'The history file could not be read (it may be encrypted for another Windows account), so history starts empty.' 'Warn'
    }
}

function Save-History {
    try {
        $items = @($script:HistoryRows | Where-Object { -not $_.Private } | Select-Object -First $script:HistoryLimit)
        Write-DataFile -Path $script:HistoryFile -Text (ConvertTo-Json -InputObject $items -Depth 3)
        $script:HistoryDirty = $false
    } catch {
        Set-Status "Could not save history: $($_.Exception.Message)" 'Error'
    }
}

function Read-Cache {
    if (-not (Test-Path -LiteralPath $script:CacheFile)) { return }
    try {
        $text = Read-DataFile -Path $script:CacheFile
        if ([string]::IsNullOrWhiteSpace($text)) { return }
        foreach ($entry in (ConvertFrom-Json -InputObject $text)) {
            if ($null -eq $entry -or -not $entry.Key -or $null -eq $entry.Result) { continue }
            try {
                if ($entry.Time -is [datetime]) {
                    $time = $entry.Time.ToUniversalTime()
                } else {
                    $time = [datetime]::Parse([string]$entry.Time, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::RoundtripKind).ToUniversalTime()
                }
                $ttl = if ($null -ne $entry.TtlHours) { [double]$entry.TtlHours } else { [double]$script:Settings.StableCacheHours }
                if (([datetime]::UtcNow - $time).TotalHours -lt $ttl) {
                    $script:Shared.Cache[[string]$entry.Key] = @{ Time = $time; Result = $entry.Result; TtlHours = $ttl }
                }
            } catch { }
        }
    } catch { }
}

function Save-Cache {
    try {
        $now  = [datetime]::UtcNow
        $list = New-Object System.Collections.Generic.List[object]
        $cache = $script:Shared.Cache
        [System.Threading.Monitor]::Enter($cache.SyncRoot)
        try {
            foreach ($key in @($cache.Keys)) {
                if ($script:PrivateKeys.Contains([string]$key)) { continue }
                $entry = $cache[$key]
                if ($null -eq $entry) { continue }
                $ttl = if ($entry.ContainsKey('TtlHours') -and $null -ne $entry['TtlHours']) { [double]$entry['TtlHours'] } else { [double]$script:CacheHours }
                if (($now - $entry.Time).TotalHours -lt $ttl) {
                    $list.Add([pscustomobject]@{ Key = $key; Time = $entry.Time.ToString('o'); TtlHours = $ttl; Result = $entry.Result })
                }
            }
        } finally {
            [System.Threading.Monitor]::Exit($cache.SyncRoot)
        }
        Write-DataFile -Path $script:CacheFile -Text (ConvertTo-Json -InputObject $list.ToArray() -Depth 4)
    } catch { }
}

function Clear-LookupCache {
    $cache = $script:Shared.Cache
    [System.Threading.Monitor]::Enter($cache.SyncRoot)
    try { $cache.Clear() } finally { [System.Threading.Monitor]::Exit($cache.SyncRoot) }
    $domainCache = $script:Shared.DomainCache
    [System.Threading.Monitor]::Enter($domainCache.SyncRoot)
    try { $domainCache.Clear() } finally { [System.Threading.Monitor]::Exit($domainCache.SyncRoot) }
    Save-Cache
    Set-Status 'Lookup cache cleared. The next checks will contact Microsoft again.' 'Success'
}

function Get-TodayLogFile {
    return (Join-Path $script:LogsDir ('MS_Account_Checker_Logs_{0}.txt' -f (Get-Date -Format 'yyyy-MM-dd')))
}

function Write-AccountLog {
    param($Result)
    if (-not $script:Settings.EnableLogging -or $script:PrivateSession) { return }
    try {
        $file = Get-TodayLogFile
        if (-not (Test-Path -LiteralPath $file)) {
            $script:LogHeader | Out-File -FilePath $file -Encoding UTF8
        }
        $fields = @(
            $Result.Timestamp, $Result.Username, $Result.FoundInCloud, $Result.CloudURL, $Result.FederationRedirectUrl,
            $Result.AccountType, $Result.Organization, $Result.TenantId, $Result.TenantRegion, $Result.FederationProvider, $Result.Status
        ) | ForEach-Object { ([string]$_) -replace '[\t\r\n]+', ' ' }
        ($fields -join "`t") | Out-File -FilePath $file -Append -Encoding UTF8
    } catch {
        if (-not $script:LogErrorShown) {
            $script:LogErrorShown = $true
            Set-Status "Logging failed: $($_.Exception.Message)" 'Error'
        }
    }
}

#endregion

#region UI helpers
function Get-InnerMessage {
    param($Exception)
    $e = $Exception
    while ($null -ne $e.InnerException) { $e = $e.InnerException }
    return $e.Message
}

function Set-Status {
    param([string]$Text, [string]$Kind = 'Info')
    if (-not $script:UI -or -not $script:UI.StatusText) { return }
    $first = ($Text -split "`r?`n")[0]
    $script:UI.StatusText.Text = $first
    $key = switch ($Kind) {
        'Error'   { 'CatNotMicrosoft' }
        'Success' { 'CatEntra' }
        'Warn'    { 'CatError' }
        default   { 'Text' }
    }
    $script:UI.StatusText.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, $key)
}

function Show-AppError {
    param($ErrorObject, [string]$Context = '')
    if ($ErrorObject -is [System.Management.Automation.ErrorRecord]) { $message = Get-InnerMessage $ErrorObject.Exception }
    elseif ($ErrorObject -is [System.Exception]) { $message = Get-InnerMessage $ErrorObject }
    else { $message = [string]$ErrorObject }
    if ($Context) { $message = "$Context`n`n$message" }
    Set-Status $message 'Error'
    try {
        if ($script:Window -and $script:Window.IsVisible) {
            [void][System.Windows.MessageBox]::Show($script:Window, $message, $script:AppName, [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        } else {
            [void][System.Windows.MessageBox]::Show($message, $script:AppName, [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        }
    } catch { }
}

function Set-ClipboardText {
    param([string]$Text)
    for ($i = 0; $i -lt 5; $i++) {
        try {
            [System.Windows.Clipboard]::SetText($Text)
            # The clipboard watcher must not collect what the app itself copied.
            try { $script:ClipboardSeq = [MsAccCheckerNative3]::GetClipboardSequenceNumber() } catch { }
            return $true
        } catch { Start-Sleep -Milliseconds 60 }
    }
    return $false
}

function Get-ClipboardText {
    try { return [string][System.Windows.Clipboard]::GetText() } catch { return '' }
}

function Set-ToggleSilently {
    param($Toggle, [bool]$Value)
    $script:SuppressToggle = $true
    try { $Toggle.IsChecked = $Value } finally { $script:SuppressToggle = $false }
}

function Open-Url {
    param($Uri)
    if ($null -eq $Uri) { return }
    try { Start-Process ([string]$Uri) } catch { Show-AppError $_ "Could not open $Uri" }
}

function Set-HintVisibility {
    param($TextBox, $Hint)
    $Hint.Visibility = if ([string]::IsNullOrEmpty($TextBox.Text)) { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }
}

function Set-WindowBounds {
    $w  = $script:Window
    $wa = [System.Windows.SystemParameters]::WorkArea
    $w.MinWidth  = [Math]::Min(960, $wa.Width)
    $w.MinHeight = [Math]::Min(620, $wa.Height)

    $width  = [Math]::Max($w.MinWidth,  [Math]::Min(1280, $wa.Width * 0.92))
    $height = [Math]::Max($w.MinHeight, [Math]::Min(840,  $wa.Height * 0.92))
    $left   = $wa.Left + ($wa.Width - $width) / 2
    $top    = $wa.Top + ($wa.Height - $height) / 2

    $s = $script:Settings
    if ($s.WindowWidth -and $s.WindowHeight -and $null -ne $s.WindowLeft -and $null -ne $s.WindowTop) {
        $vLeft   = [System.Windows.SystemParameters]::VirtualScreenLeft
        $vTop    = [System.Windows.SystemParameters]::VirtualScreenTop
        $vWidth  = [System.Windows.SystemParameters]::VirtualScreenWidth
        $vHeight = [System.Windows.SystemParameters]::VirtualScreenHeight
        $onScreen = ($s.WindowLeft -ge ($vLeft - 50)) -and ($s.WindowTop -ge ($vTop - 10)) -and
                    (($s.WindowLeft + 120) -le ($vLeft + $vWidth)) -and (($s.WindowTop + 60) -le ($vTop + $vHeight))
        if ($onScreen) {
            $width  = [Math]::Max($w.MinWidth,  [Math]::Min($s.WindowWidth,  $vWidth))
            $height = [Math]::Max($w.MinHeight, [Math]::Min($s.WindowHeight, $vHeight))
            $left   = $s.WindowLeft
            $top    = $s.WindowTop
        }
    }
    $w.WindowStartupLocation = [System.Windows.WindowStartupLocation]::Manual
    $w.Width  = $width
    $w.Height = $height
    $w.Left   = $left
    $w.Top    = $top
    if ($s.WindowMaximized) { $w.WindowState = [System.Windows.WindowState]::Maximized }
}

function Save-WindowBounds {
    $w = $script:Window
    if ($w.WindowState -eq [System.Windows.WindowState]::Normal) {
        $bounds = [System.Windows.Rect]::new($w.Left, $w.Top, $w.ActualWidth, $w.ActualHeight)
    } else {
        $bounds = $w.RestoreBounds
    }
    if (-not $bounds.IsEmpty -and $bounds.Width -gt 0) {
        $script:Settings.WindowLeft   = $bounds.Left
        $script:Settings.WindowTop    = $bounds.Top
        $script:Settings.WindowWidth  = $bounds.Width
        $script:Settings.WindowHeight = $bounds.Height
    }
    $script:Settings.WindowMaximized = ($w.WindowState -eq [System.Windows.WindowState]::Maximized)
}
#endregion

#region Background job plumbing
function New-WorkerShell {
    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.ThreadOptions = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
    $runspace.Open()
    $shell = [powershell]::Create()
    $shell.Runspace = $runspace
    return $shell
}

function Close-WorkerShell {
    param($Shell, [int]$WaitMs = 0, $Handle = $null)
    if ($null -eq $Shell) { return }
    try {
        if ($null -ne $Handle -and -not $Handle.IsCompleted) {
            if ($WaitMs -gt 0) { [void]$Handle.AsyncWaitHandle.WaitOne($WaitMs) }
            if (-not $Handle.IsCompleted) { [void]$Shell.BeginStop($null, $null); return }
        }
        $runspace = $Shell.Runspace
        $Shell.Dispose()
        if ($null -ne $runspace) { $runspace.Dispose() }
    } catch { }
}

function Start-CheckJob {
    param([string]$Mode, [string[]]$Emails, [int]$DelayMs = 0, [int]$Workers = 1, [bool]$SkipCache = $false, $RetryIndex = $null, [double]$MaxAgeHours = 0)
    $emailQueue = New-Object 'System.Collections.Concurrent.ConcurrentQueue[string]'
    foreach ($address in $Emails) { $emailQueue.Enqueue($address) }
    $queue   = New-Object 'System.Collections.Concurrent.ConcurrentQueue[object]'
    $sync    = [hashtable]::Synchronized(@{ Cancel = $false; Paused = $false })
    $count   = [Math]::Max(1, [Math]::Min($Workers, $Emails.Count))
    $shells  = New-Object System.Collections.Generic.List[object]
    $handles = New-Object System.Collections.Generic.List[object]
    for ($i = 0; $i -lt $count; $i++) {
        $shell = New-WorkerShell
        [void]$shell.AddScript($script:WorkerScript.ToString())
        [void]$shell.AddParameter('CoreText', $script:CoreText)
        [void]$shell.AddParameter('EmailQueue', $emailQueue)
        [void]$shell.AddParameter('Queue', $queue)
        [void]$shell.AddParameter('Sync', $sync)
        [void]$shell.AddParameter('Cache', $script:Shared.Cache)
        [void]$shell.AddParameter('DomainCache', $script:Shared.DomainCache)
        [void]$shell.AddParameter('Stats', $script:Shared.Stats)
        [void]$shell.AddParameter('DelayMs', $DelayMs)
        [void]$shell.AddParameter('TtlHours', [double]$script:CacheHours)
        [void]$shell.AddParameter('SkipCache', $SkipCache)
        [void]$shell.AddParameter('Policy', $script:Shared['CachePolicy'])
        [void]$shell.AddParameter('MaxAgeHours', $MaxAgeHours)
        $shells.Add($shell)
        $handles.Add($shell.BeginInvoke())
    }
    $script:Jobs[$Mode] = @{
        Shells = $shells; Handles = $handles; Queue = $queue; Sync = $sync; Workers = $count
        Total = $Emails.Count; Done = 0; FromCache = 0; Changed = 0; Started = Get-Date; RetryIndex = $RetryIndex
        Origin = ''; WatchFile = ''
    }
}

function Complete-CheckJob {
    param([string]$Mode)
    $job = $script:Jobs[$Mode]
    $script:Jobs.Remove($Mode)
    $failure = $null
    for ($i = 0; $i -lt $job.Shells.Count; $i++) {
        try { [void]$job.Shells[$i].EndInvoke($job.Handles[$i]) } catch { if (-not $failure) { $failure = Get-InnerMessage $_.Exception } }
        if (-not $failure -and $job.Shells[$i].Streams.Error.Count -gt 0) { $failure = $job.Shells[$i].Streams.Error[0].ToString() }
        Close-WorkerShell -Shell $job.Shells[$i]
    }

    if ($Mode -eq 'Single') {
        Set-SingleBusy $false
        if ($failure) { Set-Status "The check could not finish: $failure" 'Error' }
    } elseif ($Mode -eq 'Batch') {
        Set-BatchBusy $false
        Update-BatchProgress -Job $job
        if ($script:BatchGrouped) { Update-DomainView }
        $what    = if ($null -ne $job.RetryIndex) { 'Retry' } elseif ($job.WatchFile) { 'Watched file' } else { 'Batch' }
        $state   = if ($job.Sync.Cancel) { 'stopped' } else { 'finished' }
        $seconds = [int]((Get-Date) - $job.Started).TotalSeconds
        $failed  = $script:BatchCounts['Throttled'] + $script:BatchCounts['Unavailable']
        if ($failure) {
            Set-Status "$what stopped by an error after $($job.Done) of $($job.Total): $failure" 'Error'
        } else {
            $tail = ''
            if ($job.Changed -gt 0) { $tail += " $($job.Changed) changed since their previous check." }
            if ($failed -gt 0) { $tail += " $failed still need a retry." }
            $kind = if ($failed -gt 0 -or $job.Changed -gt 0) { 'Warn' } else { 'Success' }
            Set-Status ("{0} {1}: {2} of {3} checked, {4} from cache, {5}s.{6}" -f $what, $state, $job.Done, $job.Total, $job.FromCache, $seconds, $tail) $kind
        }
        if ($job.WatchFile) { Complete-WatchFile -Job $job }
    }
}

function Receive-CheckResult {
    param([string]$Mode, $Result, $Job)
    if ($Result.Source -eq 'Cache') { $Job.FromCache = $Job.FromCache + 1 }
    $changes = Set-ResultChanges -Result $Result
    if ($changes) { $Job.Changed = [int]$Job.Changed + 1 }
    if ($Mode -eq 'Single') {
        Show-SingleResult -Result $Result
        Show-ChangeCallout -Result $Result
        Add-HistoryEntry -Result $Result -Origin 'Single'
        if ($changes) { Set-Status ("{0} changed since the previous check: {1}" -f $Result.Username, $changes) 'Warn' }
    } else {
        $replaced = $false
        if ($null -ne $Job.RetryIndex) {
            $key = ([string]$Result.Username).ToLowerInvariant()
            if ($Job.RetryIndex.ContainsKey($key)) {
                Update-BatchRow -Result $Result -Index ([int]$Job.RetryIndex[$key])
                $replaced = $true
            }
        }
        if (-not $replaced) { Add-BatchRow -Result $Result }
        $origin = if ($Job.Origin) { [string]$Job.Origin } else { 'Batch' }
        Add-HistoryEntry -Result $Result -Origin $origin
    }
    if ($Result.Status -ne 'Invalid') { Write-AccountLog -Result $Result }
}

function Invoke-TimerTick {
    foreach ($mode in @($script:Jobs.Keys)) {
        $job = $script:Jobs[$mode]
        $item = $null
        $count = 0
        while ($count -lt 40 -and $job.Queue.TryDequeue([ref]$item)) {
            $count++
            $job.Done = $job.Done + 1
            Receive-CheckResult -Mode $mode -Result $item -Job $job
        }
        if ($mode -eq 'Batch' -and $count -gt 0) { Update-BatchProgress -Job $job }
        if ((Test-JobFinished -Job $job) -and $job.Queue.IsEmpty) { Complete-CheckJob -Mode $mode }
    }

    $item = $null
    while ($script:Shared.ApiResults.TryDequeue([ref]$item)) {
        $changes = Set-ResultChanges -Result $item
        Add-HistoryEntry -Result $item -Origin 'API'
        Write-AccountLog -Result $item
        if ($changes) { Show-ChangeAlert -Result $item }
    }
    $item = $null
    while ($script:Shared.ApiLog.TryDequeue([ref]$item)) {
        $script:ApiLogRows.Insert(0, $item)
        while ($script:ApiLogRows.Count -gt $script:ApiLogLimit) { $script:ApiLogRows.RemoveAt($script:ApiLogRows.Count - 1) }
    }

    if ($null -ne $script:Api -and $script:Api.Handle.IsCompleted) {
        $reason = ''
        try { [void]$script:Api.Shell.EndInvoke($script:Api.Handle) } catch { $reason = Get-InnerMessage $_.Exception }
        Stop-ApiServer -Quiet
        Set-Status "The API server stopped unexpectedly. $reason" 'Error'
    }

    if ($null -ne $script:ApiTest -and $script:ApiTest.Handle.IsCompleted) {
        $text = ''
        try {
            $output = $script:ApiTest.Shell.EndInvoke($script:ApiTest.Handle)
            $text = (@($output) | ForEach-Object { [string]$_ }) -join "`r`n"
        } catch {
            $text = "Request failed: $(Get-InnerMessage $_.Exception)"
        }
        Close-WorkerShell -Shell $script:ApiTest.Shell
        $script:ApiTest = $null
        $script:UI.ApiTestResponse.Text = $text
        $script:UI.ApiTestButton.IsEnabled = $true
    }

    $cacheText = "Cache: $($script:Shared.Cache.Count)"
    if ($script:UI.StatusCache.Text -ne $cacheText) { $script:UI.StatusCache.Text = $cacheText }

    if ($script:DiagClock.ElapsedMilliseconds -ge 1000) {
        $script:DiagClock.Restart()
        Update-Diagnostics
    }
    if ($script:HistoryDirty -and $script:SaveClock.Elapsed.TotalSeconds -ge 30) {
        $script:SaveClock.Restart()
        Save-History
    }
    if ($script:RetentionClock.Elapsed.TotalMinutes -ge 60) {
        $script:RetentionClock.Restart()
        Invoke-Retention
    }
    if ($script:ClipboardClock.ElapsedMilliseconds -ge 700) {
        $script:ClipboardClock.Restart()
        Invoke-ClipboardWatch
    }
    if ($script:WatchClock.Elapsed.TotalSeconds -ge 5) {
        $script:WatchClock.Restart()
        Invoke-WatchFolder
    }
    if ($script:BatchGrouped -and $script:DomainDirty -and $script:DomainClock.ElapsedMilliseconds -ge 1000) {
        $script:DomainClock.Restart()
        Update-DomainView
    }
    if ($script:Tasks.Count -gt 0) { Invoke-BackgroundTasks }
}
#endregion

#region Single check
function Show-EmailError {
    param([string]$Message)
    $script:UI.EmailError.Text = $Message
    $script:UI.EmailError.Visibility = [System.Windows.Visibility]::Visible
    Set-Status $Message 'Error'
}

function Hide-EmailError {
    $script:UI.EmailError.Visibility = [System.Windows.Visibility]::Collapsed
}

function Set-SingleBusy {
    param([bool]$Busy)
    $ui = $script:UI
    $ui.CheckButton.IsEnabled = -not $Busy
    $ui.CheckButton.Content   = if ($Busy) { 'Checking...' } else { 'Check' }
    $ui.EmailBox.IsReadOnly   = $Busy
}

function Reset-ResultView {
    $ui = $script:UI
    $collapsed = [System.Windows.Visibility]::Collapsed
    $ui.ResultEmpty.Visibility   = [System.Windows.Visibility]::Visible
    $ui.ResultFields.Visibility  = $collapsed
    $ui.ResultBadge.Visibility   = $collapsed
    $ui.StatusBadge.Visibility   = $collapsed
    $ui.ResultCallout.Visibility = $collapsed
    $ui.ChangeCallout.Visibility = $collapsed
    $ui.ResultHint.Visibility    = $collapsed
    $ui.RawBox.Visibility        = $collapsed
    $ui.RecheckButton.Visibility = $collapsed
    $ui.RawBox.Text              = ''
    $ui.RawToggleButton.Content  = 'Show raw response'
    $ui.RawToggleButton.IsEnabled = $false
    $ui.CopyButton.IsEnabled     = $false
    $script:LastResult = $null
}

function Set-LinkValue {
    param($Link, $Run, [string]$Url)
    if ($Url) {
        $Run.Text = $Url
        $Link.NavigateUri = [Uri]$Url
    } else {
        $Run.Text = 'N/A'
        $Link.NavigateUri = $null
    }
}

function Show-SingleResult {
    param($Result)
    $ui = $script:UI
    $visible   = [System.Windows.Visibility]::Visible
    $collapsed = [System.Windows.Visibility]::Collapsed
    $script:LastResult = $Result
    $category = [string]$Result.Category
    $status   = [string]$Result.Status
    $advice   = Get-StatusAdvice -Status $status
    $statusLabel = if ($Result.StatusLabel) { [string]$Result.StatusLabel } else { $advice.Label }
    $nextAction  = if ($Result.NextAction) { [string]$Result.NextAction } else { $advice.Action }
    $statusKey   = Get-StatusColorKey -Status $status

    $ui.ResultEmpty.Visibility  = $collapsed
    $ui.ResultFields.Visibility = $visible

    $ui.ValUsername.Text = $Result.Username
    $ui.ValCloud.Text    = $Result.FoundInCloud
    Set-LinkValue -Link $ui.CloudUrlLink -Run $ui.CloudUrlRun -Url $Result.CloudURL
    $ui.CopyCloudUrlButton.Visibility = if ($Result.CloudURL) { $visible } else { $collapsed }

    $fedVisibility = if ($Result.FederationRedirectUrl) { $visible } else { $collapsed }
    $ui.FedLabel.Visibility = $fedVisibility
    $ui.FedValue.Visibility = $fedVisibility
    $ui.CopyFedButton.Visibility = $fedVisibility
    if ($Result.FederationRedirectUrl) { Set-LinkValue -Link $ui.FedUrlLink -Run $ui.FedUrlRun -Url $Result.FederationRedirectUrl }

    $ui.ValAccountType.Text = $Result.AccountType
    $ui.ValAccountType.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, "Cat$category")

    $anyDetail = $false
    foreach ($pair in @(@('Org', 'Organization'), @('Provider', 'FederationProvider'), @('TenantId', 'TenantId'), @('Region', 'TenantRegion'), @('Namespace', 'NamespaceType'))) {
        $value = [string]$Result.($pair[1])
        $visibility = if ($value) { $visible } else { $collapsed }
        $ui["Lbl$($pair[0])"].Visibility = $visibility
        $ui["Val$($pair[0])"].Visibility = $visibility
        $ui["Val$($pair[0])"].Text = $value
        if ($value) { $anyDetail = $true }
    }
    $ui.CopyOrgButton.Visibility    = if ($Result.Organization) { $visible } else { $collapsed }
    $ui.CopyTenantButton.Visibility = if ($Result.TenantId) { $visible } else { $collapsed }
    $ui.DetailDivider.Visibility    = if ($anyDetail) { $visible } else { $collapsed }

    $label = $script:CategoryLabels[$category]
    if (-not $label) { $label = $category }
    $ui.ResultBadgeText.Text = $label
    $ui.ResultBadgeText.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, "Cat$category")
    $ui.ResultBadge.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, "Cat$($category)Bg")
    $ui.ResultBadge.Visibility = $visible

    $ui.StatusBadgeText.Text = $statusLabel
    $ui.StatusBadgeText.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, $statusKey)
    $ui.StatusBadge.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, "$($statusKey)Bg")
    $ui.StatusBadge.Visibility = $visible

    # Anything that is not a definitive answer gets a prominent callout with the next step.
    if ($status -ne 'OK') {
        $calloutText = (@([string]$Result.Note, $nextAction) | Where-Object { $_ }) -join ' '
        $ui.ResultCalloutTitle.Text = $statusLabel
        $ui.ResultCalloutText.Text  = $calloutText
        $ui.ResultCalloutIcon.Text  = if ($status -eq 'Unavailable' -or $status -eq 'Invalid' -or $status -eq 'Error') { [string][char]0xE783 } else { [string][char]0xE7BA }
        $ui.ResultCallout.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, "$($statusKey)Bg")
        $ui.ResultCalloutIcon.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, $statusKey)
        $ui.ResultCalloutTitle.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, $statusKey)
        $ui.ResultCallout.Visibility = $visible
    } else {
        $ui.ResultCallout.Visibility = $collapsed
    }

    $sourceText = if ($Result.Source -eq 'Cache') { 'from the 24-hour cache' } else { 'live from Microsoft' }
    $ui.ResultMeta.Text = "Checked $($Result.Timestamp), $sourceText."
    $ui.RecheckButton.Visibility = if ($Result.Source -eq 'Cache') { $visible } else { $collapsed }

    $hint = ''
    if ($category -eq 'NotMicrosoft' -and $status -eq 'OK') {
        $hint = 'If you believe this result is incorrect, please confirm with your administrator that the email address is a valid Microsoft account and is set as the primary alias.'
    } elseif (($status -eq 'Unavailable' -or $status -eq 'Error') -and $Result.Error) {
        $hint = "Details: $($Result.Error)"
    }
    $ui.ResultHint.Text = $hint
    $ui.ResultHint.Visibility = if ($hint) { $visible } else { $collapsed }

    $ui.RawBox.Text = if ($Result.Raw) { $Result.Raw } else { 'No raw response is available for this result.' }
    $ui.RawToggleButton.IsEnabled = $true
    $ui.CopyButton.IsEnabled = $true

    switch ($status) {
        'OK'           { Set-Status "Checked $($Result.Username): $($Result.AccountType)" 'Success' }
        'Inconclusive' { Set-Status "Inconclusive answer for $($Result.Username). See the recommended next step." 'Warn' }
        'Throttled'    { Set-Status "Microsoft throttled the lookup for $($Result.Username). Check again in a few minutes." 'Warn' }
        default        { Set-Status "Check failed for $($Result.Username)." 'Error' }
    }
}

function Invoke-SingleCheck {
    param([switch]$SkipCache, [string]$Email = '')
    $ui = $script:UI
    if ($script:Jobs.ContainsKey('Single')) { return }
    if ($Email) { $ui.EmailBox.Text = $Email }
    $address = $ui.EmailBox.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($address)) { Show-EmailError 'Please enter an email address.'; return }
    if (-not (Test-EmailFormat $address)) { Show-EmailError 'Invalid email format. Please check and try again.'; return }
    Hide-EmailError
    Set-SingleBusy $true
    Set-Status $(if ($SkipCache) { 'Contacting Microsoft (cache skipped)...' } else { 'Contacting Microsoft...' }) 'Info'
    Start-CheckJob -Mode 'Single' -Emails @($address) -SkipCache ([bool]$SkipCache)
}
#endregion

#region Batch check
function Get-BatchTokens {
    param([string]$Text)
    $valid   = New-Object System.Collections.Generic.List[string]
    $invalid = New-Object System.Collections.Generic.List[string]
    $seen    = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $dupes   = 0
    $trimChars = [char[]]@('"', "'", '<', '>', '(', ')', '[', ']')
    foreach ($token in ([string]$Text -split '[\s,;]+')) {
        if (-not $token) { continue }
        $clean = $token.Trim().Trim($trimChars)
        if (-not $clean) { continue }
        if (-not $seen.Add($clean)) { $dupes++; continue }
        if (Test-EmailFormat $clean) { $valid.Add($clean) } else { $invalid.Add($clean) }
    }
    return @{ Valid = $valid; Invalid = $invalid; Duplicates = $dupes }
}

function Update-BatchCount {
    $tokens = Get-BatchTokens -Text $script:UI.BatchInput.Text
    $parts = @("$($tokens.Valid.Count) valid")
    if ($tokens.Invalid.Count -gt 0) { $parts += "$($tokens.Invalid.Count) invalid" }
    if ($tokens.Duplicates -gt 0) { $parts += "$($tokens.Duplicates) duplicates ignored" }
    $script:UI.BatchCount.Text = if (($tokens.Valid.Count + $tokens.Invalid.Count) -eq 0) { 'No addresses yet' } else { $parts -join ', ' }
}

function Initialize-BatchChips {
    $script:BatchChips = @{}
    $definitions = @(
        @{ Bucket = 'Entra';        Label = 'Work/School';      Key = 'CatEntra' }
        @{ Bucket = 'MSA';          Label = 'Personal';         Key = 'CatMSA' }
        @{ Bucket = 'Federated';    Label = 'Federated';        Key = 'CatFederated' }
        @{ Bucket = 'Dual';         Label = 'Dual identity';    Key = 'CatDual' }
        @{ Bucket = 'Sovereign';    Label = 'GCC High / China'; Key = 'CatGCCHigh' }
        @{ Bucket = 'NotMicrosoft'; Label = 'Not Microsoft';    Key = 'CatNotMicrosoft' }
        @{ Bucket = 'Error';        Label = 'Errors / unknown'; Key = 'CatError' }
        @{ Bucket = 'Inconclusive'; Label = 'Inconclusive';     Key = 'CatError' }
        @{ Bucket = 'Throttled';    Label = 'Throttled';        Key = 'CatError' }
        @{ Bucket = 'Unavailable';  Label = 'Unavailable';      Key = 'CatNotMicrosoft' }
        @{ Bucket = 'Changed';      Label = 'Changed';          Key = 'CatDual' }
    )
    foreach ($d in $definitions) {
        $border = New-Object System.Windows.Controls.Border
        $border.CornerRadius = [System.Windows.CornerRadius]::new(12)
        $border.Padding = [System.Windows.Thickness]::new(10, 3, 10, 3)
        $border.Margin  = [System.Windows.Thickness]::new(0, 0, 8, 6)
        $border.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, "$($d.Key)Bg")
        $text = New-Object System.Windows.Controls.TextBlock
        $text.FontSize = 12
        $text.FontWeight = [System.Windows.FontWeights]::SemiBold
        $text.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, $d.Key)
        $border.Child = $text
        [void]$script:UI.BatchChips.Children.Add($border)
        $script:BatchChips[$d.Bucket] = @{ Text = $text; Label = $d.Label }
    }
    Update-BatchChips
}

function Update-BatchChips {
    foreach ($bucket in @($script:BatchChips.Keys)) {
        $chip = $script:BatchChips[$bucket]
        $chip.Text.Text = '{0}  {1}' -f $chip.Label, $script:BatchCounts[$bucket]
    }
}

function Set-BatchBusy {
    param([bool]$Busy)
    $ui = $script:UI
    $ui.BatchStartButton.IsEnabled        = -not $Busy
    $ui.BatchPauseButton.IsEnabled        = $Busy
    $ui.BatchPauseButton.Content          = 'Pause'
    $ui.BatchStopButton.IsEnabled         = $Busy
    $ui.BatchImportButton.IsEnabled       = -not $Busy
    $ui.BatchPasteButton.IsEnabled        = -not $Busy
    $ui.BatchClearInputButton.IsEnabled   = -not $Busy
    $ui.BatchClearResultsButton.IsEnabled = -not $Busy
    $ui.BatchSkipCacheCheck.IsEnabled     = -not $Busy
    $ui.RefreshStaleCheck.IsEnabled       = -not $Busy
    $ui.BatchInput.IsReadOnly             = $Busy
    Update-RetryButton -Busy $Busy
}

function Update-BatchProgress {
    param($Job)
    if ($null -eq $Job) { return }
    $ui = $script:UI
    $ui.BatchProgress.Maximum = [Math]::Max(1, $Job.Total)
    $ui.BatchProgress.Value   = $Job.Done
    $text = '{0} of {1} checked' -f $Job.Done, $Job.Total
    if ($null -ne $Job.RetryIndex) { $text = "Retrying: $text" }
    if ($Job.Sync.Paused) { $text += ' - paused' }
    $ui.BatchProgressText.Text = $text
}

function Add-BatchRow {
    param($Result)
    $script:BatchIndex++
    $row = $Result.PSObject.Copy()
    $row | Add-Member -NotePropertyName 'Index' -NotePropertyValue $script:BatchIndex -Force
    $row | Add-Member -NotePropertyName 'Extra' -NotePropertyValue (Get-BatchExtraFor -Username $Result.Username) -Force
    if ($null -eq $row.PSObject.Properties['Changes']) { $row | Add-Member -NotePropertyName 'Changes' -NotePropertyValue '' -Force }
    $script:BatchRows.Add($row)
    Add-ToCounts -Counts $script:BatchCounts -Category ([string]$Result.Category) -Status ([string]$Result.Status)
    if ([string]$Result.Changes) { $script:BatchCounts['Changed'] = $script:BatchCounts['Changed'] + 1 }
    Update-BatchChips
    $script:DomainDirty = $true
}

function Clear-BatchResults {
    $script:BatchRows.Clear()
    $script:DomainRows.Clear()
    $script:BatchCounts = New-Counts
    $script:BatchIndex = 0
    $script:UI.BatchProgress.Value = 0
    $script:UI.BatchProgressText.Text = ''
    Update-BatchChips
    Update-RetryButton
}

function Start-Batch {
    param([string]$WatchFile = '')
    if ($script:Jobs.ContainsKey('Batch')) { return }
    $ui = $script:UI
    $tokens = Get-BatchTokens -Text $ui.BatchInput.Text
    if (($tokens.Valid.Count + $tokens.Invalid.Count) -eq 0) {
        Set-Status 'Add email addresses to the list first, or import a file.' 'Warn'
        return
    }
    Clear-BatchResults
    foreach ($bad in $tokens.Invalid) { Add-BatchRow -Result (Get-AccountCheck -Email $bad) }
    if ($tokens.Valid.Count -eq 0) {
        Update-RetryButton
        Set-Status 'None of the entries is a valid email address.' 'Error'
        return
    }
    $skip   = [bool]$ui.BatchSkipCacheCheck.IsChecked
    $maxAge = if ([bool]$ui.RefreshStaleCheck.IsChecked) { [double]$script:Settings.StaleAfterHours } else { 0 }
    Set-BatchBusy $true
    Start-CheckJob -Mode 'Batch' -Emails ([string[]]$tokens.Valid.ToArray()) -DelayMs $script:Settings.BatchDelayMs `
        -Workers $script:Settings.BatchConcurrency -SkipCache $skip -MaxAgeHours $maxAge
    $job = $script:Jobs['Batch']
    if ($WatchFile) { $job.WatchFile = $WatchFile; $job.Origin = 'Folder' }
    Update-BatchProgress -Job $job
    $notes = @()
    if ($tokens.Duplicates -gt 0) { $notes += "$($tokens.Duplicates) duplicates skipped" }
    if ($tokens.Invalid.Count -gt 0) { $notes += "$($tokens.Invalid.Count) invalid entries listed as errors" }
    if ($skip) { $notes += 'cache skipped' } elseif ($maxAge -gt 0) { $notes += "answers older than $maxAge hours refreshed" }
    $suffix = if ($notes.Count -gt 0) { ' (' + ($notes -join ', ') + ')' } else { '' }
    Set-Status "Checking $($tokens.Valid.Count) addresses$suffix..." 'Info'
}

function Stop-Batch {
    $job = $script:Jobs['Batch']
    if ($null -eq $job) { return }
    $job.Sync.Cancel = $true
    $job.Sync.Paused = $false
    $script:UI.BatchStopButton.IsEnabled = $false
    $script:UI.BatchPauseButton.IsEnabled = $false
    Set-Status 'Stopping after the current lookup...' 'Warn'
}

function Import-BatchFile {
    $dialog = New-Object Microsoft.Win32.OpenFileDialog
    $dialog.Filter = 'Email lists (*.csv;*.txt)|*.csv;*.txt|All files (*.*)|*.*'
    $dialog.Title  = 'Import email addresses'
    if ($dialog.ShowDialog($script:Window) -ne $true) { return }
    $path = $dialog.FileName
    $import = Import-EmailFile -Path $path
    if ([System.IO.Path]::GetExtension($path) -ieq '.csv' -and @($import.Columns).Count -gt 1) {
        Show-ImportDialog -Path $path -Import $import
        return
    }
    Set-BatchImport -Import $import -Name ([System.IO.Path]::GetFileName($path))
}
#endregion

#region History
function Add-HistoryEntry {
    param($Result, [string]$Origin)
    $entry = [pscustomobject][ordered]@{
        Timestamp             = $Result.Timestamp
        Username              = $Result.Username
        FoundInCloud          = $Result.FoundInCloud
        CloudURL              = $Result.CloudURL
        FederationRedirectUrl = $Result.FederationRedirectUrl
        AccountType           = $Result.AccountType
        Category              = $Result.Category
        Organization          = $Result.Organization
        TenantId              = $Result.TenantId
        TenantRegion          = $Result.TenantRegion
        NamespaceType         = $Result.NamespaceType
        FederationProvider    = $Result.FederationProvider
        Status                = $Result.Status
        StatusLabel           = $Result.StatusLabel
        NextAction            = $Result.NextAction
        Note                  = $Result.Note
        Changes               = [string]$Result.Changes
        ChangedSince          = [string]$Result.ChangedSince
        Source                = $Result.Source
        Origin                = $Origin
        Private               = [bool]$script:PrivateSession
    }
    $script:HistoryRows.Insert(0, $entry)
    [void]$script:Shared.History.Add($entry)
    Add-ToCounts -Counts $script:HistoryCounts -Category ([string]$Result.Category) -Status ([string]$Result.Status)
    if ($entry.Changes) { $script:HistoryCounts['Changed'] = $script:HistoryCounts['Changed'] + 1 }
    if ([string]$entry.Status -eq 'OK' -or [string]$entry.Status -eq 'Inconclusive') {
        $script:LatestByUser[([string]$entry.Username).ToLowerInvariant()] = $entry
    }
    if ($script:HistoryRows.Count -gt $script:HistoryLimit) {
        while ($script:HistoryRows.Count -gt $script:HistoryLimit) { $script:HistoryRows.RemoveAt($script:HistoryRows.Count - 1) }
        while ($script:Shared.History.Count -gt $script:HistoryLimit) { $script:Shared.History.RemoveAt(0) }
        $script:HistoryCounts = New-Counts
        foreach ($h in $script:HistoryRows) { Add-ToCounts -Counts $script:HistoryCounts -Category ([string]$h.Category) -Status ([string]$h.Status) }
    }
    $script:HistoryDirty = $true
    Update-HistoryStats
}

function Update-HistoryStats {
    foreach ($name in @('Total', 'Entra', 'MSA', 'Federated', 'Dual', 'Sovereign', 'NotMicrosoft', 'Error')) {
        $script:UI["Stat$name"].Text = [string]$script:HistoryCounts[$name]
    }
    Update-HistoryShown
}

function Clear-History {
    $answer = [System.Windows.MessageBox]::Show($script:Window, 'Delete all saved history? This cannot be undone.', $script:AppName,
        [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
    if ($answer -ne [System.Windows.MessageBoxResult]::Yes) { return }
    $script:HistoryRows.Clear()
    $script:Shared.History.Clear()
    $script:HistoryCounts = New-Counts
    $script:LatestByUser = @{}
    Update-HistoryStats
    Save-History
    Set-Status 'History cleared.' 'Success'
}

function Open-TodayLog {
    $file = Get-TodayLogFile
    if (Test-Path -LiteralPath $file) {
        try { Start-Process notepad.exe -ArgumentList ('"{0}"' -f $file) } catch { Show-AppError $_ 'Could not open the log file.' }
    } elseif (Test-Path -LiteralPath $script:LogsDir) {
        $msg = if ($script:Settings.EnableLogging) { 'No checks have been logged today yet. Opening the logs folder.' } else { 'Logging is off, so there is no log for today. Turn on the Logging switch to start one. Opening the logs folder.' }
        Set-Status $msg 'Warn'
        try { Start-Process explorer.exe -ArgumentList ('"{0}"' -f $script:LogsDir) } catch { }
    }
}

function Export-TodayLog {
    $file = Get-TodayLogFile
    if (-not (Test-Path -LiteralPath $file)) {
        Set-Status 'There is no log file for today to export.' 'Warn'
        return
    }
    $dialog = New-Object Microsoft.Win32.SaveFileDialog
    $dialog.Filter   = 'CSV files (*.csv)|*.csv|All files (*.*)|*.*'
    $dialog.FileName = 'MS_Account_Checker_Results_{0}.csv' -f (Get-Date -Format 'yyyy-MM-dd')
    if ($dialog.ShowDialog($script:Window) -ne $true) { return }
    Import-Csv -LiteralPath $file -Delimiter "`t" | Export-Csv -LiteralPath $dialog.FileName -NoTypeInformation -Encoding UTF8
    Set-Status "Log exported to $($dialog.FileName)" 'Success'
}
#endregion

#region Export
function Get-ExportRows {
    param($Items, [string[]]$ExtraColumns = @(), [switch]$Redact)
    foreach ($i in $Items) {
        $row = [ordered]@{
            Timestamp             = $i.Timestamp
            Username              = if ($Redact) { ConvertTo-RedactedEmail ([string]$i.Username) } else { $i.Username }
            FoundInCloud          = $i.FoundInCloud
            CloudURL              = $i.CloudURL
            FederationRedirectUrl = $i.FederationRedirectUrl
            AccountType           = $i.AccountType
            Organization          = $i.Organization
            TenantId              = $i.TenantId
            TenantRegion          = $i.TenantRegion
            DomainType            = $i.NamespaceType
            FederationProvider    = $i.FederationProvider
            Status                = $i.Status
            StatusLabel           = $i.StatusLabel
            NextAction            = $i.NextAction
            Changes               = $i.Changes
            Note                  = $i.Note
            Source                = $i.Source
            Origin                = $i.Origin
        }
        for ($c = 0; $c -lt @($ExtraColumns).Count; $c++) {
            $name  = $ExtraColumns[$c]
            $key   = if ($row.Contains($name)) { "Input $name" } else { $name }
            $value = ''
            if ($null -ne $i.Extra) { $value = [string]$i.Extra["c$c"] }
            $row[$key] = if ($Redact -and $value) { '[redacted]' } else { $value }
        }
        [pscustomobject]$row
    }
}

function New-HtmlReport {
    param($Items, [string]$Title, [switch]$Redact)
    $list = @($Items)
    $enc  = { param($value) [System.Net.WebUtility]::HtmlEncode([string]$value) }
    $who  = { param($value) if ($Redact) { ConvertTo-RedactedEmail ([string]$value) } else { [string]$value } }
    $pct  = { param($part, $whole) if ($whole -gt 0) { [int][Math]::Round(100.0 * $part / $whole) } else { 0 } }
    $counts = New-Counts
    foreach ($i in $list) {
        Add-ToCounts -Counts $counts -Category ([string]$i.Category) -Status ([string]$i.Status)
        if ([string]$i.Changes) { $counts['Changed'] = $counts['Changed'] + 1 }
    }
    $total      = $list.Count
    $definitive = @($list | Where-Object { [string]$_.Status -eq 'OK' }).Count
    $exceptions = @($list | Where-Object { [string]$_.Status -ne 'OK' -or [string]$_.Changes })
    $unique     = @($list | ForEach-Object { ([string]$_.Username).ToLowerInvariant() } | Sort-Object -Unique).Count
    $tenants    = @($list | ForEach-Object { [string]$_.TenantId } | Where-Object { $_ } | Sort-Object -Unique).Count
    $federated  = @($list | Where-Object { Test-IsFederated $_ }).Count
    $stamps     = @($list | ForEach-Object { [string]$_.Timestamp } | Where-Object { $_ } | Sort-Object)
    $period     = if ($stamps.Count) { "$($stamps[0]) to $($stamps[-1])" } else { 'no dates' }

    $css = @'
*{box-sizing:border-box}
body{font-family:"Segoe UI",Arial,sans-serif;background:#F4F7FC;color:#1F2937;margin:0;padding:28px;font-size:13px}
.wrap{max-width:1180px;margin:0 auto}
header{background:linear-gradient(90deg,#3659C9,#4668D8);color:#FFFFFF;border-radius:12px;padding:22px 28px;margin-bottom:18px}
header h1{font-size:22px;margin:0 0 4px 0}
header p{margin:0;color:#F1F5FF}
section{background:#FFFFFF;border:1px solid #D8E1EF;border-radius:12px;padding:18px 22px;margin-bottom:16px;break-inside:avoid}
h2{font-size:15px;margin:0 0 12px 0;color:#243E9E}
.kpis{display:grid;grid-template-columns:repeat(auto-fill,minmax(150px,1fr));gap:10px}
.kpi{border:1px solid #D8E1EF;border-radius:10px;padding:12px 14px}
.kpi b{display:block;font-size:24px}
.kpi span{color:#5D6C83;font-size:12px}
.cols{display:grid;grid-template-columns:1fr 1fr;gap:16px}
table{border-collapse:collapse;width:100%}
th{background:#F8FAFD;color:#5D6C83;text-align:left;padding:7px 9px;border-bottom:1px solid #D8E1EF;font-size:12px}
td{padding:6px 9px;border-bottom:1px solid #EEF2F8;vertical-align:top}
.num{text-align:right;white-space:nowrap}
.bar{height:9px;border-radius:5px;background:#3659C9;min-width:2px}
.barcell{width:40%}
.pill{display:inline-block;border-radius:12px;padding:1px 9px;font-weight:600}
.Entra{color:#15803D;background:#15803D1F}.MSA{color:#1D4ED8;background:#1D4ED81F}
.Federated{color:#475569;background:#4755691F}.Dual{color:#7E22CE;background:#7E22CE1F}
.GCCHigh{color:#0E7490;background:#0E74901F}.China{color:#A21CAF;background:#A21CAF1F}
.NotMicrosoft{color:#C62828;background:#C628281F}.Error{color:#B45309;background:#B453091F}.Unknown{color:#5D6C83;background:#5D6C831F}
.warn{color:#B45309;font-weight:600}
.change{color:#7E22CE;font-weight:600}
.trend{display:flex;align-items:stretch;gap:4px;height:120px;border-bottom:1px solid #D8E1EF;padding-top:6px}
.trend div{flex:1;display:flex;flex-direction:column;justify-content:flex-end;min-width:6px}
.trend .ok{background:#3659C9;border-radius:3px 3px 0 0}
.trend .bad{background:#B45309}
.trend-labels{display:flex;justify-content:space-between;color:#5D6C83;font-size:11px;margin-top:4px}
.muted{color:#5D6C83}
.no-print{margin-bottom:14px}
.no-print button{font:inherit;font-weight:600;border:0;border-radius:6px;padding:8px 14px;background:#3659C9;color:#FFFFFF;cursor:pointer}
@media (max-width:760px){.cols{grid-template-columns:1fr}}
@media print{body{background:#FFFFFF;padding:0}.no-print{display:none}section,header{break-inside:avoid}header{-webkit-print-color-adjust:exact;print-color-adjust:exact}.bar,.pill,.trend div{-webkit-print-color-adjust:exact;print-color-adjust:exact}.all{break-before:page}}
'@

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<!DOCTYPE html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>').Append((& $enc $Title)).Append('</title><style>').Append($css).Append('</style></head><body><div class="wrap">')
    [void]$sb.Append('<div class="no-print"><button type="button" onclick="window.print()">Print or save as PDF</button></div>')
    [void]$sb.Append('<header><h1>').Append((& $enc $Title)).Append('</h1><p>Generated ').Append((Get-Date -Format 'yyyy-MM-dd HH:mm')).Append(' by ').Append((& $enc "$($script:AppName) $($script:AppVersion)")).Append('. Checks from ').Append((& $enc $period)).Append('.')
    if ($Redact) { [void]$sb.Append(' Email addresses are redacted.') }
    [void]$sb.Append('</p></header>')

    # Key figures
    [void]$sb.Append('<section><h2>Summary</h2><div class="kpis">')
    foreach ($kpi in @(
        @($total, 'Checks'), @($unique, 'Unique addresses'), @(('{0}%' -f (& $pct $definitive $total)), 'Definitive answers'),
        @($exceptions.Count, 'Need attention or changed'), @($counts['Changed'], 'Changed since previous check'),
        @($tenants, 'Distinct tenants'), @(('{0}%' -f (& $pct $federated $total)), 'Federated'))) {
        [void]$sb.Append('<div class="kpi"><b>').Append((& $enc $kpi[0])).Append('</b><span>').Append((& $enc $kpi[1])).Append('</span></div>')
    }
    [void]$sb.Append('</div></section>')

    $breakdown = {
        param([string]$Heading, $Rows, [string]$FirstColumn)
        $max = 1
        foreach ($r in $Rows) { if ($r[1] -gt $max) { $max = $r[1] } }
        [void]$sb.Append('<section><h2>').Append((& $enc $Heading)).Append('</h2><table><thead><tr><th>').Append((& $enc $FirstColumn)).Append('</th><th class="num">Checks</th><th class="num">Share</th><th class="barcell"></th></tr></thead><tbody>')
        foreach ($r in $Rows) {
            [void]$sb.Append('<tr><td>').Append((& $enc $r[0])).Append('</td><td class="num">').Append($r[1]).Append('</td><td class="num">').Append((& $pct $r[1] $total)).Append('%</td><td class="barcell"><div class="bar" style="width:').Append([int](100 * $r[1] / $max)).Append('%"></div></td></tr>')
        }
        if (@($Rows).Count -eq 0) { [void]$sb.Append('<tr><td colspan="4" class="muted">None</td></tr>') }
        [void]$sb.Append('</tbody></table></section>')
    }
    $typeRows = @(foreach ($pair in @(@('Entra', 'Work/School'), @('MSA', 'Personal'), @('Federated', 'Federated'), @('Dual', 'Dual identity'),
                                       @('Sovereign', 'GCC High / China'), @('NotMicrosoft', 'Not Microsoft'), @('Error', 'Errors / unknown'))) {
        if ($counts[$pair[0]] -gt 0) { ,@($pair[1], $counts[$pair[0]]) } })
    $cloudRows = @($list | Group-Object { if ($_.FoundInCloud) { [string]$_.FoundInCloud } else { 'N/A' } } | Sort-Object Count -Descending | ForEach-Object { ,@($_.Name, $_.Count) })
    $providerRows = @($list | Where-Object { $_.FederationProvider } | Group-Object { [string]$_.FederationProvider } | Sort-Object Count -Descending | Select-Object -First 10 | ForEach-Object { ,@($_.Name, $_.Count) })
    $statusRows = @($list | Group-Object { if ($_.StatusLabel) { [string]$_.StatusLabel } else { [string]$_.Status } } | Sort-Object Count -Descending | ForEach-Object { ,@($_.Name, $_.Count) })
    [void]$sb.Append('<div class="cols">')
    & $breakdown 'Account types' $typeRows 'Account type'
    & $breakdown 'Certainty' $statusRows 'Status'
    & $breakdown 'Clouds' $cloudRows 'Cloud'
    & $breakdown 'Federation providers' $providerRows 'Provider'
    [void]$sb.Append('</div>')

    # Organizations and tenants
    $orgs = @($list | Where-Object { $_.Organization -or $_.TenantId } | Group-Object { "{0}|{1}" -f $_.Organization, $_.TenantId } | Sort-Object Count -Descending | Select-Object -First 15)
    [void]$sb.Append('<section><h2>Organizations and tenants</h2><table><thead><tr><th>Organization</th><th>Tenant ID</th><th>Region</th><th class="num">Checks</th></tr></thead><tbody>')
    foreach ($o in $orgs) {
        $first = $o.Group[0]
        [void]$sb.Append('<tr><td>').Append((& $enc $first.Organization)).Append('</td><td>').Append((& $enc $first.TenantId)).Append('</td><td>').Append((& $enc $first.TenantRegion)).Append('</td><td class="num">').Append($o.Count).Append('</td></tr>')
    }
    if ($orgs.Count -eq 0) { [void]$sb.Append('<tr><td colspan="4" class="muted">No tenant details in these results.</td></tr>') }
    [void]$sb.Append('</tbody></table></section>')

    # Trend by day
    $days = @($list | Where-Object { [string]$_.Timestamp -match '^\d{4}-\d{2}-\d{2}' } | Group-Object { ([string]$_.Timestamp).Substring(0, 10) } | Sort-Object Name | Select-Object -Last 30)
    if ($days.Count -gt 1) {
        $peak = ($days | Measure-Object Count -Maximum).Maximum
        [void]$sb.Append('<section><h2>Trend: checks per day</h2><div class="trend">')
        foreach ($d in $days) {
            $bad = @($d.Group | Where-Object { [string]$_.Status -ne 'OK' }).Count
            $okHeight  = [int](100 * ($d.Count - $bad) / $peak)
            $badHeight = [int](100 * $bad / $peak)
            [void]$sb.Append('<div title="').Append((& $enc ("{0}: {1} checks, {2} need attention" -f $d.Name, $d.Count, $bad))).Append('"><span class="bad" style="height:').Append($badHeight).Append('%"></span><span class="ok" style="height:').Append($okHeight).Append('%"></span></div>')
        }
        [void]$sb.Append('</div><div class="trend-labels"><span>').Append((& $enc $days[0].Name)).Append('</span><span>Blue: definitive. Amber: need attention.</span><span>').Append((& $enc $days[-1].Name)).Append('</span></div></section>')
    }

    # Exceptions
    [void]$sb.Append('<section><h2>Exceptions: need attention or changed (').Append($exceptions.Count).Append(')</h2><table><thead><tr><th>Time</th><th>Username</th><th>Account type</th><th>Status</th><th>What to do or what changed</th></tr></thead><tbody>')
    foreach ($i in ($exceptions | Select-Object -First 200)) {
        $detail = @([string]$i.NextAction, $(if ($i.Changes) { "Changed: $($i.Changes)" })) | Where-Object { $_ }
        $statusText = if ($i.StatusLabel) { $i.StatusLabel } else { $i.Status }
        $statusClass = if ([string]$i.Status -eq 'OK') { 'change' } else { 'warn' }
        [void]$sb.Append('<tr><td>').Append((& $enc $i.Timestamp)).Append('</td><td>').Append((& $enc (& $who $i.Username))).Append('</td><td>').Append((& $enc $i.AccountType)).Append('</td><td class="').Append($statusClass).Append('">').Append((& $enc $statusText)).Append('</td><td>').Append((& $enc ($detail -join ' '))).Append('</td></tr>')
    }
    if ($exceptions.Count -eq 0) { [void]$sb.Append('<tr><td colspan="5" class="muted">None. Every answer was definitive and nothing changed.</td></tr>') }
    if ($exceptions.Count -gt 200) { [void]$sb.Append('<tr><td colspan="5" class="muted">The first 200 are shown; the full list is below.</td></tr>') }
    [void]$sb.Append('</tbody></table></section>')

    # Domains
    $domains = @(Get-DomainSummary -Items $list | Select-Object -First 50)
    [void]$sb.Append('<section><h2>Domains</h2><table><thead><tr><th>Domain</th><th class="num">Addresses</th><th>Organization</th><th>Cloud</th><th>Federation provider</th><th>Account types</th><th class="num">Need attention</th></tr></thead><tbody>')
    foreach ($d in $domains) {
        [void]$sb.Append('<tr><td>').Append((& $enc $d.Domain)).Append('</td><td class="num">').Append($d.Addresses).Append('</td><td>').Append((& $enc $d.Organization)).Append('</td><td>').Append((& $enc $d.Cloud)).Append('</td><td>').Append((& $enc $d.FederationProvider)).Append('</td><td>').Append((& $enc $d.AccountTypes)).Append('</td><td class="num">').Append($d.NeedsAttention).Append('</td></tr>')
    }
    [void]$sb.Append('</tbody></table></section>')

    # Every result
    [void]$sb.Append('<section class="all"><h2>All results (').Append($total).Append(')</h2><table><thead><tr><th>Time</th><th>Username</th><th>FoundInCloud</th><th>Cloud URL</th><th>Account type</th><th>Organization</th><th>Federation provider</th><th>Tenant region</th><th>Status</th><th>Changes</th><th>Source</th></tr></thead><tbody>')
    foreach ($i in $list) {
        $cat = [string]$i.Category
        if ($cat -notmatch '^(Entra|MSA|Federated|Dual|GCCHigh|China|NotMicrosoft|Error|Unknown)$') { $cat = 'Unknown' }
        [void]$sb.Append('<tr><td>').Append((& $enc $i.Timestamp)).Append('</td><td>').Append((& $enc (& $who $i.Username))).Append('</td><td>').Append((& $enc $i.FoundInCloud))
        [void]$sb.Append('</td><td>').Append((& $enc $i.CloudURL)).Append('</td><td><span class="pill ').Append($cat).Append('">').Append((& $enc $i.AccountType))
        [void]$sb.Append('</span></td><td>').Append((& $enc $i.Organization)).Append('</td><td>').Append((& $enc $i.FederationProvider)).Append('</td><td>').Append((& $enc $i.TenantRegion))
        $statusAttr = if ([string]$i.Status -eq 'OK') { '' } else { ' class="warn"' }
        $statusText = if ($i.StatusLabel) { $i.StatusLabel } else { $i.Status }
        [void]$sb.Append('</td><td').Append($statusAttr).Append('>').Append((& $enc $statusText)).Append('</td><td class="change">').Append((& $enc $i.Changes)).Append('</td><td>').Append((& $enc $i.Source)).Append('</td></tr>')
    }
    [void]$sb.Append('</tbody></table></section></div></body></html>')
    return $sb.ToString()
}

function Export-Items {
    param($Items, [string]$Format, [string]$BaseName, [string[]]$ExtraColumns = @(), [string]$Title = 'Microsoft account check report', [switch]$Raw)
    $list = @($Items)
    if ($list.Count -eq 0) { Set-Status 'There is nothing to export yet.' 'Warn'; return }
    $redact = [bool]$script:Settings.RedactExports
    $dialog = New-Object Microsoft.Win32.SaveFileDialog
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    switch ($Format) {
        'CSV'  { $dialog.Filter = 'CSV files (*.csv)|*.csv';    $dialog.FileName = "$($BaseName)_$stamp.csv" }
        'JSON' { $dialog.Filter = 'JSON files (*.json)|*.json'; $dialog.FileName = "$($BaseName)_$stamp.json" }
        'HTML' { $dialog.Filter = 'HTML files (*.html)|*.html'; $dialog.FileName = "$($BaseName)_$stamp.html" }
        'PDF'  { $dialog.Filter = 'PDF files (*.pdf)|*.pdf';    $dialog.FileName = "$($BaseName)_$stamp.pdf" }
    }
    if ($dialog.ShowDialog($script:Window) -ne $true) { return }
    $path = $dialog.FileName
    $rows = if ($Raw) { $list } else { @(Get-ExportRows $list -ExtraColumns $ExtraColumns -Redact:$redact) }
    switch ($Format) {
        'CSV'  { $rows | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding UTF8 }
        'JSON' { ConvertTo-Json -InputObject @($rows) -Depth 3 | Set-Content -LiteralPath $path -Encoding UTF8 }
        'HTML' { New-HtmlReport -Items $list -Title $Title -Redact:$redact | Set-Content -LiteralPath $path -Encoding UTF8 }
        'PDF'  {
            Set-Status 'Creating the PDF...' 'Info'
            $temp = Join-Path ([System.IO.Path]::GetTempPath()) ('msac-report-' + [guid]::NewGuid().ToString('N') + '.html')
            New-HtmlReport -Items $list -Title $Title -Redact:$redact | Set-Content -LiteralPath $temp -Encoding UTF8
            try { Convert-HtmlToPdf -HtmlPath $temp -PdfPath $path } finally { Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue }
        }
    }
    Register-ExportFile -Path $path
    $note = if ($redact) { ' Addresses are redacted.' } else { '' }
    Set-Status "Exported $($list.Count) rows to $path.$note" 'Success'
    if ($Format -eq 'HTML' -or $Format -eq 'PDF') {
        $answer = [System.Windows.MessageBox]::Show($script:Window, 'Report saved. Open it now?', $script:AppName,
            [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
        if ($answer -eq [System.Windows.MessageBoxResult]::Yes) { Open-Url $path }
    }
}
#endregion

#region About
function Show-About {
    $ui = $script:UI
    $ui.AboutVersion.Text     = "Version $($script:AppVersion)"
    $ui.AboutAuthorRun.Text   = $script:AppAuthor
    $ui.AboutAuthorLink.NavigateUri = [Uri]("https://github.com/$($script:AppAuthor)")
    $ui.AboutRepoRun.Text     = $script:AppRepo -replace '^https://', ''
    $ui.AboutRepoLink.NavigateUri = [Uri]$script:AppRepo
    $ui.AboutLicense.Text     = $script:AppLicense
    $ui.AboutPsVersion.Text   = "$($PSVersionTable.PSVersion) ($($PSVersionTable.PSEdition))"
    $ui.AboutDataFolder.Text  = $script:AppDir
    $signature = 'Not signed. Compare the zip with the SHA-256 checksum on the release page.'
    try {
        if ($script:ScriptFile -and (Get-Command Get-AuthenticodeSignature -ErrorAction SilentlyContinue)) {
            $sig = Get-AuthenticodeSignature -FilePath $script:ScriptFile
            if ([string]$sig.Status -eq 'Valid') { $signature = "Signed by $($sig.SignerCertificate.Subject)" }
            elseif ([string]$sig.Status -ne 'NotSigned') { $signature = "Signature problem: $($sig.Status)" }
        }
    } catch { }
    $ui.AboutSignature.Text = $signature
    $ui.AboutOverlay.Visibility = [System.Windows.Visibility]::Visible
    [void]$ui.AboutCloseButton.Focus()
}

function Hide-About {
    $script:UI.AboutOverlay.Visibility = [System.Windows.Visibility]::Collapsed
}
#endregion

#region API server
function Get-ApiPort {
    $port = 0
    if (-not [int]::TryParse($script:UI.ApiPortBox.Text.Trim(), [ref]$port) -or $port -lt 1 -or $port -gt 65535) {
        throw 'Enter a port number between 1 and 65535 on the API server tab.'
    }
    return $port
}

function Get-EndpointHelp {
    param([int]$Port)
    $base  = "http://localhost:$Port"
    $token = [string]$script:Shared['ApiToken']
    $lines = @(
        'GET    /                    web page for the browser',
        'GET    /api/check?email=user@domain.com',
        'POST   /api/batch   body: ["a@x.com","b@y.com"]  (max 200)',
        'POST   /api/jobs    same body, up to 5,000, runs in the background',
        'GET    /api/jobs/{id}       progress and results; DELETE cancels',
        'GET    /api/history?limit=200&search=text',
        '       filters: type, cloud, status, federation, source, days',
        'GET    /api/health, /api/stats, /api/endpoints',
        'GET    /api/openapi.json    OpenAPI 3 description',
        'GET    /health              liveness, no token needed',
        'GET    /metrics             Prometheus metrics',
        ''
    )
    if ($token) {
        $lines += 'A token is required (Settings tab). PowerShell:'
        $lines += "Invoke-RestMethod `"$base/api/check?email=user@domain.com`" -Headers @{ 'X-API-Key' = '$token' }"
        $lines += ''
        $lines += 'curl:'
        $lines += "curl -H `"X-API-Key: $token`" `"$base/api/check?email=user@domain.com`""
    } else {
        $lines += 'PowerShell:'
        $lines += "Invoke-RestMethod `"$base/api/check?email=user@domain.com`""
        $lines += ''
        $lines += 'curl:'
        $lines += "curl `"$base/api/check?email=user@domain.com`""
    }
    return $lines -join "`r`n"
}

function Update-ApiUi {
    $ui = $script:UI
    $running = $null -ne $script:Api
    if ($running) {
        $port = $script:Api.Port
    } else {
        $port = 0
        if (-not [int]::TryParse($ui.ApiPortBox.Text.Trim(), [ref]$port) -or $port -lt 1 -or $port -gt 65535) { $port = $script:Settings.ApiPort }
    }
    $base = "http://localhost:$port/"
    $ui.ApiStatusText.Text = if ($running) { "Running on port $port" } else { 'Stopped' }
    $dotKey = if ($running) { 'CatEntra' } else { 'Muted' }
    $ui.ApiDot.SetResourceReference([System.Windows.Shapes.Shape]::FillProperty, $dotKey)
    $ui.ApiPortBox.IsEnabled = -not $running
    $ui.ApiUrlRun.Text = $base
    $ui.ApiUrlLink.NavigateUri = [Uri]($base + 'api/endpoints')
    $ui.ApiEndpointsBox.Text = Get-EndpointHelp -Port $port
    $ui.StatusApi.Text = if ($running) { "API: $base" } else { 'API: off' }
    if ([bool]$ui.ApiToggle.IsChecked -ne $running) { Set-ToggleSilently -Toggle $ui.ApiToggle -Value $running }
    $ui.ApiOpenWebButton.IsEnabled = $running
    Update-Tray
}

function Start-ApiServer {
    if ($null -ne $script:Api) { return }
    $port = Get-ApiPort
    $listener = [System.Net.HttpListener]::new()
    $listener.Prefixes.Add("http://localhost:$port/")
    try {
        $listener.Start()
    } catch {
        $reason = Get-InnerMessage $_.Exception
        try { $listener.Close() } catch { }
        throw "Port $port could not be opened: $reason Another program may be using it; pick a different port on the API server tab."
    }
    Sync-SharedSettings
    $shell = New-WorkerShell
    [void]$shell.AddScript($script:ApiScript.ToString())
    [void]$shell.AddParameter('CoreText', $script:CoreText)
    [void]$shell.AddParameter('Listener', $listener)
    [void]$shell.AddParameter('Shared', $script:Shared)
    [void]$shell.AddParameter('StartedUtc', [datetime]::UtcNow)
    [void]$shell.AddParameter('TtlHours', [double]$script:CacheHours)
    [void]$shell.AddParameter('Version', $script:AppVersion)
    [void]$shell.AddParameter('WebPage', $script:WebPage)
    [void]$shell.AddParameter('Port', $port)
    $handle = $shell.BeginInvoke()
    $control = [hashtable]::Synchronized(@{ Stop = $false })
    $runner = New-WorkerShell
    [void]$runner.AddScript($script:JobRunnerScript.ToString())
    [void]$runner.AddParameter('CoreText', $script:CoreText)
    [void]$runner.AddParameter('Shared', $script:Shared)
    [void]$runner.AddParameter('TtlHours', [double]$script:CacheHours)
    [void]$runner.AddParameter('Control', $control)
    $runnerHandle = $runner.BeginInvoke()
    $script:Api = @{ Listener = $listener; Shell = $shell; Handle = $handle; Port = $port; Started = Get-Date; Runner = $runner; RunnerHandle = $runnerHandle; Control = $control }
    $script:Settings.ApiPort = $port
    Save-Settings
    Update-ApiUi
    Set-Status "API server is listening on http://localhost:$port/" 'Success'
    if ($script:Settings.OpenBrowserOnApiStart -and -not $script:NoBrowser) { Open-Url "http://localhost:$port/" }
}

function Stop-ApiServer {
    param([switch]$Quiet)
    $api = $script:Api
    if ($null -eq $api) { Update-ApiUi; return }
    $script:Api = $null
    try { $api.Control.Stop = $true } catch { }
    try { $api.Listener.Stop() } catch { }
    try { $api.Listener.Close() } catch { }
    Close-WorkerShell -Shell $api.Shell -WaitMs 2000 -Handle $api.Handle
    Close-WorkerShell -Shell $api.Runner -WaitMs 1000 -Handle $api.RunnerHandle
    Update-ApiUi
    if (-not $Quiet) { Set-Status 'API server stopped.' 'Info' }
    if ($script:Tray -and -not $script:Window.IsVisible -and -not $script:ExitRequested) {
        $script:Tray.ShowBalloonTip(3000, $script:AppName, 'The API server stopped. Right-click the tray icon to start it again or open the window.', [System.Windows.Forms.ToolTipIcon]::Info)
    }
}

function Invoke-ApiTest {
    $ui = $script:UI
    if ($null -eq $script:Api) { Set-Status 'Turn on the API server switch first.' 'Warn'; return }
    if ($null -ne $script:ApiTest) { return }
    $email = $ui.ApiTestEmail.Text.Trim()
    if (-not $email) { $email = $ui.EmailBox.Text.Trim() }
    if (-not $email) { Set-Status 'Enter an email address to send to the API.' 'Warn'; return }
    $url = "http://localhost:$($script:Api.Port)/api/check?email=" + [Uri]::EscapeDataString($email)
    $shell = New-WorkerShell
    [void]$shell.AddScript($script:ApiTestScript.ToString())
    [void]$shell.AddParameter('Url', $url)
    [void]$shell.AddParameter('Token', [string]$script:Shared['ApiToken'])
    $handle = $shell.BeginInvoke()
    $script:ApiTest = @{ Shell = $shell; Handle = $handle }
    $ui.ApiTestButton.IsEnabled = $false
    $ui.ApiTestResponse.Text = "GET $url`r`n`r`nWaiting for the response..."
}
#endregion

#region 2.0 features: filters, retention, diagnostics, API token, shortcuts, menus
function Test-HighContrast {
    try { return [bool][System.Windows.SystemParameters]::HighContrast } catch { return $false }
}

function ConvertTo-HexColor {
    param($Color)
    return '#{0:X2}{1:X2}{2:X2}' -f $Color.R, $Color.G, $Color.B
}

function Get-HighContrastPalette {
    $sc = [System.Windows.SystemColors]
    $window        = ConvertTo-HexColor $sc::WindowColor
    $text          = ConvertTo-HexColor $sc::WindowTextColor
    $highlight     = ConvertTo-HexColor $sc::HighlightColor
    $highlightText = ConvertTo-HexColor $sc::HighlightTextColor
    $gray          = ConvertTo-HexColor $sc::GrayTextColor
    $control       = ConvertTo-HexColor $sc::ControlColor
    $controlText   = ConvertTo-HexColor $sc::ControlTextColor
    $link          = ConvertTo-HexColor $sc::HotTrackColor
    $palette = [ordered]@{
        Bg = $window; Surface = $window; Text = $text; Muted = $gray; Border = $text
        Primary = $highlight; PrimaryText = $highlightText; Secondary = $control; SecondaryText = $controlText
        Danger = $highlight; DangerText = $highlightText; InputBg = $window; GridBg = $window; GridAlt = $window
        Selection = $highlight; Link = $link; HeaderText = $text; HeaderMuted = $text; SwitchOn = $highlight
        HeaderStart = $window; HeaderEnd = $window
        HeaderButtonBg = $control; HeaderButtonBorder = $text; SwitchTrack = $window; SwitchTrackBorder = $text; SwitchThumb = $text
    }
    foreach ($key in @('CatEntra', 'CatMSA', 'CatFederated', 'CatDual', 'CatGCCHigh', 'CatChina', 'CatNotMicrosoft', 'CatError', 'CatUnknown')) {
        $palette[$key] = $text
    }
    return $palette
}

function Test-JobFinished {
    param($Job)
    foreach ($handle in $Job.Handles) { if (-not $handle.IsCompleted) { return $false } }
    return $true
}

function Stop-JobShells {
    param($Job, [int]$WaitMs = 0)
    for ($i = 0; $i -lt $Job.Shells.Count; $i++) { Close-WorkerShell -Shell $Job.Shells[$i] -WaitMs $WaitMs -Handle $Job.Handles[$i] }
}

function Get-StatusColorKey {
    param([string]$Status)
    switch ($Status) {
        'OK'           { return 'CatEntra' }
        'Inconclusive' { return 'CatError' }
        'Throttled'    { return 'CatError' }
        'Unavailable'  { return 'CatNotMicrosoft' }
        'Error'        { return 'CatNotMicrosoft' }
        'Invalid'      { return 'CatNotMicrosoft' }
        default        { return 'CatUnknown' }
    }
}

function Update-RetryButton {
    param([bool]$Busy = $false)
    $failed = 0
    foreach ($row in $script:BatchRows) {
        $status = [string]$row.Status
        if ($status -eq 'Throttled' -or $status -eq 'Unavailable' -or $status -eq 'Error') { $failed++ }
    }
    $script:UI.BatchRetryButton.IsEnabled = (-not $Busy) -and ($failed -gt 0)
    $script:UI.BatchRetryButton.Content   = if ($failed -gt 0) { "Retry failed ($failed)" } else { 'Retry failed' }
}

function Update-BatchRow {
    param($Result, [int]$Index)
    $old = $script:BatchRows[$Index]
    $row = $Result.PSObject.Copy()
    $row | Add-Member -NotePropertyName 'Index' -NotePropertyValue $old.Index -Force
    $row | Add-Member -NotePropertyName 'Extra' -NotePropertyValue $old.Extra -Force
    if ($null -eq $row.PSObject.Properties['Changes']) { $row | Add-Member -NotePropertyName 'Changes' -NotePropertyValue '' -Force }
    $script:BatchRows[$Index] = $row
    Add-ToCounts -Counts $script:BatchCounts -Category ([string]$old.Category) -Status ([string]$old.Status) -Delta -1
    Add-ToCounts -Counts $script:BatchCounts -Category ([string]$Result.Category) -Status ([string]$Result.Status)
    if ([string]$old.Changes) { $script:BatchCounts['Changed'] = $script:BatchCounts['Changed'] - 1 }
    if ([string]$Result.Changes) { $script:BatchCounts['Changed'] = $script:BatchCounts['Changed'] + 1 }
    Update-BatchChips
    $script:DomainDirty = $true
}

function Update-BatchSettingsText {
    $s = $script:Settings
    $parallel = if ($s.BatchConcurrency -eq 1) { '1 lookup at a time' } else { "$($s.BatchConcurrency) lookups at a time" }
    $script:UI.BatchSettingsText.Text = "$parallel, $($s.BatchDelayMs) ms apart. Change this in Settings."
}

function Start-BatchRetry {
    if ($script:Jobs.ContainsKey('Batch')) { return }
    $index  = @{}
    $emails = New-Object System.Collections.Generic.List[string]
    for ($i = 0; $i -lt $script:BatchRows.Count; $i++) {
        $row = $script:BatchRows[$i]
        $status = [string]$row.Status
        if ($status -eq 'Throttled' -or $status -eq 'Unavailable' -or $status -eq 'Error') {
            $key = ([string]$row.Username).ToLowerInvariant()
            if (-not $index.ContainsKey($key)) { $index[$key] = $i; $emails.Add([string]$row.Username) }
        }
    }
    if ($emails.Count -eq 0) { Set-Status 'There are no failed results to retry.' 'Info'; return }
    Set-BatchBusy $true
    Start-CheckJob -Mode 'Batch' -Emails ([string[]]$emails.ToArray()) -DelayMs $script:Settings.BatchDelayMs `
        -Workers $script:Settings.BatchConcurrency -SkipCache $true -RetryIndex $index
    Update-BatchProgress -Job $script:Jobs['Batch']
    Set-Status "Retrying $($emails.Count) failed lookups..." 'Info'
}

function Switch-BatchPause {
    $job = $script:Jobs['Batch']
    if ($null -eq $job) { return }
    $job.Sync.Paused = -not $job.Sync.Paused
    $script:UI.BatchPauseButton.Content = if ($job.Sync.Paused) { 'Resume' } else { 'Pause' }
    Update-BatchProgress -Job $job
    if ($job.Sync.Paused) { Set-Status 'Batch paused after the current lookup. Select Resume to continue.' 'Info' } else { Set-Status 'Batch resumed.' 'Info' }
}

function Get-ComboTag {
    param($Combo)
    $item = $Combo.SelectedItem
    if ($null -eq $item) { return 'All' }
    return [string]$item.Tag
}

function Update-HistoryFilter {
    if ($script:SuppressFilter) { return }
    $ui = $script:UI
    $state = @{
        Text       = $ui.HistorySearch.Text.Trim()
        Type       = Get-ComboTag $ui.FilterType
        Cloud      = Get-ComboTag $ui.FilterCloud
        Status     = Get-ComboTag $ui.FilterStatus
        Federation = Get-ComboTag $ui.FilterFederation
        Source     = Get-ComboTag $ui.FilterSource
        Since      = $null
    }
    $dateTag = Get-ComboTag $ui.FilterDate
    if ($dateTag -ne 'All') { $state.Since = (Get-Date).Date.AddDays(-[int]$dateTag) }
    $script:HistoryFilterState = $state
    $script:HistoryView.Refresh()
    Update-HistoryShown
}

function Test-HistoryItem {
    param($Item)
    $f = $script:HistoryFilterState
    if ($null -eq $f) { return $true }
    if ($f.Type -ne 'All' -and (Get-Bucket ([string]$Item.Category)) -ne $f.Type) { return $false }
    if ($f.Cloud -ne 'All' -and [string]$Item.FoundInCloud -ne $f.Cloud) { return $false }
    if ($f.Status -eq 'Changed') {
        if (-not [string]$Item.Changes) { return $false }
    } elseif ($f.Status -ne 'All') {
        $status = [string]$Item.Status
        if ($status -eq 'Error') { $status = 'Unavailable' }
        if ($status -ne $f.Status) { return $false }
    }
    if ($f.Federation -ne 'All') {
        $federated = ([string]$Item.Category -eq 'Federated') -or [bool]$Item.FederationProvider
        if (($f.Federation -eq 'Yes') -ne $federated) { return $false }
    }
    if ($f.Source -ne 'All' -and [string]$Item.Source -ne $f.Source) { return $false }
    if ($null -ne $f.Since) {
        $time = [datetime]::MinValue
        $parsed = [datetime]::TryParseExact([string]$Item.Timestamp, 'yyyy-MM-dd HH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$time)
        if (-not $parsed -or $time -lt $f.Since) { return $false }
    }
    if ($f.Text) {
        foreach ($value in @($Item.Username, $Item.AccountType, $Item.FoundInCloud, $Item.Organization, $Item.Origin, $Item.StatusLabel, $Item.FederationProvider)) {
            if (([string]$value).IndexOf($f.Text, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { return $true }
        }
        return $false
    }
    return $true
}

function Reset-HistoryFilters {
    $ui = $script:UI
    $script:SuppressFilter = $true
    try {
        foreach ($combo in @($ui.FilterType, $ui.FilterCloud, $ui.FilterStatus, $ui.FilterFederation, $ui.FilterSource, $ui.FilterDate)) { $combo.SelectedIndex = 0 }
        $ui.HistorySearch.Text = ''
    } finally {
        $script:SuppressFilter = $false
    }
    Update-HistoryFilter
}

function Update-HistoryShown {
    if ($null -eq $script:HistoryView -or -not $script:UI.HistoryShownText) { return }
    $total = $script:HistoryRows.Count
    $shown = $script:HistoryView.Count
    $script:UI.HistoryShownText.Text = if ($shown -eq $total) { "$total checks" } else { "Showing $shown of $total checks" }
}

function Get-ShownHistory {
    $list = New-Object System.Collections.Generic.List[object]
    foreach ($item in $script:HistoryView) { $list.Add($item) }
    return $list.ToArray()
}

function Reset-HistoryMirror {
    $script:Shared.History.Clear()
    for ($i = $script:HistoryRows.Count - 1; $i -ge 0; $i--) { [void]$script:Shared.History.Add($script:HistoryRows[$i]) }
    $script:HistoryCounts = New-Counts
    $script:LatestByUser = @{}
    foreach ($h in $script:HistoryRows) {
        Add-ToCounts -Counts $script:HistoryCounts -Category ([string]$h.Category) -Status ([string]$h.Status)
        $key = ([string]$h.Username).ToLowerInvariant()
        if (([string]$h.Status -eq 'OK' -or [string]$h.Status -eq 'Inconclusive') -and -not $script:LatestByUser.ContainsKey($key)) { $script:LatestByUser[$key] = $h }
    }
    Update-HistoryStats
}

function Invoke-Retention {
    $removed = 0
    $days = [int]$script:Settings.HistoryRetentionDays
    if ($days -gt 0) {
        $cutoff = (Get-Date).AddDays(-$days)
        for ($i = $script:HistoryRows.Count - 1; $i -ge 0; $i--) {
            $time = [datetime]::MinValue
            $ok = [datetime]::TryParseExact([string]$script:HistoryRows[$i].Timestamp, 'yyyy-MM-dd HH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$time)
            if ($ok -and $time -lt $cutoff) { $script:HistoryRows.RemoveAt($i); $removed++ }
        }
        if ($removed -gt 0) { Reset-HistoryMirror; $script:HistoryDirty = $true }
    }
    $deleted = 0
    $logDays = [int]$script:Settings.LogRetentionDays
    if ($logDays -gt 0 -and (Test-Path -LiteralPath $script:LogsDir)) {
        $logCutoff = (Get-Date).Date.AddDays(-$logDays)
        foreach ($file in @(Get-ChildItem -LiteralPath $script:LogsDir -Filter 'MS_Account_Checker_Logs_*.txt' -File -ErrorAction SilentlyContinue)) {
            if ($file.LastWriteTime -lt $logCutoff) {
                try { Remove-Item -LiteralPath $file.FullName -Force -ErrorAction Stop; $deleted++ } catch { }
            }
        }
    }
    if ($removed -gt 0 -or $deleted -gt 0) {
        Set-Status ("Retention: removed {0} old history entries and {1} old log files." -f $removed, $deleted) 'Info'
    }
}

function Remove-AllLogs {
    $files = @(Get-ChildItem -LiteralPath $script:LogsDir -Filter 'MS_Account_Checker_Logs_*' -File -ErrorAction SilentlyContinue)
    if ($files.Count -eq 0) { Set-Status 'There are no log files to delete.' 'Info'; return }
    $answer = [System.Windows.MessageBox]::Show($script:Window, "Delete all $($files.Count) log files? This cannot be undone.", $script:AppName,
        [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
    if ($answer -ne [System.Windows.MessageBoxResult]::Yes) { return }
    $deleted = 0
    foreach ($file in $files) { try { Remove-Item -LiteralPath $file.FullName -Force -ErrorAction Stop; $deleted++ } catch { } }
    $script:LogChecked = @{}
    Set-Status "Deleted $deleted log files." 'Success'
}

function New-ApiToken {
    $bytes = New-Object byte[] 24
    $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    try { $rng.GetBytes($bytes) } finally { $rng.Dispose() }
    return (($bytes | ForEach-Object { $_.ToString('x2') }) -join '')
}

function Set-ApiTokenState {
    param([bool]$Required, [switch]$NewToken)
    $s = $script:Settings
    if ($NewToken -or ($Required -and -not $s.ApiToken)) { $s.ApiToken = New-ApiToken }
    $s.ApiRequireToken = $Required
    $script:Shared['ApiToken'] = if ($Required) { [string]$s.ApiToken } else { '' }
    $ui = $script:UI
    $ui.ApiTokenBox.Text = if ($Required) { [string]$s.ApiToken } else { 'Not required' }
    $ui.ApiTokenCopyButton.IsEnabled = $Required
    $ui.ApiTokenNewButton.IsEnabled  = $Required
    if ([bool]$ui.ApiTokenCheck.IsChecked -ne $Required) { Set-ToggleSilently -Toggle $ui.ApiTokenCheck -Value $Required }
    Save-Settings
    Update-ApiUi
}

function Update-Diagnostics {
    $ui = $script:UI
    if ($ui.MainTabs.SelectedItem -ne $ui.TabDiagnostics) { return }
    $snap = Get-CoreStatsSnapshot -Stats $script:Shared.Stats
    if ($null -eq $snap) { return }
    $ui.DiagLive.Text         = [string]$snap.LiveLookups
    $ui.DiagCacheHits.Text    = [string]$snap.CacheHits
    $ui.DiagHitRate.Text      = if ($snap.CacheHitRate) { [string]$snap.CacheHitRate } else { '-' }
    $ui.DiagRetries.Text      = [string]$snap.Retries
    $ui.DiagThrottle.Text     = [string]$snap.ThrottleEvents
    $ui.DiagUnavailable.Text  = [string]$snap.Unavailable
    $ui.DiagInconclusive.Text = [string]$snap.Inconclusive
    $ui.DiagApiRequests.Text  = [string]$script:Shared['RequestCount']
    $ui.DiagLastSuccess.Text  = if ($snap.LastSuccess) { [string]$snap.LastSuccess } else { 'None yet' }
    $ui.DiagLastError.Text    = if ($snap.LastError) { [string]$snap.LastError } else { 'None' }
    if ($null -ne $script:Api) {
        $last = if ($script:Shared['LastRequest']) { "last request $($script:Shared['LastRequest'])" } else { 'no requests yet' }
        $jobs = @($script:Shared.ApiJobs.Values | Where-Object { $_.Status -eq 'queued' -or $_.Status -eq 'running' }).Count
        $ui.DiagApiState.Text = "Running on port $($script:Api.Port) since $($script:Api.Started.ToString('yyyy-MM-dd HH:mm:ss')), $last. Background jobs waiting or running: $jobs."
    } else {
        $ui.DiagApiState.Text = 'Off'
    }
    $ui.DiagApiRejected.Text = [string]$script:Shared['RejectedCount']
    $ui.DiagSince.Text = [string]$snap.Since
    $script:DiagRows.Clear()
    foreach ($row in $snap.Endpoints) { $script:DiagRows.Add($row) }

    $recent = $snap.Recent
    $script:LastAdvice = $recent
    $ui.DiagRecentAvailability.Text = if ($null -ne $recent.AvailabilityPct) { '{0:0.0}% of {1} calls' -f $recent.AvailabilityPct, $recent.Calls } else { 'No calls in the last 5 minutes' }
    $ui.DiagRecentP50.Text      = if ($recent.LookupCalls) { "$($recent.MedianMs) ms" } else { '-' }
    $ui.DiagRecentP95.Text      = if ($recent.LookupCalls) { "$($recent.P95Ms) ms" } else { '-' }
    $ui.DiagRecentThrottle.Text = if ($recent.LookupCalls) { '{0:0.0}% ({1} events)' -f $recent.ThrottleRatePct, $recent.Throttles } else { '-' }
    $parallel = if ($recent.SuggestedParallel -eq 1) { '1 lookup at a time' } else { "$($recent.SuggestedParallel) lookups at a time" }
    $ui.DiagSuggestion.Text = "$parallel, $($recent.SuggestedDelayMs) ms apart. $($recent.Suggestion)"
    $same = ([int]$script:Settings.BatchConcurrency -eq [int]$recent.SuggestedParallel) -and ([int]$script:Settings.BatchDelayMs -eq [int]$recent.SuggestedDelayMs)
    $ui.DiagApplyButton.IsEnabled = -not $same
    $ui.DiagApplyButton.Content = if ($same) { 'Batch settings already match' } else { 'Apply to batch settings' }
    Update-DiagChart -Timeline $snap.Timeline
}

function Set-NumberSetting {
    param($Box)
    $definition = $script:NumberFields[[string]$Box.Name]
    if ($null -eq $definition) { return }
    $name = $definition.Name
    $value = 0
    if (-not [int]::TryParse($Box.Text.Trim(), [ref]$value)) {
        $Box.Text = [string]$script:Settings[$name]
        Set-Status "Enter a whole number from $($definition.Min) to $($definition.Max)." 'Warn'
        return
    }
    $value = [Math]::Max($definition.Min, [Math]::Min($definition.Max, $value))
    $Box.Text = [string]$value
    if ([int]$script:Settings[$name] -ne $value) {
        $script:Settings[$name] = $value
        Save-Settings
        Sync-SharedSettings
        if ($name -like '*Retention*') { Invoke-Retention }
        Update-BatchSettingsText
        Set-Status 'Setting saved.' 'Success'
    }
}

function Copy-Field {
    param([string]$Value, [string]$Name)
    if (-not $Value) { return }
    if (Set-ClipboardText $Value) { Set-Status "$Name copied to the clipboard." 'Success' }
    else { Set-Status 'The clipboard is busy. Try again.' 'Warn' }
}

function New-GridContextMenu {
    param($Grid)
    $menu = New-Object System.Windows.Controls.ContextMenu
    foreach ($definition in @(@('Copy username', 'Username'), @('Copy tenant ID', 'TenantId'), @('Copy rows', 'Row'), @('-', ''), @('Check again live', 'Recheck'))) {
        if ($definition[0] -eq '-') { [void]$menu.Items.Add((New-Object System.Windows.Controls.Separator)); continue }
        $item = New-Object System.Windows.Controls.MenuItem
        $item.Header = $definition[0]
        $item.Tag = $definition[1]
        $item.Add_Click({ param($sender, $e) try { Invoke-GridMenuAction -Item $sender } catch { Show-AppError $_ } })
        [void]$menu.Items.Add($item)
    }
    $menu.Add_Opened({
        param($sender, $e)
        try {
            $sender.Background  = $script:Window.Resources['Surface']
            $sender.Foreground  = $script:Window.Resources['Text']
            $sender.BorderBrush = $script:Window.Resources['Border']
        } catch { }
    })
    $Grid.ContextMenu = $menu
}

function Invoke-GridMenuAction {
    param($Item)
    $grid = $Item.Parent.PlacementTarget
    if ($null -eq $grid) { return }
    $rows = @($grid.SelectedItems)
    if ($rows.Count -eq 0 -and $null -ne $grid.CurrentItem) { $rows = @($grid.CurrentItem) }
    if ($rows.Count -eq 0) { return }
    switch ([string]$Item.Tag) {
        'Username' { Copy-Field -Value (($rows | ForEach-Object { [string]$_.Username }) -join "`r`n") -Name 'Username' }
        'TenantId' {
            $ids = @($rows | ForEach-Object { [string]$_.TenantId } | Where-Object { $_ })
            if ($ids.Count -eq 0) { Set-Status 'The selected rows have no tenant ID.' 'Info' } else { Copy-Field -Value ($ids -join "`r`n") -Name 'Tenant ID' }
        }
        'Row' {
            $columns = @('Timestamp', 'Username', 'FoundInCloud', 'AccountType', 'StatusLabel', 'Organization', 'TenantId', 'TenantRegion', 'FederationProvider', 'Source')
            $lines = @($columns -join "`t")
            foreach ($row in $rows) { $lines += (($columns | ForEach-Object { ([string]$row.$_) -replace '[\t\r\n]+', ' ' }) -join "`t") }
            Copy-Field -Value ($lines -join "`r`n") -Name 'Rows'
        }
        'Recheck' {
            $script:UI.MainTabs.SelectedItem = $script:UI.TabSingle
            Invoke-SingleCheck -SkipCache -Email ([string]$rows[0].Username)
        }
    }
}

function Invoke-Shortcut {
    param($KeyEvent)
    $ui = $script:UI
    $key = $KeyEvent.Key
    if ($key -eq [System.Windows.Input.Key]::System) { $key = $KeyEvent.SystemKey }
    $modifiers = [System.Windows.Input.Keyboard]::Modifiers
    $ctrl  = ($modifiers -band [System.Windows.Input.ModifierKeys]::Control) -ne 0
    $shift = ($modifiers -band [System.Windows.Input.ModifierKeys]::Shift) -ne 0
    $visible = [System.Windows.Visibility]::Visible

    foreach ($overlay in @(@('ImportOverlay', 'Import'), @('ChangelogOverlay', 'Changelog'), @('AboutOverlay', 'About'))) {
        if ($ui[$overlay[0]].Visibility -eq $visible) {
            if ($key -eq [System.Windows.Input.Key]::Escape) {
                switch ($overlay[1]) { 'Import' { Hide-ImportDialog } 'Changelog' { Hide-Changelog } 'About' { Hide-About } }
                return $true
            }
            return $false
        }
    }
    if ($key -eq [System.Windows.Input.Key]::F1) { Show-About; return $true }
    if (-not $ctrl) { return $false }

    $name = $key.ToString()
    $tabs = @{ D1 = 0; D2 = 1; D3 = 2; D4 = 3; D5 = 4; D6 = 5; NumPad1 = 0; NumPad2 = 1; NumPad3 = 2; NumPad4 = 3; NumPad5 = 4; NumPad6 = 5 }
    if ($tabs.ContainsKey($name) -and -not $shift) { $ui.MainTabs.SelectedIndex = $tabs[$name]; return $true }
    $selected = $ui.MainTabs.SelectedItem
    $onBatch = $selected -eq $ui.TabBatch
    switch ($name) {
        'OemComma' { $ui.MainTabs.SelectedItem = $ui.TabSettings; return $true }
        'L' {
            $ui.MainTabs.SelectedItem = $ui.TabSingle
            [void]$script:Window.Dispatcher.BeginInvoke([Action]{ [void]$script:UI.EmailBox.Focus(); $script:UI.EmailBox.SelectAll() }, [System.Windows.Threading.DispatcherPriority]::Input)
            return $true
        }
        'C' { if ($shift -and $null -ne $script:LastResult) { Copy-Field -Value $script:LastResult.Summary -Name 'Result'; return $true } }
        'F' {
            $ui.MainTabs.SelectedItem = $ui.TabHistory
            [void]$script:Window.Dispatcher.BeginInvoke([Action]{ [void]$script:UI.HistorySearch.Focus(); $script:UI.HistorySearch.SelectAll() }, [System.Windows.Threading.DispatcherPriority]::Input)
            return $true
        }
        'E' {
            if ($onBatch) { Export-BatchResults -Format 'CSV'; return $true }
            if ($selected -eq $ui.TabHistory) { Export-Items -Items (Get-ShownHistory) -Format 'CSV' -BaseName 'check_history'; return $true }
        }
        'D' { if ($ui.DarkToggle.IsEnabled) { $ui.DarkToggle.IsChecked = -not [bool]$ui.DarkToggle.IsChecked; return $true } }
        'Return' { if ($onBatch -and -not $script:Jobs.ContainsKey('Batch')) { Start-Batch; return $true } }
        'P' { if ($onBatch -and $script:Jobs.ContainsKey('Batch')) { Switch-BatchPause; return $true } }
        'S' { if ($onBatch -and $shift -and $script:Jobs.ContainsKey('Batch')) { Stop-Batch; return $true } }
        'R' { if ($onBatch -and $ui.BatchRetryButton.IsEnabled) { Start-BatchRetry; return $true } }
        'G' { if ($onBatch) { Set-BatchView -Grouped (-not $script:BatchGrouped); return $true } }
    }
    return $false
}
#endregion

#region 2.0 features: privacy, change detection, domains, imports, watchers, updates
function Sync-SharedSettings {
    $s = $script:Settings
    $script:CacheHours = [double]$s.StableCacheHours
    $script:Shared['CachePolicy'] = @{ StableHours = [double]$s.StableCacheHours; InconclusiveHours = [double]$s.InconclusiveCacheHours; FailureMinutes = [double]$s.FailureCacheMinutes }
    $script:Shared['RateLimitPerMinute'] = [int]$s.ApiRateLimitPerMinute
    $script:Shared['BatchDelayMs'] = [int]$s.BatchDelayMs
}

function Write-DataFile {
    param([string]$Path, [string]$Text)
    if ($script:Settings.EncryptStoredData -and ('MsAccCheckerNative3' -as [type])) {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
        $payload = $script:EncryptedPrefix + [Convert]::ToBase64String([MsAccCheckerNative3]::Protect($bytes))
        Set-Content -LiteralPath $Path -Value $payload -Encoding ASCII -NoNewline
    } else {
        Set-Content -LiteralPath $Path -Value $Text -Encoding UTF8
    }
}

function Read-DataFile {
    param([string]$Path)
    $text = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
    if ($text -and $text.StartsWith($script:EncryptedPrefix)) {
        $bytes = [Convert]::FromBase64String($text.Substring($script:EncryptedPrefix.Length).Trim())
        return [System.Text.Encoding]::UTF8.GetString([MsAccCheckerNative3]::Unprotect($bytes))
    }
    return $text
}

function Set-EncryptStoredData {
    param([bool]$On)
    if ($On -and -not ('MsAccCheckerNative3' -as [type])) {
        Set-ToggleSilently -Toggle $script:UI.EncryptStoredCheck -Value $false
        Show-AppError 'Windows data protection is not available in this PowerShell session, so the files cannot be encrypted.'
        return
    }
    $script:Settings.EncryptStoredData = $On
    Save-Settings
    Save-History
    Save-Cache
    if ($On) { Set-Status 'History and the cache are now encrypted for your Windows account.' 'Success' } else { Set-Status 'History and the cache are now saved without encryption.' 'Info' }
}

function Set-PrivateSession {
    param([bool]$On)
    $script:PrivateSession = $On
    $script:UI.StatusPrivate.Visibility = if ($On) { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }
    if ($On) { Set-Status 'Private session on. New checks will not be saved to history, the cache or logs.' 'Warn' }
    else { Set-Status 'Private session off. Checks made while it was on stay unsaved.' 'Info' }
}

function Set-CompactTables {
    param([bool]$On)
    $script:Settings.CompactTables = $On
    foreach ($name in @('BatchGrid', 'DomainGrid', 'HistoryGrid', 'ApiLogGrid', 'DiagEndpointGrid')) {
        $grid = $script:UI[$name]
        if ($null -eq $grid) { continue }
        $grid.RowHeight = if ($On) { 24 } else { 32 }
        $grid.FontSize  = if ($On) { 12 } else { 13 }
    }
}

function Set-ResultChanges {
    param($Result)
    $key = ([string]$Result.Username).ToLowerInvariant()
    $previous = $script:LatestByUser[$key]
    $changes = ''
    if ($null -ne $previous -and [string]$Result.Source -ne 'Cache') { $changes = Get-ResultChanges -Previous $previous -Current $Result }
    $since = if ($changes) { [string]$previous.Timestamp } else { '' }
    $Result | Add-Member -NotePropertyName 'Changes' -NotePropertyValue $changes -Force
    $Result | Add-Member -NotePropertyName 'ChangedSince' -NotePropertyValue $since -Force
    if ($script:PrivateSession -and $key) { [void]$script:PrivateKeys.Add($key) }
    return $changes
}

function Show-ChangeCallout {
    param($Result)
    $ui = $script:UI
    $changes = [string]$Result.Changes
    if ($changes) {
        $ui.ChangeCalloutTitle.Text = "Changed since the previous check on $($Result.ChangedSince)"
        $ui.ChangeCalloutText.Text  = $changes.Substring(0, 1).ToUpper() + $changes.Substring(1) + '.'
        $ui.ChangeCallout.Visibility = [System.Windows.Visibility]::Visible
    } else {
        $ui.ChangeCallout.Visibility = [System.Windows.Visibility]::Collapsed
    }
}

function Show-ChangeAlert {
    param($Result)
    Set-Status ("Change detected for {0}: {1}" -f $Result.Username, $Result.Changes) 'Warn'
    if ($null -ne $script:Tray -and -not $script:Window.IsVisible -and $script:AlertClock.Elapsed.TotalSeconds -ge 30) {
        $script:AlertClock.Restart()
        $script:Tray.ShowBalloonTip(5000, 'Account change detected', ("{0}: {1}" -f $Result.Username, $Result.Changes), [System.Windows.Forms.ToolTipIcon]::Warning)
    }
}

function Get-BatchExtraFor {
    param([string]$Username)
    $bound = @{}
    $extra = $script:BatchExtras[([string]$Username).ToLowerInvariant()]
    for ($i = 0; $i -lt $script:BatchExtraColumns.Count; $i++) {
        $bound["c$i"] = if ($null -ne $extra) { [string]$extra[$script:BatchExtraColumns[$i]] } else { '' }
    }
    return $bound
}

function Set-BatchView {
    param([bool]$Grouped)
    $ui = $script:UI
    $script:BatchGrouped = $Grouped
    $ui.BatchGrid.Visibility  = if ($Grouped) { [System.Windows.Visibility]::Collapsed } else { [System.Windows.Visibility]::Visible }
    $ui.DomainGrid.Visibility = if ($Grouped) { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }
    if ($Grouped) { Update-DomainView }
    if ([bool]$ui.GroupByDomainCheck.IsChecked -ne $Grouped) { Set-ToggleSilently -Toggle $ui.GroupByDomainCheck -Value $Grouped }
}

function Update-DomainView {
    $script:DomainRows.Clear()
    foreach ($row in (Get-DomainSummary -Items $script:BatchRows)) { $script:DomainRows.Add($row) }
    $script:DomainDirty = $false
}

function Set-BatchExtraColumns {
    param([string[]]$Columns, $Extras)
    $grid = $script:UI.BatchGrid
    foreach ($column in @($script:BatchExtraGridColumns)) { [void]$grid.Columns.Remove($column) }
    $script:BatchExtraGridColumns = New-Object System.Collections.Generic.List[object]
    $script:BatchExtraColumns = @($Columns | Where-Object { $_ })
    $script:BatchExtras = if ($null -ne $Extras) { $Extras } else { @{} }
    for ($i = 0; $i -lt $script:BatchExtraColumns.Count; $i++) {
        $column = New-Object System.Windows.Controls.DataGridTextColumn
        $column.Header  = $script:BatchExtraColumns[$i]
        $column.Binding = New-Object System.Windows.Data.Binding ("Extra[c$i]")
        $column.Width   = New-Object System.Windows.Controls.DataGridLength 130
        $column.ElementStyle = $script:Window.Resources['CellText']
        $grid.Columns.Insert([Math]::Min(2 + $i, $grid.Columns.Count), $column)
        $script:BatchExtraGridColumns.Add($column)
    }
}

function Set-BatchImport {
    param($Import, [string]$Name)
    $script:UI.BatchInput.Text = ($Import.Emails -join "`r`n")
    Set-BatchExtraColumns -Columns @($Import.KeepColumns) -Extras $Import.Extras
    $kept = if (@($Import.KeepColumns).Count -gt 0) { ' Keeping ' + (@($Import.KeepColumns) -join ', ') + '.' } else { '' }
    Set-Status ("Imported {0} addresses from {1}.{2}" -f $Import.Emails.Count, $Name, $kept) 'Success'
}

function Show-ImportDialog {
    param([string]$Path, $Import)
    $ui = $script:UI
    $script:PendingImport = @{ Path = $Path; Import = $Import }
    $ui.ImportFileText.Text = '{0}: {1} columns.' -f [System.IO.Path]::GetFileName($Path), @($Import.Columns).Count
    $script:SuppressImportRefresh = $true
    try {
        $ui.ImportEmailColumn.Items.Clear()
        foreach ($column in $Import.Columns) { [void]$ui.ImportEmailColumn.Items.Add([string]$column) }
        if ($Import.EmailColumn) { $ui.ImportEmailColumn.SelectedItem = [string]$Import.EmailColumn } else { $ui.ImportEmailColumn.SelectedIndex = 0 }
    } finally {
        $script:SuppressImportRefresh = $false
    }
    Update-ImportColumnChoices
    $ui.ImportOverlay.Visibility = [System.Windows.Visibility]::Visible
    [void]$ui.ImportEmailColumn.Focus()
}

function Update-ImportColumnChoices {
    if ($script:SuppressImportRefresh -or $null -eq $script:PendingImport) { return }
    $ui = $script:UI
    $import = $script:PendingImport.Import
    $emailColumn = [string]$ui.ImportEmailColumn.SelectedItem
    $ui.ImportColumnsPanel.Children.Clear()
    $style = $script:Window.Resources['BodySwitch']
    foreach ($column in $import.Columns) {
        if ([string]$column -eq $emailColumn) { continue }
        $box = New-Object System.Windows.Controls.CheckBox
        $box.Style = $style
        $box.Content = [string]$column
        $box.Tag = [string]$column
        $box.IsChecked = $true
        $box.Margin = [System.Windows.Thickness]::new(0, 4, 18, 4)
        [void]$ui.ImportColumnsPanel.Children.Add($box)
    }
    $preview = Import-EmailFile -Path $script:PendingImport.Path -EmailColumn $emailColumn -KeepColumns @()
    $ui.ImportPreviewText.Text = "$($preview.Emails.Count) addresses in the '$emailColumn' column."
}

function Complete-ImportDialog {
    $ui = $script:UI
    if ($null -eq $script:PendingImport) { return }
    $emailColumn = [string]$ui.ImportEmailColumn.SelectedItem
    $keep = @(foreach ($box in $ui.ImportColumnsPanel.Children) { if ($box.IsChecked) { [string]$box.Tag } })
    $path = $script:PendingImport.Path
    $import = Import-EmailFile -Path $path -EmailColumn $emailColumn -KeepColumns $keep
    Hide-ImportDialog
    Set-BatchImport -Import $import -Name ([System.IO.Path]::GetFileName($path))
}

function Hide-ImportDialog {
    $script:UI.ImportOverlay.Visibility = [System.Windows.Visibility]::Collapsed
    $script:PendingImport = $null
}

function Save-ImportTemplate {
    $dialog = New-Object Microsoft.Win32.SaveFileDialog
    $dialog.Filter = 'CSV files (*.csv)|*.csv'
    $dialog.FileName = 'email_import_template.csv'
    if ($dialog.ShowDialog($script:Window) -ne $true) { return }
    $lines = @('Email,EmployeeId,DisplayName,Department', 'ada@contoso.com,1001,Ada Lovelace,Engineering', 'grace@fabrikam.com,1002,Grace Hopper,Research')
    Set-Content -LiteralPath $dialog.FileName -Value $lines -Encoding UTF8
    Set-Status "Template saved to $($dialog.FileName). Add your own columns; they are kept with the results." 'Success'
}

function Invoke-ClipboardWatch {
    if (-not [bool]$script:UI.ClipboardWatchCheck.IsChecked -or $script:Jobs.ContainsKey('Batch')) { return }
    $sequence = 0
    try { $sequence = [MsAccCheckerNative3]::GetClipboardSequenceNumber() } catch { return }
    if ($sequence -eq $script:ClipboardSeq) { return }
    $script:ClipboardSeq = $sequence
    $text = Get-ClipboardText
    if (-not $text) { return }
    $tokens = Get-BatchTokens -Text $script:UI.BatchInput.Text
    $known = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($t in $tokens.Valid) { [void]$known.Add($t) }
    foreach ($t in $tokens.Invalid) { [void]$known.Add($t) }
    $added = New-Object System.Collections.Generic.List[string]
    foreach ($match in $script:EmailFinder.Matches($text)) { if ($known.Add($match.Value)) { $added.Add($match.Value) } }
    if ($added.Count -eq 0) { return }
    $box = $script:UI.BatchInput
    if ($box.Text -and -not $box.Text.EndsWith("`n")) { $box.AppendText("`r`n") }
    $box.AppendText(($added -join "`r`n"))
    Set-Status "Added $($added.Count) addresses from the clipboard." 'Info'
}

function Update-WatchStatus {
    param([string]$Text)
    $script:UI.WatchStatusText.Text = $Text
}

function Select-WatchFolder {
    try { Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop } catch { }
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = 'Choose the folder to watch for new email lists'
    if ($script:Settings.WatchFolder -and (Test-Path -LiteralPath $script:Settings.WatchFolder)) { $dialog.SelectedPath = $script:Settings.WatchFolder }
    if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
    $script:Settings.WatchFolder = $dialog.SelectedPath
    $script:UI.WatchFolderBox.Text = $dialog.SelectedPath
    Save-Settings
    Update-WatchStatus $(if ($script:Settings.WatchFolderEnabled) { "Watching $($dialog.SelectedPath)." } else { 'Turn on the switch below to start watching.' })
}

function Invoke-WatchFolder {
    $s = $script:Settings
    if (-not $s.WatchFolderEnabled -or -not $s.WatchFolder -or $script:Jobs.ContainsKey('Batch')) { return }
    if (-not (Test-Path -LiteralPath $s.WatchFolder -PathType Container)) { Update-WatchStatus 'The watched folder does not exist.'; return }
    $file = Get-ChildItem -LiteralPath $s.WatchFolder -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Extension -eq '.csv' -or $_.Extension -eq '.txt' } | Sort-Object LastWriteTime | Select-Object -First 1
    if ($null -eq $file) { return }
    if (((Get-Date) - $file.LastWriteTime).TotalSeconds -lt 3) { return }   # still being written
    try {
        $processing = Join-Path $s.WatchFolder 'processing'
        if (-not (Test-Path -LiteralPath $processing)) { [void](New-Item -ItemType Directory -Path $processing -Force) }
        $target = Join-Path $processing $file.Name
        Move-Item -LiteralPath $file.FullName -Destination $target -Force
        $import = Import-EmailFile -Path $target
        if ($import.Emails.Count -eq 0) {
            $processed = Join-Path $s.WatchFolder 'processed'
            if (-not (Test-Path -LiteralPath $processed)) { [void](New-Item -ItemType Directory -Path $processed -Force) }
            Move-Item -LiteralPath $target -Destination (Join-Path $processed ("{0}_empty_{1}{2}" -f $file.BaseName, (Get-Date -Format 'yyyyMMdd_HHmmss'), $file.Extension)) -Force
            Update-WatchStatus "$($file.Name) had no email addresses."
            return
        }
        Set-BatchImport -Import $import -Name $file.Name
        Update-WatchStatus "Checking $($import.Emails.Count) addresses from $($file.Name)..."
        Start-Batch -WatchFile $target
    } catch {
        Update-WatchStatus "Could not process $($file.Name): $(Get-InnerMessage $_.Exception)"
    }
}

function Complete-WatchFile {
    param($Job)
    $s = $script:Settings
    $file = [string]$Job.WatchFile
    try {
        $results = Join-Path $s.WatchFolder 'results'
        $processed = Join-Path $s.WatchFolder 'processed'
        foreach ($dir in @($results, $processed)) { if (-not (Test-Path -LiteralPath $dir)) { [void](New-Item -ItemType Directory -Path $dir -Force) } }
        $base  = [System.IO.Path]::GetFileNameWithoutExtension($file)
        $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
        $out   = Join-Path $results ("{0}_results_{1}.csv" -f $base, $stamp)
        $rows  = @(Get-ExportRows -Items @($script:BatchRows) -ExtraColumns $script:BatchExtraColumns -Redact:([bool]$s.RedactExports))
        $rows | Export-Csv -LiteralPath $out -NoTypeInformation -Encoding UTF8
        Register-ExportFile -Path $out
        Move-Item -LiteralPath $file -Destination (Join-Path $processed ("{0}_{1}{2}" -f $base, $stamp, [System.IO.Path]::GetExtension($file))) -Force
        Update-WatchStatus ("Last file: {0}. {1} results saved to the results folder at {2}." -f [System.IO.Path]::GetFileName($file), $rows.Count, (Get-Date -Format 'HH:mm'))
        if ($null -ne $script:Tray -and -not $script:Window.IsVisible) {
            $script:Tray.ShowBalloonTip(4000, $script:AppName, "Checked $($rows.Count) addresses from $([System.IO.Path]::GetFileName($file)).", [System.Windows.Forms.ToolTipIcon]::Info)
        }
    } catch {
        Update-WatchStatus "Could not save the results for $([System.IO.Path]::GetFileName($file)): $(Get-InnerMessage $_.Exception)"
    }
}

function Get-ExportedFiles {
    if (-not (Test-Path -LiteralPath $script:ExportsFile)) { return @() }
    try {
        $data = Get-Content -LiteralPath $script:ExportsFile -Raw -Encoding UTF8 | ConvertFrom-Json
        return @(foreach ($item in $data) { [string]$item })
    } catch { return @() }
}

function Register-ExportFile {
    param([string]$Path)
    try {
        $list = @(@(Get-ExportedFiles) + $Path | Where-Object { $_ } | Select-Object -Unique)
        ConvertTo-Json -InputObject $list | Set-Content -LiteralPath $script:ExportsFile -Encoding UTF8
    } catch { }
    Update-ExportCount
}

function Update-ExportCount {
    $count = @(Get-ExportedFiles | Where-Object { Test-Path -LiteralPath $_ }).Count
    $script:UI.ExportDeleteButton.Content   = if ($count -gt 0) { "Delete exported files ($count)" } else { 'Delete exported files' }
    $script:UI.ExportDeleteButton.IsEnabled = $count -gt 0
}

function Remove-ExportedFiles {
    $files = @(Get-ExportedFiles | Where-Object { Test-Path -LiteralPath $_ })
    if ($files.Count -eq 0) { Set-Status 'There are no exported files to delete.' 'Info'; return }
    $answer = [System.Windows.MessageBox]::Show($script:Window, "Delete the $($files.Count) files this app has exported? This cannot be undone.`n`n" + (($files | Select-Object -First 8) -join "`n"),
        $script:AppName, [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
    if ($answer -ne [System.Windows.MessageBoxResult]::Yes) { return }
    $deleted = 0
    foreach ($file in $files) { try { Remove-Item -LiteralPath $file -Force -ErrorAction Stop; $deleted++ } catch { } }
    try { '[]' | Set-Content -LiteralPath $script:ExportsFile -Encoding UTF8 } catch { }
    Update-ExportCount
    Set-Status "Deleted $deleted exported files." 'Success'
}

function Find-PdfBrowser {
    $candidates = New-Object System.Collections.Generic.List[string]
    foreach ($root in @(${env:ProgramFiles(x86)}, $env:ProgramFiles, $env:LOCALAPPDATA)) {
        if (-not $root) { continue }
        $candidates.Add((Join-Path $root 'Microsoft\Edge\Application\msedge.exe'))
        $candidates.Add((Join-Path $root 'Google\Chrome\Application\chrome.exe'))
    }
    foreach ($candidate in $candidates) { if (Test-Path -LiteralPath $candidate) { return $candidate } }
    return $null
}

function Convert-HtmlToPdf {
    param([string]$HtmlPath, [string]$PdfPath)
    $browser = Find-PdfBrowser
    if (-not $browser) { throw 'Microsoft Edge or Google Chrome is needed to create a PDF. Save the HTML report instead, open it and print it to PDF.' }
    if (Test-Path -LiteralPath $PdfPath) { Remove-Item -LiteralPath $PdfPath -Force }
    $profileDir = Join-Path ([System.IO.Path]::GetTempPath()) ('msac-pdf-' + [guid]::NewGuid().ToString('N'))
    $uri = ([Uri]$HtmlPath).AbsoluteUri
    $arguments = @('--headless=new', '--disable-gpu', '--no-first-run', '--no-pdf-header-footer', '--print-to-pdf-no-header',
                   ('--user-data-dir="{0}"' -f $profileDir), ('--print-to-pdf="{0}"' -f $PdfPath), ('"{0}"' -f $uri)) -join ' '
    $process = Start-Process -FilePath $browser -ArgumentList $arguments -PassThru -WindowStyle Hidden
    if (-not $process.WaitForExit(90000)) {
        try { $process.Kill() } catch { }
        throw 'Creating the PDF took too long. Save the HTML report and print it to PDF instead.'
    }
    Start-Sleep -Milliseconds 300
    Remove-Item -LiteralPath $profileDir -Recurse -Force -ErrorAction SilentlyContinue
    if (-not (Test-Path -LiteralPath $PdfPath)) { throw 'The browser did not create the PDF. Save the HTML report and print it to PDF instead.' }
}

function Export-BatchResults {
    param([string]$Format)
    if ($script:BatchGrouped -and ($Format -eq 'CSV' -or $Format -eq 'JSON')) {
        Export-Items -Items $script:DomainRows -Format $Format -BaseName 'domain_summary' -Raw
        return
    }
    $base = if ($Format -eq 'HTML' -or $Format -eq 'PDF') { 'batch_report' } else { 'batch_results' }
    Export-Items -Items $script:BatchRows -Format $Format -BaseName $base -ExtraColumns $script:BatchExtraColumns -Title 'Batch check report'
}

function Update-DiagChart {
    param($Timeline)
    $ui = $script:UI
    $canvas = $ui.DiagChart
    $canvas.Children.Clear()
    $points = @($Timeline | Where-Object { $null -ne $_.AverageMs })
    $failures = @($Timeline | Where-Object { $_.Failures -gt 0 })
    $ui.DiagChartEmpty.Visibility = if ($points.Count -gt 0 -or $failures.Count -gt 0) { [System.Windows.Visibility]::Collapsed } else { [System.Windows.Visibility]::Visible }
    if ($points.Count -eq 0 -and $failures.Count -eq 0) { $ui.DiagChartCaption.Text = ''; return }
    $width = $canvas.ActualWidth
    $height = $canvas.ActualHeight
    if ($width -lt 40 -or $height -lt 40) { return }
    $peak = 100
    foreach ($p in $points) { if ($p.AverageMs -gt $peak) { $peak = $p.AverageMs } }
    $step = $width / 29
    $toY = { param($ms) $height - 6 - ($ms / $peak) * ($height - 14) }

    $grid = New-Object System.Windows.Shapes.Line
    $grid.X1 = 0; $grid.X2 = $width; $grid.Y1 = & $toY ($peak / 2); $grid.Y2 = $grid.Y1
    $grid.StrokeThickness = 1
    $grid.StrokeDashArray = New-Object System.Windows.Media.DoubleCollection (,[double[]]@(4, 4))
    $grid.SetResourceReference([System.Windows.Shapes.Shape]::StrokeProperty, 'Border')
    [void]$canvas.Children.Add($grid)

    $line = $null
    foreach ($bucket in $Timeline) {
        $x = (29 - $bucket.MinutesAgo) * $step
        if ($bucket.Failures -gt 0) {
            $dot = New-Object System.Windows.Shapes.Ellipse
            $dot.Width = 7; $dot.Height = 7
            $dot.SetResourceReference([System.Windows.Shapes.Shape]::FillProperty, 'CatNotMicrosoft')
            $dot.ToolTip = "$($bucket.Failures) failed calls, $($bucket.MinutesAgo) minutes ago"
            [System.Windows.Controls.Canvas]::SetLeft($dot, $x - 3.5)
            [System.Windows.Controls.Canvas]::SetTop($dot, $height - 8)
            [void]$canvas.Children.Add($dot)
        }
        if ($null -eq $bucket.AverageMs) { $line = $null; continue }
        $y = & $toY $bucket.AverageMs
        if ($null -eq $line) {
            $line = New-Object System.Windows.Shapes.Polyline
            $line.StrokeThickness = 2
            $line.SetResourceReference([System.Windows.Shapes.Shape]::StrokeProperty, 'Primary')
            [void]$canvas.Children.Add($line)
        }
        $line.Points.Add([System.Windows.Point]::new($x, $y))
        $marker = New-Object System.Windows.Shapes.Ellipse
        $marker.Width = 6; $marker.Height = 6
        $marker.SetResourceReference([System.Windows.Shapes.Shape]::FillProperty, 'Primary')
        $marker.ToolTip = "$($bucket.AverageMs) ms average, $($bucket.Calls) calls, $($bucket.MinutesAgo) minutes ago"
        [System.Windows.Controls.Canvas]::SetLeft($marker, $x - 3)
        [System.Windows.Controls.Canvas]::SetTop($marker, $y - 3)
        [void]$canvas.Children.Add($marker)
    }
    $calls = ($Timeline | Measure-Object Calls -Sum).Sum
    $failed = ($Timeline | Measure-Object Failures -Sum).Sum
    $ui.DiagChartCaption.Text = "Average per minute, oldest on the left. Peak $peak ms; dashed line $([int]($peak / 2)) ms. $calls lookups, $failed failed (red dots)."
}

function Set-SuggestedBatchSettings {
    $advice = $script:LastAdvice
    if ($null -eq $advice) { return }
    $script:Settings.BatchConcurrency = [int]$advice.SuggestedParallel
    $script:Settings.BatchDelayMs = [int]$advice.SuggestedDelayMs
    $script:UI.BatchConcurrencyBox.Text = [string]$script:Settings.BatchConcurrency
    $script:UI.BatchDelayBox.Text = [string]$script:Settings.BatchDelayMs
    Save-Settings
    Sync-SharedSettings
    Update-BatchSettingsText
    Update-Diagnostics
    Set-Status "Batch settings changed to $($script:Settings.BatchConcurrency) at a time, $($script:Settings.BatchDelayMs) ms apart." 'Success'
}

function Start-BackgroundTask {
    param([string]$Name, [scriptblock]$Script, [hashtable]$Parameters)
    if ($script:Tasks.ContainsKey($Name)) { return $false }
    $shell = New-WorkerShell
    [void]$shell.AddScript($Script.ToString())
    foreach ($key in $Parameters.Keys) { [void]$shell.AddParameter($key, $Parameters[$key]) }
    $script:Tasks[$Name] = @{ Shell = $shell; Handle = $shell.BeginInvoke() }
    return $true
}

function Invoke-BackgroundTasks {
    foreach ($name in @($script:Tasks.Keys)) {
        $task = $script:Tasks[$name]
        if (-not $task.Handle.IsCompleted) { continue }
        $script:Tasks.Remove($name)
        $output = $null
        try { $output = @($task.Shell.EndInvoke($task.Handle)) | Select-Object -Last 1 }
        catch { $output = [pscustomobject]@{ Ok = $false; Error = (Get-InnerMessage $_.Exception) } }
        Close-WorkerShell -Shell $task.Shell
        switch ($name) {
            'UpdateCheck'    { Receive-UpdateCheck -Result $output }
            'UpdateDownload' { Receive-UpdateDownload -Result $output }
        }
    }
}

function Start-UpdateCheck {
    param([switch]$Quiet)
    $script:UpdateQuiet = [bool]$Quiet
    $started = Start-BackgroundTask -Name 'UpdateCheck' -Script $script:UpdateCheckScript -Parameters @{ ApiUrl = "https://api.github.com/repos/$($script:AppRepo -replace '^https://github.com/', '')/releases/latest" }
    if ($started) {
        $script:UI.UpdateStatusText.Text = 'Checking GitHub for the latest release...'
        $script:UI.UpdateCheckButton.IsEnabled = $false
    }
}

function Receive-UpdateCheck {
    param($Result)
    $ui = $script:UI
    $ui.UpdateCheckButton.IsEnabled = $true
    if ($null -eq $Result -or -not $Result.Ok) {
        $reason = if ($null -ne $Result) { $Result.Error } else { 'no answer' }
        $ui.UpdateStatusText.Text = "Could not check for updates: $reason"
        return
    }
    $script:LatestRelease = $Result
    $latest = $null
    try { $latest = [version](([string]$Result.Tag).TrimStart('v', 'V')) } catch { $latest = $null }
    $current = [version]$script:AppVersion
    if ($null -ne $latest -and $latest -gt $current) {
        $ui.UpdateStatusText.Text = "Version $latest is available (you have $current), published $(([string]$Result.Published).Substring(0, [Math]::Min(10, ([string]$Result.Published).Length))). Select What's new to read the release notes."
        $ui.UpdateDownloadButton.Visibility = [System.Windows.Visibility]::Visible
        if ($script:UpdateQuiet) { Set-Status "Version $latest is available. See Settings, Updates." 'Info' }
    } else {
        $ui.UpdateStatusText.Text = "You have the latest version ($current). Checked $(Get-Date -Format 'yyyy-MM-dd HH:mm')."
        $ui.UpdateDownloadButton.Visibility = [System.Windows.Visibility]::Collapsed
    }
}

function Start-UpdateDownload {
    $release = $script:LatestRelease
    if ($null -eq $release) { return }
    $zip = @($release.Assets | Where-Object { $_.Name -like '*.zip' }) | Select-Object -First 1
    if ($null -eq $zip) { Open-Url $release.Url; return }
    $sums = @($release.Assets | Where-Object { $_.Name -match '^SHA256SUMS' }) | Select-Object -First 1
    $sumsUrl = if ($null -ne $sums) { [string]$sums.Url } else { '' }
    $folder = Join-Path ([Environment]::GetFolderPath('UserProfile')) 'Downloads'
    $started = Start-BackgroundTask -Name 'UpdateDownload' -Script $script:UpdateDownloadScript -Parameters @{ ZipUrl = [string]$zip.Url; ZipName = [string]$zip.Name; SumsUrl = $sumsUrl; Folder = $folder }
    if ($started) {
        $script:UI.UpdateDownloadButton.IsEnabled = $false
        $script:UI.UpdateStatusText.Text = "Downloading $($zip.Name)..."
    }
}

function Receive-UpdateDownload {
    param($Result)
    $ui = $script:UI
    $ui.UpdateDownloadButton.IsEnabled = $true
    if ($null -eq $Result -or -not $Result.Ok) {
        $reason = if ($null -ne $Result) { $Result.Error } else { 'no answer' }
        $ui.UpdateStatusText.Text = "Download failed: $reason"
        return
    }
    $check = if ($Result.Verified) { 'The SHA-256 checksum matches the published release.' } else { 'The release has no published checksum, so the download could not be verified.' }
    $ui.UpdateStatusText.Text = "Downloaded to $($Result.Path). $check Script signature: $($Result.Signature). Close the app, extract the zip and replace your files to update."
    try { Start-Process explorer.exe -ArgumentList ('/select,"{0}"' -f $Result.Path) } catch { }
}

function Get-ChangelogText {
    foreach ($path in @((Join-Path (Split-Path -Parent $script:AppDir) 'CHANGELOG.md'), (Join-Path $script:AppDir 'CHANGELOG.md'))) {
        if (Test-Path -LiteralPath $path) { return [string](Get-Content -LiteralPath $path -Raw -Encoding UTF8) }
    }
    return ''
}

function Show-Changelog {
    $ui = $script:UI
    $parts = @()
    if ($null -ne $script:LatestRelease -and $script:LatestRelease.Notes) {
        $parts += "Latest release on GitHub: $($script:LatestRelease.Tag)`r`n`r`n$($script:LatestRelease.Notes)`r`n`r`n----------------------------------------`r`n"
    }
    $local = Get-ChangelogText
    if ($local) { $parts += $local } else { $parts += 'The changelog file is not next to the app. Select Open on GitHub to read it.' }
    $ui.ChangelogTitle.Text = "What's new - you have version $($script:AppVersion)"
    $ui.ChangelogText.Text = ($parts -join '')
    $ui.ChangelogText.ScrollToHome()
    $ui.ChangelogOverlay.Visibility = [System.Windows.Visibility]::Visible
    [void]$ui.ChangelogCloseButton.Focus()
}

function Hide-Changelog {
    $script:UI.ChangelogOverlay.Visibility = [System.Windows.Visibility]::Collapsed
}
#endregion

#region System tray and window lifetime
function New-AppIconBitmap {
    param([int]$Size = 32)
    $bitmap   = New-Object System.Drawing.Bitmap $Size, $Size
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $graphics.SmoothingMode     = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
    $rect  = New-Object System.Drawing.Rectangle 0, 0, $Size, $Size
    $start = [System.Drawing.ColorTranslator]::FromHtml('#3659C9')
    $end   = [System.Drawing.ColorTranslator]::FromHtml('#5A7BEA')
    $brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush($rect, $start, $end, [single]45)
    $radius = [int]($Size * 0.28)
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $path.AddArc(0, 0, $radius, $radius, 180, 90)
    $path.AddArc($Size - $radius - 1, 0, $radius, $radius, 270, 90)
    $path.AddArc($Size - $radius - 1, $Size - $radius - 1, $radius, $radius, 0, 90)
    $path.AddArc(0, $Size - $radius - 1, $radius, $radius, 90, 90)
    $path.CloseFigure()
    $graphics.FillPath($brush, $path)
    $font = New-Object System.Drawing.Font('Segoe UI', [single]($Size * 0.58), [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
    $format = New-Object System.Drawing.StringFormat
    $format.Alignment     = [System.Drawing.StringAlignment]::Center
    $format.LineAlignment = [System.Drawing.StringAlignment]::Center
    $graphics.DrawString('@', $font, [System.Drawing.Brushes]::White, (New-Object System.Drawing.RectangleF 0, -1, $Size, $Size), $format)
    $format.Dispose(); $font.Dispose(); $path.Dispose(); $brush.Dispose(); $graphics.Dispose()
    return $bitmap
}

function Initialize-Tray {
    # The app works without a tray icon if any of this fails; it just exits when the window closes.
    try {
        Add-Type -AssemblyName System.Windows.Forms, System.Drawing -ErrorAction Stop
        [System.Windows.Forms.Application]::EnableVisualStyles()

        $bitmap = New-AppIconBitmap -Size 32
        $script:AppIconHandle = $bitmap.GetHicon()
        $icon = [System.Drawing.Icon]::FromHandle($script:AppIconHandle)
        try {
            $script:Window.Icon = [System.Windows.Interop.Imaging]::CreateBitmapSourceFromHIcon(
                $script:AppIconHandle, [System.Windows.Int32Rect]::Empty,
                [System.Windows.Media.Imaging.BitmapSizeOptions]::FromEmptyOptions())
        } catch { }

        $menu = New-Object System.Windows.Forms.ContextMenuStrip
        $script:TrayOpenWeb = $menu.Items.Add('Open web page')
        $script:TrayShow    = $menu.Items.Add('Show window')
        [void]$menu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
        $script:TrayApi     = $menu.Items.Add('Stop API server')
        [void]$menu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
        $script:TrayExit    = $menu.Items.Add('Exit')
        $script:TrayShow.Font = New-Object System.Drawing.Font($script:TrayShow.Font, [System.Drawing.FontStyle]::Bold)

        $script:TrayOpenWeb.Add_Click({ try { if ($null -ne $script:Api) { Open-Url "http://localhost:$($script:Api.Port)/" } } catch { } })
        $script:TrayShow.Add_Click({ try { Show-MainWindow } catch { } })
        $script:TrayApi.Add_Click({
            try {
                if ($null -ne $script:Api) { Stop-ApiServer } else { Start-ApiServer }
            } catch {
                Show-MainWindow
                Show-AppError $_ 'The API server could not be started.'
            }
        })
        $script:TrayExit.Add_Click({ try { Exit-App } catch { } })

        $tray = New-Object System.Windows.Forms.NotifyIcon
        $tray.Icon = $icon
        $tray.Text = $script:AppName
        $tray.ContextMenuStrip = $menu
        $tray.Visible = $false
        $tray.Add_MouseClick({
            param($sender, $e)
            if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) { try { Show-MainWindow } catch { } }
        })
        $tray.Add_BalloonTipClicked({ try { Show-MainWindow } catch { } })
        $script:Tray = $tray
    } catch {
        $script:Tray = $null
    }
}

function Update-Tray {
    if ($null -eq $script:Tray -or $script:CleanedUp) { return }
    $running = $null -ne $script:Api
    $hidden  = -not $script:Window.IsVisible
    $script:Tray.Visible = ($running -or $hidden)
    $text = if ($running) { "Account Checker - API on port $($script:Api.Port)" } else { $script:AppName }
    $script:Tray.Text = $text.Substring(0, [Math]::Min(63, $text.Length))
    $script:TrayApi.Text = if ($running) { 'Stop API server' } else { 'Start API server' }
    $script:TrayOpenWeb.Enabled = $running
}

function Test-KeepInTray {
    return (-not $script:ExitRequested) -and ($null -ne $script:Api) -and ($null -ne $script:Tray) -and [bool]$script:Settings.TrayOnClose
}

function Show-MainWindow {
    $w = $script:Window
    if (-not $w.IsVisible) { $w.Show() }
    if ($w.WindowState -eq [System.Windows.WindowState]::Minimized) { $w.WindowState = [System.Windows.WindowState]::Normal }
    [void]$w.Activate()
    Update-Tray
}

function Hide-ToTray {
    $script:Window.Hide()
    Update-Tray
    if (-not $script:TrayTipShown -and $null -ne $script:Api) {
        $script:TrayTipShown = $true
        $script:Tray.ShowBalloonTip(4000, $script:AppName,
            "Still running in the system tray. The API is at http://localhost:$($script:Api.Port)/. Right-click the icon for options.",
            [System.Windows.Forms.ToolTipIcon]::Info)
    }
}

function Invoke-AppCleanup {
    if ($script:CleanedUp) { return }
    $script:CleanedUp = $true
    try { $script:Timer.Stop() } catch { }
    foreach ($mode in @($script:Jobs.Keys)) {
        try { $job = $script:Jobs[$mode]; $job.Sync.Cancel = $true; $job.Sync.Paused = $false; Stop-JobShells -Job $job -WaitMs 500 } catch { }
    }
    $script:Jobs.Clear()
    try { Stop-ApiServer -Quiet } catch { }
    if ($null -ne $script:ApiTest) { try { Close-WorkerShell -Shell $script:ApiTest.Shell -WaitMs 200 -Handle $script:ApiTest.Handle } catch { } }
    if ($script:WindowCreated) { try { Save-WindowBounds } catch { } }
    if ($script:Settings.ClearOnExit) {
        try { $script:HistoryRows.Clear(); $script:Shared.History.Clear() } catch { }
        try {
            $cache = $script:Shared.Cache
            [System.Threading.Monitor]::Enter($cache.SyncRoot)
            try { $cache.Clear() } finally { [System.Threading.Monitor]::Exit($cache.SyncRoot) }
        } catch { }
    }
    foreach ($taskName in @($script:Tasks.Keys)) {
        try { Close-WorkerShell -Shell $script:Tasks[$taskName].Shell -WaitMs 200 -Handle $script:Tasks[$taskName].Handle } catch { }
    }
    if ($null -ne $script:HighContrastHandler) {
        try { [System.Windows.SystemParameters]::remove_StaticPropertyChanged($script:HighContrastHandler) } catch { }
    }
    try { Save-Settings } catch { }
    try { Save-History } catch { }
    try { Save-Cache } catch { }
    if ($null -ne $script:Tray) {
        try { $script:Tray.Visible = $false; $script:Tray.Dispose() } catch { }
    }
}

function Exit-App {
    $script:ExitRequested = $true
    if ($script:WindowCreated) {
        $script:Window.Close()
    } else {
        # The window was never shown (started with -ApiMode), so Close would not raise its events.
        Invoke-AppCleanup
        $script:Frame.Continue = $false
    }
}
#endregion

#region Build the window
$script:Settings      = Read-Settings
$script:UI            = @{}
$script:Jobs          = @{}
$script:Api           = $null
$script:ApiTest       = $null
$script:Hwnd          = [IntPtr]::Zero
$script:LastResult    = $null
$script:SuppressToggle = $false
$script:HistoryDirty  = $false
$script:LogErrorShown = $false
$script:BatchIndex    = 0
$script:BatchCounts   = New-Counts
$script:HistoryCounts = New-Counts
$script:HistoryFilterState = $null
$script:SuppressFilter = $false
$script:LogChecked    = @{}
$script:DiagClock     = [System.Diagnostics.Stopwatch]::StartNew()
$script:RetentionClock = [System.Diagnostics.Stopwatch]::StartNew()
$script:DiagRows      = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
$script:HighContrastHandler = $null
$script:EncryptedPrefix = 'MSAC-DPAPI-1:'
$script:ExportsFile   = Join-Path $script:ConfigDir 'exported_files.json'
$script:PrivateSession = $false
$script:PrivateKeys   = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
$script:LatestByUser  = @{}
$script:DomainRows    = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
$script:BatchGrouped  = $false
$script:DomainDirty   = $false
$script:BatchExtras   = @{}
$script:BatchExtraColumns = @()
$script:BatchExtraGridColumns = New-Object System.Collections.Generic.List[object]
$script:PendingImport = $null
$script:SuppressImportRefresh = $false
$script:ClipboardSeq  = 0
$script:Tasks         = @{}
$script:LatestRelease = $null
$script:UpdateQuiet   = $false
$script:LastAdvice    = $null
$script:ClipboardClock = [System.Diagnostics.Stopwatch]::StartNew()
$script:WatchClock    = [System.Diagnostics.Stopwatch]::StartNew()
$script:DomainClock   = [System.Diagnostics.Stopwatch]::StartNew()
$script:AlertClock    = [System.Diagnostics.Stopwatch]::StartNew()
$script:NumberFields  = @{
    HistoryRetentionBox = @{ Name = 'HistoryRetentionDays'; Min = 0; Max = 3650 }
    LogRetentionBox     = @{ Name = 'LogRetentionDays';     Min = 0; Max = 3650 }
    BatchDelayBox       = @{ Name = 'BatchDelayMs';         Min = 0; Max = 10000 }
    BatchConcurrencyBox = @{ Name = 'BatchConcurrency';     Min = 1; Max = 4 }
    StableCacheBox      = @{ Name = 'StableCacheHours';     Min = 0; Max = 720 }
    InconclusiveCacheBox = @{ Name = 'InconclusiveCacheHours'; Min = 0; Max = 720 }
    FailureCacheBox     = @{ Name = 'FailureCacheMinutes';  Min = 0; Max = 60 }
    StaleAfterBox       = @{ Name = 'StaleAfterHours';      Min = 1; Max = 720 }
    RateLimitBox        = @{ Name = 'ApiRateLimitPerMinute'; Min = 0; Max = 6000 }
}
$script:SaveClock     = [System.Diagnostics.Stopwatch]::StartNew()
$script:EmailFinder   = [regex]'[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}'
$script:BatchRows     = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
$script:HistoryRows   = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
$script:ApiLogRows    = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
$script:CategoryLabels = @{
    Entra = 'Work/School (Entra ID)'; MSA = 'Personal (MSA)'; Federated = 'Federated (Work/School)'
    Dual = 'Dual identity'; GCCHigh = 'GCC High'; China = 'M365 China'; NotMicrosoft = 'Not a Microsoft account'
    Error = 'Error'; Unknown = 'Unknown'
}
$script:Tray          = $null
$script:TrayTipShown  = $false
$script:ExitRequested = $false
$script:CleanedUp     = $false
$script:NoBrowser     = [bool]$NoBrowser
$script:WebPage       = $script:WebPageTemplate.Replace('{{VERSION}}', $script:AppVersion)
$script:Shared        = New-SharedState

try {
    [xml]$xamlDocument = $script:Xaml
    $script:Window = [System.Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xamlDocument))
} catch {
    $layoutError = "The window layout could not be loaded: $(Get-InnerMessage $_.Exception)"
    Write-Error $layoutError
    try { [void][System.Windows.MessageBox]::Show($layoutError, 'Microsoft Account Checker', 'OK', 'Error') } catch { }
    return
}

foreach ($match in [regex]::Matches($script:Xaml, 'x:Name="([A-Za-z0-9_]+)"')) {
    $name = $match.Groups[1].Value
    $element = $script:Window.FindName($name)
    if ($null -ne $element) { $script:UI[$name] = $element }
}

Set-Theme -Dark $script:Settings.DarkMode
Set-WindowBounds

$ui = $script:UI
$ui.BatchGrid.ItemsSource   = $script:BatchRows
$ui.HistoryGrid.ItemsSource = $script:HistoryRows
$ui.ApiLogGrid.ItemsSource  = $script:ApiLogRows
$script:HistoryView = [System.Windows.Data.CollectionViewSource]::GetDefaultView($script:HistoryRows)
$script:HistoryView.Filter = [Predicate[object]]{ param($item) return (Test-HistoryItem -Item $item) }
$ui.DiagEndpointGrid.ItemsSource = $script:DiagRows
New-GridContextMenu -Grid $ui.BatchGrid
New-GridContextMenu -Grid $ui.HistoryGrid

$ui.DarkToggle.IsChecked = $script:Settings.DarkMode
$ui.LogToggle.IsChecked  = $script:Settings.EnableLogging
$ui.ApiToggle.IsChecked  = $false
$ui.ApiPortBox.Text      = [string]$script:Settings.ApiPort
if ($Port -gt 0) { $ui.ApiPortBox.Text = [string]$Port }
$ui.ApiOpenBrowserCheck.IsChecked = $script:Settings.OpenBrowserOnApiStart
$ui.ApiTrayCheck.IsChecked        = $script:Settings.TrayOnClose
$ui.StatusVersion.Text   = "v$($script:AppVersion)"
$script:Window.Title     = "$($script:AppName) $($script:AppVersion)"
$ui.ApiTestResponse.Text = 'Turn on the API server, enter an address and select Send.'
$ui.SettingsDataFolder.Text       = $script:AppDir
$ui.HistoryRetentionBox.Text      = [string]$script:Settings.HistoryRetentionDays
$ui.LogRetentionBox.Text          = [string]$script:Settings.LogRetentionDays
$ui.BatchDelayBox.Text            = [string]$script:Settings.BatchDelayMs
$ui.BatchConcurrencyBox.Text      = [string]$script:Settings.BatchConcurrency
$ui.ClearOnExitCheck.IsChecked    = $script:Settings.ClearOnExit
$ui.BatchSkipCacheCheck.IsChecked = $script:Settings.BatchSkipCache
$ui.DomainGrid.ItemsSource        = $script:DomainRows
foreach ($boxName in @('StableCacheBox', 'InconclusiveCacheBox', 'FailureCacheBox', 'StaleAfterBox', 'RateLimitBox')) {
    $ui[$boxName].Text = [string]$script:Settings[$script:NumberFields[$boxName].Name]
}
$ui.EncryptStoredCheck.IsChecked  = $script:Settings.EncryptStoredData
$ui.RedactExportsCheck.IsChecked  = $script:Settings.RedactExports
$ui.WatchFolderBox.Text           = [string]$script:Settings.WatchFolder
$ui.WatchFolderCheck.IsChecked    = $script:Settings.WatchFolderEnabled
$ui.UpdateOnStartCheck.IsChecked  = $script:Settings.CheckUpdatesOnStart
$ui.CompactTablesCheck.IsChecked  = $script:Settings.CompactTables
Sync-SharedSettings

Read-History
Read-Cache
Invoke-Retention
Update-HistoryFilter
Update-HistoryStats
Initialize-BatchChips
Update-BatchSettingsText
Update-RetryButton
Initialize-Tray
Set-ApiTokenState -Required ([bool]$script:Settings.ApiRequireToken)
Set-CompactTables -On ([bool]$script:Settings.CompactTables)
Update-ExportCount
Update-WatchStatus $(if ($script:Settings.WatchFolderEnabled -and $script:Settings.WatchFolder) { "Watching $($script:Settings.WatchFolder)." } elseif ($script:Settings.WatchFolder) { 'Watching is off.' } else { 'No folder chosen.' })
New-GridContextMenu -Grid $ui.DomainGrid
if ($script:Settings.CheckUpdatesOnStart) { Start-UpdateCheck -Quiet }
Update-ApiUi
Reset-ResultView
Set-HintVisibility -TextBox $ui.EmailBox -Hint $ui.EmailHint
Set-HintVisibility -TextBox $ui.HistorySearch -Hint $ui.HistorySearchHint
Set-Status 'Ready' 'Info'
#endregion

#region Event handlers
# Each handler calls a named function inside try/catch so one failure never closes the window.

# Header switches
$ui.DarkToggle.Add_Checked({ if ($script:SuppressToggle) { return }; try { Set-Theme -Dark $true; Save-Settings; Set-Status 'Dark mode on.' 'Info' } catch { Show-AppError $_ } })
$ui.DarkToggle.Add_Unchecked({ if ($script:SuppressToggle) { return }; try { Set-Theme -Dark $false; Save-Settings; Set-Status 'Dark mode off.' 'Info' } catch { Show-AppError $_ } })

$ui.ApiToggle.Add_Checked({
    if ($script:SuppressToggle) { return }
    try { Start-ApiServer }
    catch { Set-ToggleSilently -Toggle $script:UI.ApiToggle -Value $false; Show-AppError $_ 'The API server could not be started.' }
})
$ui.ApiToggle.Add_Unchecked({ if ($script:SuppressToggle) { return }; try { Stop-ApiServer } catch { Show-AppError $_ } })

$ui.LogToggle.Add_Checked({
    if ($script:SuppressToggle) { return }
    try { $script:Settings.EnableLogging = $true; $script:LogErrorShown = $false; Save-Settings; Set-Status "Logging on. Checks are written to $(Get-TodayLogFile)" 'Info' } catch { Show-AppError $_ }
})
$ui.LogToggle.Add_Unchecked({
    if ($script:SuppressToggle) { return }
    try { $script:Settings.EnableLogging = $false; Save-Settings; Set-Status 'Logging off.' 'Info' } catch { Show-AppError $_ }
})

$ui.AboutButton.Add_Click({ try { Show-About } catch { Show-AppError $_ } })
$ui.SettingsButton.Add_Click({ $script:UI.MainTabs.SelectedItem = $script:UI.TabSettings })
$ui.AboutChangelogButton.Add_Click({ try { Hide-About; Show-Changelog } catch { Show-AppError $_ } })
$ui.AboutCloseButton.Add_Click({ Hide-About })
$ui.AboutCloseX.Add_Click({ Hide-About })
$ui.AboutRepoButton.Add_Click({ Open-Url $script:AppRepo })
$ui.AboutIssueButton.Add_Click({ Open-Url $script:AppIssues })
$ui.AboutFolderButton.Add_Click({ try { Start-Process explorer.exe -ArgumentList ('"{0}"' -f $script:AppDir) } catch { Show-AppError $_ } })
$ui.AboutOverlay.Add_MouseLeftButtonDown({
    param($sender, $e)
    if ($e.OriginalSource -eq $script:UI.AboutOverlay) { Hide-About }
})
$script:Window.Add_PreviewKeyDown({
    param($sender, $e)
    try { if (Invoke-Shortcut -KeyEvent $e) { $e.Handled = $true } } catch { Show-AppError $_ }
})

# Single check
$ui.CheckButton.Add_Click({ try { Invoke-SingleCheck } catch { Set-SingleBusy $false; Show-AppError $_ } })
$ui.EmailBox.Add_KeyDown({
    param($sender, $e)
    if ($e.Key -eq [System.Windows.Input.Key]::Return) {
        $e.Handled = $true
        try { Invoke-SingleCheck } catch { Set-SingleBusy $false; Show-AppError $_ }
    }
})
$ui.EmailBox.Add_TextChanged({ try { Set-HintVisibility -TextBox $script:UI.EmailBox -Hint $script:UI.EmailHint; Hide-EmailError } catch { } })
$ui.PasteButton.Add_Click({
    try {
        $text = (Get-ClipboardText).Trim()
        if ($text) { $script:UI.EmailBox.Text = $text; $script:UI.EmailBox.Focus() | Out-Null; $script:UI.EmailBox.CaretIndex = $text.Length }
        else { Set-Status 'The clipboard does not contain text.' 'Warn' }
    } catch { Show-AppError $_ }
})
$ui.ClearButton.Add_Click({
    try {
        if ($script:Jobs.ContainsKey('Single')) { return }
        $script:UI.EmailBox.Clear(); Hide-EmailError; Reset-ResultView; Set-Status 'Ready' 'Info'; $script:UI.EmailBox.Focus() | Out-Null
    } catch { Show-AppError $_ }
})
$ui.CopyButton.Add_Click({
    try {
        if ($null -eq $script:LastResult) { return }
        if (Set-ClipboardText $script:LastResult.Summary) { Set-Status 'Result copied to the clipboard.' 'Success' }
        else { Set-Status 'The clipboard is busy. Try Copy result again.' 'Warn' }
    } catch { Show-AppError $_ }
})
$ui.RecheckButton.Add_Click({ try { if ($null -ne $script:LastResult) { Invoke-SingleCheck -SkipCache -Email ([string]$script:LastResult.Username) } } catch { Set-SingleBusy $false; Show-AppError $_ } })
$ui.CopyCloudUrlButton.Add_Click({ if ($script:LastResult) { Copy-Field -Value ([string]$script:LastResult.CloudURL) -Name 'Cloud URL' } })
$ui.CopyFedButton.Add_Click({ if ($script:LastResult) { Copy-Field -Value ([string]$script:LastResult.FederationRedirectUrl) -Name 'FederationRedirectUrl' } })
$ui.CopyOrgButton.Add_Click({ if ($script:LastResult) { Copy-Field -Value ([string]$script:LastResult.Organization) -Name 'Organization' } })
$ui.CopyTenantButton.Add_Click({ if ($script:LastResult) { Copy-Field -Value ([string]$script:LastResult.TenantId) -Name 'Tenant ID' } })
$ui.RawToggleButton.Add_Click({
    try {
        $box = $script:UI.RawBox
        if ($box.Visibility -eq [System.Windows.Visibility]::Visible) {
            $box.Visibility = [System.Windows.Visibility]::Collapsed; $script:UI.RawToggleButton.Content = 'Show raw response'
        } else {
            $box.Visibility = [System.Windows.Visibility]::Visible; $script:UI.RawToggleButton.Content = 'Hide raw response'
        }
    } catch { Show-AppError $_ }
})

# Hyperlinks
foreach ($linkName in @('CloudUrlLink', 'FedUrlLink', 'ApiUrlLink', 'AboutAuthorLink', 'AboutRepoLink')) {
    $ui[$linkName].Add_RequestNavigate({
        param($sender, $e)
        $e.Handled = $true
        Open-Url $e.Uri.AbsoluteUri
    })
}

# Batch check
$ui.BatchInput.Add_TextChanged({ try { Update-BatchCount } catch { } })
$ui.BatchImportButton.Add_Click({ try { Import-BatchFile } catch { Show-AppError $_ 'The file could not be imported.' } })
$ui.BatchPasteButton.Add_Click({
    try {
        $text = Get-ClipboardText
        if (-not $text) { Set-Status 'The clipboard does not contain text.' 'Warn'; return }
        $box = $script:UI.BatchInput
        if ($box.Text -and -not $box.Text.EndsWith("`n")) { $box.AppendText("`r`n") }
        $box.AppendText($text)
    } catch { Show-AppError $_ }
})
$ui.BatchClearInputButton.Add_Click({ try { $script:UI.BatchInput.Clear(); Set-BatchExtraColumns -Columns @() -Extras @{} } catch { } })
$ui.BatchStartButton.Add_Click({ try { Start-Batch } catch { Set-BatchBusy $false; Show-AppError $_ 'The batch could not start.' } })
$ui.BatchStopButton.Add_Click({ try { Stop-Batch } catch { Show-AppError $_ } })
$ui.BatchPauseButton.Add_Click({ try { Switch-BatchPause } catch { Show-AppError $_ } })
$ui.BatchRetryButton.Add_Click({ try { Start-BatchRetry } catch { Set-BatchBusy $false; Show-AppError $_ 'The retry could not start.' } })
$ui.BatchSkipCacheCheck.Add_Checked({ $script:Settings.BatchSkipCache = $true; Save-Settings })
$ui.BatchSkipCacheCheck.Add_Unchecked({ $script:Settings.BatchSkipCache = $false; Save-Settings })
$ui.BatchClearResultsButton.Add_Click({ try { if (-not $script:Jobs.ContainsKey('Batch')) { Clear-BatchResults; Set-Status 'Batch results cleared.' 'Info' } } catch { Show-AppError $_ } })
$ui.BatchExportCsvButton.Add_Click({ try { Export-BatchResults -Format 'CSV' } catch { Show-AppError $_ 'Export failed.' } })
$ui.BatchExportJsonButton.Add_Click({ try { Export-BatchResults -Format 'JSON' } catch { Show-AppError $_ 'Export failed.' } })
$ui.BatchExportHtmlButton.Add_Click({ try { Export-BatchResults -Format 'HTML' } catch { Show-AppError $_ 'Export failed.' } })
$ui.BatchExportPdfButton.Add_Click({ try { Export-BatchResults -Format 'PDF' } catch { Show-AppError $_ 'The PDF could not be created.' } })
$ui.GroupByDomainCheck.Add_Checked({ if ($script:SuppressToggle) { return }; try { Set-BatchView -Grouped $true } catch { Show-AppError $_ } })
$ui.GroupByDomainCheck.Add_Unchecked({ if ($script:SuppressToggle) { return }; try { Set-BatchView -Grouped $false } catch { Show-AppError $_ } })
$ui.ClipboardWatchCheck.Add_Checked({
    try { $script:ClipboardSeq = [MsAccCheckerNative3]::GetClipboardSequenceNumber() } catch { }
    Set-Status 'Collecting addresses you copy from now on. Turn the switch off to stop.' 'Info'
})
$ui.ClipboardWatchCheck.Add_Unchecked({ Set-Status 'Stopped collecting addresses from the clipboard.' 'Info' })

# History
$ui.HistorySearch.Add_TextChanged({
    try {
        Set-HintVisibility -TextBox $script:UI.HistorySearch -Hint $script:UI.HistorySearchHint
        Update-HistoryFilter
    } catch { }
})
foreach ($comboName in @('FilterType', 'FilterCloud', 'FilterStatus', 'FilterFederation', 'FilterSource', 'FilterDate')) {
    $ui[$comboName].Add_SelectionChanged({ try { Update-HistoryFilter } catch { } })
}
$ui.FilterResetButton.Add_Click({ try { Reset-HistoryFilters } catch { Show-AppError $_ } })
$ui.HistoryPrivacyLink.Add_Click({ $script:UI.MainTabs.SelectedItem = $script:UI.TabSettings })
$ui.HistoryExportButton.Add_Click({ try { Export-Items -Items (Get-ShownHistory) -Format 'CSV' -BaseName 'check_history' } catch { Show-AppError $_ 'Export failed.' } })
$ui.HistoryExportHtmlButton.Add_Click({ try { Export-Items -Items (Get-ShownHistory) -Format 'HTML' -BaseName 'history_report' -Title 'History report' } catch { Show-AppError $_ 'Export failed.' } })
$ui.HistoryExportPdfButton.Add_Click({ try { Export-Items -Items (Get-ShownHistory) -Format 'PDF' -BaseName 'history_report' -Title 'History report' } catch { Show-AppError $_ 'The PDF could not be created.' } })
$ui.HistoryClearButton.Add_Click({ try { Clear-History } catch { Show-AppError $_ } })
$ui.CacheClearButton.Add_Click({ try { Clear-LookupCache } catch { Show-AppError $_ } })
$ui.LogOpenButton.Add_Click({ try { Open-TodayLog } catch { Show-AppError $_ } })
$ui.LogExportButton.Add_Click({ try { Export-TodayLog } catch { Show-AppError $_ 'The log could not be exported.' } })

# API tab
$ui.ApiPortBox.Add_TextChanged({ try { if ($null -eq $script:Api) { Update-ApiUi } } catch { } })
$ui.ApiTestButton.Add_Click({ try { Invoke-ApiTest } catch { $script:UI.ApiTestButton.IsEnabled = $true; Show-AppError $_ } })
$ui.ApiTestEmail.Add_KeyDown({
    param($sender, $e)
    if ($e.Key -eq [System.Windows.Input.Key]::Return) { $e.Handled = $true; try { Invoke-ApiTest } catch { Show-AppError $_ } }
})
$ui.ApiClearLogButton.Add_Click({ try { $script:ApiLogRows.Clear() } catch { } })
$ui.ApiOpenWebButton.Add_Click({ if ($null -ne $script:Api) { Open-Url "http://localhost:$($script:Api.Port)/" } })
$ui.ApiOpenBrowserCheck.Add_Checked({ $script:Settings.OpenBrowserOnApiStart = $true; Save-Settings })
$ui.ApiOpenBrowserCheck.Add_Unchecked({ $script:Settings.OpenBrowserOnApiStart = $false; Save-Settings })
$ui.ApiTrayCheck.Add_Checked({ $script:Settings.TrayOnClose = $true; Save-Settings })
$ui.ApiTrayCheck.Add_Unchecked({ $script:Settings.TrayOnClose = $false; Save-Settings })

# Diagnostics and settings
$ui.MainTabs.Add_SelectionChanged({
    param($sender, $e)
    if ($e.OriginalSource -eq $script:UI.MainTabs) { try { Update-Diagnostics } catch { } }
})
$ui.DiagResetButton.Add_Click({ try { Reset-CoreStats -Stats $script:Shared.Stats; $script:Shared['RejectedCount'] = 0; Update-Diagnostics; Set-Status 'Diagnostics counters reset.' 'Info' } catch { Show-AppError $_ } })
foreach ($boxName in @($script:NumberFields.Keys)) {
    $ui[$boxName].Add_LostFocus({ param($sender, $e) try { Set-NumberSetting -Box $sender } catch { } })
    $ui[$boxName].Add_KeyDown({ param($sender, $e) if ($e.Key -eq [System.Windows.Input.Key]::Return) { $e.Handled = $true; try { Set-NumberSetting -Box $sender } catch { } } })
}
$ui.ClearOnExitCheck.Add_Checked({ $script:Settings.ClearOnExit = $true; Save-Settings; Set-Status 'History and the cache will be cleared when the app closes.' 'Info' })
$ui.ClearOnExitCheck.Add_Unchecked({ $script:Settings.ClearOnExit = $false; Save-Settings })
$ui.LogDeleteButton.Add_Click({ try { Remove-AllLogs } catch { Show-AppError $_ } })
$ui.SettingsOpenFolderButton.Add_Click({ try { Start-Process explorer.exe -ArgumentList ('"{0}"' -f $script:AppDir) } catch { Show-AppError $_ } })
$ui.ApiTokenCheck.Add_Checked({ if ($script:SuppressToggle) { return }; try { Set-ApiTokenState -Required $true; Set-Status 'API token required. Copy it from the Settings tab for your scripts.' 'Info' } catch { Show-AppError $_ } })
$ui.ApiTokenCheck.Add_Unchecked({ if ($script:SuppressToggle) { return }; try { Set-ApiTokenState -Required $false; Set-Status 'API token no longer required.' 'Info' } catch { Show-AppError $_ } })
$ui.ApiTokenCopyButton.Add_Click({ Copy-Field -Value ([string]$script:Settings.ApiToken) -Name 'API token' })
$ui.ApiTokenNewButton.Add_Click({
    try {
        $answer = [System.Windows.MessageBox]::Show($script:Window, 'Create a new API token? Scripts using the current token will stop working until they are updated.', $script:AppName,
            [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
        if ($answer -eq [System.Windows.MessageBoxResult]::Yes) { Set-ApiTokenState -Required $true -NewToken; Set-Status 'New API token created.' 'Success' }
    } catch { Show-AppError $_ }
})

# Import dialog, What's new, updates
$ui.ImportEmailColumn.Add_SelectionChanged({ try { Update-ImportColumnChoices } catch { } })
$ui.ImportConfirmButton.Add_Click({ try { Complete-ImportDialog } catch { Hide-ImportDialog; Show-AppError $_ 'The file could not be imported.' } })
$ui.ImportCancelButton.Add_Click({ Hide-ImportDialog })
$ui.ImportTemplateButton.Add_Click({ try { Save-ImportTemplate } catch { Show-AppError $_ } })
$ui.ImportOverlay.Add_MouseLeftButtonDown({ param($sender, $e) if ($e.OriginalSource -eq $script:UI.ImportOverlay) { Hide-ImportDialog } })
$ui.ChangelogButton.Add_Click({ try { Show-Changelog } catch { Show-AppError $_ } })
$ui.ChangelogCloseButton.Add_Click({ Hide-Changelog })
$ui.ChangelogWebButton.Add_Click({ Open-Url "$($script:AppRepo)/blob/main/CHANGELOG.md" })
$ui.ChangelogOverlay.Add_MouseLeftButtonDown({ param($sender, $e) if ($e.OriginalSource -eq $script:UI.ChangelogOverlay) { Hide-Changelog } })
$ui.UpdateCheckButton.Add_Click({ try { Start-UpdateCheck } catch { Show-AppError $_ } })
$ui.UpdateDownloadButton.Add_Click({ try { Start-UpdateDownload } catch { Show-AppError $_ } })
$ui.UpdateOnStartCheck.Add_Checked({ $script:Settings.CheckUpdatesOnStart = $true; Save-Settings })
$ui.UpdateOnStartCheck.Add_Unchecked({ $script:Settings.CheckUpdatesOnStart = $false; Save-Settings })

# Privacy, display, watched folder, diagnostics
$ui.EncryptStoredCheck.Add_Checked({ if ($script:SuppressToggle) { return }; try { Set-EncryptStoredData -On $true } catch { Show-AppError $_ } })
$ui.EncryptStoredCheck.Add_Unchecked({ if ($script:SuppressToggle) { return }; try { Set-EncryptStoredData -On $false } catch { Show-AppError $_ } })
$ui.PrivateSessionCheck.Add_Checked({ Set-PrivateSession -On $true })
$ui.PrivateSessionCheck.Add_Unchecked({ Set-PrivateSession -On $false })
$ui.RedactExportsCheck.Add_Checked({ $script:Settings.RedactExports = $true; Save-Settings; Set-Status 'Exports will redact email addresses and imported columns.' 'Info' })
$ui.RedactExportsCheck.Add_Unchecked({ $script:Settings.RedactExports = $false; Save-Settings })
$ui.ExportDeleteButton.Add_Click({ try { Remove-ExportedFiles } catch { Show-AppError $_ } })
$ui.CompactTablesCheck.Add_Checked({ Set-CompactTables -On $true; Save-Settings })
$ui.CompactTablesCheck.Add_Unchecked({ Set-CompactTables -On $false; Save-Settings })
$ui.WatchFolderBrowseButton.Add_Click({ try { Select-WatchFolder } catch { Show-AppError $_ } })
$ui.WatchFolderCheck.Add_Checked({
    if (-not $script:Settings.WatchFolder) { Set-ToggleSilently -Toggle $script:UI.WatchFolderCheck -Value $false; Set-Status 'Choose a folder first.' 'Warn'; return }
    if ($script:SuppressToggle) { return }
    $script:Settings.WatchFolderEnabled = $true; Save-Settings; Update-WatchStatus "Watching $($script:Settings.WatchFolder)."
})
$ui.WatchFolderCheck.Add_Unchecked({ if ($script:SuppressToggle) { return }; $script:Settings.WatchFolderEnabled = $false; Save-Settings; Update-WatchStatus 'Watching is off.' })
$ui.DiagApplyButton.Add_Click({ try { Set-SuggestedBatchSettings } catch { Show-AppError $_ } })
$ui.DiagChart.Add_SizeChanged({ try { Update-Diagnostics } catch { } })

# Follow Windows high contrast changes while the app is open.
try {
    $script:HighContrastHandler = [System.ComponentModel.PropertyChangedEventHandler]{
        param($sender, $e)
        if ($e.PropertyName -eq 'HighContrast') { try { Set-Theme -Dark ([bool]$script:Settings.DarkMode) } catch { } }
    }
    [System.Windows.SystemParameters]::add_StaticPropertyChanged($script:HighContrastHandler)
} catch { $script:HighContrastHandler = $null }

# Window lifetime
$script:Window.Add_SourceInitialized({
    try {
        $script:Hwnd = (New-Object System.Windows.Interop.WindowInteropHelper $script:Window).Handle
        $script:WindowCreated = $true
        Set-TitleBarTheme -Dark ([bool]$script:Settings.DarkMode)
    } catch { }
})
$script:Window.Add_Loaded({ try { $script:UI.EmailBox.Focus() | Out-Null } catch { } })
$script:Window.Add_Closing({
    param($sender, $e)
    if (Test-KeepInTray) {
        $e.Cancel = $true
        Hide-ToTray
        return
    }
    Invoke-AppCleanup
})
$script:Window.Add_Closed({
    Invoke-AppCleanup
    $script:Frame.Continue = $false
})
$script:Window.Add_StateChanged({
    if ($script:Window.WindowState -eq [System.Windows.WindowState]::Minimized -and (Test-KeepInTray)) { Hide-ToTray }
})
$script:Window.Add_IsVisibleChanged({ try { Update-Tray } catch { } })

# Safety net: report unexpected UI exceptions instead of closing the window.
$script:Window.Dispatcher.Add_UnhandledException({
    param($sender, $e)
    $e.Handled = $true
    try { Show-AppError $e.Exception 'Unexpected error' } catch { }
})

# Pulls results from the background runspaces onto the UI thread.
$script:Timer = New-Object System.Windows.Threading.DispatcherTimer
$script:Timer.Interval = [TimeSpan]::FromMilliseconds(150)
$script:Timer.Add_Tick({ try { Invoke-TimerTick } catch { Set-Status "Background update failed: $(Get-InnerMessage $_.Exception)" 'Error' } })
$script:Timer.Start()
#endregion

$script:WindowCreated = $false
$script:Frame = New-Object System.Windows.Threading.DispatcherFrame
$startHidden = $false
if ($ApiMode) {
    try {
        Start-ApiServer
        if ($null -ne $script:Tray) {
            $startHidden = $true
            $script:TrayTipShown = $true
            $script:Tray.ShowBalloonTip(5000, $script:AppName,
                "API server running at http://localhost:$($script:Api.Port)/. Right-click the tray icon to open the web page, show the window or exit.",
                [System.Windows.Forms.ToolTipIcon]::Info)
        }
    } catch {
        Show-AppError $_ 'The API server could not be started.'
    }
}
if (-not $startHidden) { $script:Window.Show() }
Update-Tray
[System.Windows.Threading.Dispatcher]::PushFrame($script:Frame)
