# Microsoft Account Checker

![Version](https://img.shields.io/badge/version-2.0.0-3659C9)
![PowerShell](https://img.shields.io/badge/PowerShell-5.1%20%7C%207%2B-243E9E)
![Platform](https://img.shields.io/badge/platform-Windows-64748B)
![License](https://img.shields.io/badge/license-CC0--1.0-15803D)
[![Tests](https://github.com/imumeshk/MS-Account-Checker/actions/workflows/tests.yml/badge.svg)](https://github.com/imumeshk/MS-Account-Checker/actions/workflows/tests.yml)

A Windows desktop tool that tells you what kind of Microsoft account sits behind an email address, and in which Microsoft cloud it lives: Commercial, GCC High or China.

For work and school domains it also shows the organization name, tenant ID, tenant region and federation provider. Every result says how certain it is, and what to do next when Microsoft's answer is not definitive. It also flags addresses whose answer changed since their previous check.

You can check addresses in several ways:
- one at a time, or in batches, in the window
- from a watched folder or the clipboard
- from the command line, or through a PowerShell module
- through a local REST API and web page

> Version 2.0.0 is a full rewrite of the 1.x WinForms tool. See [CHANGELOG.md](CHANGELOG.md) for everything that changed.

---

## Contents

- [Features](#features)
- [Requirements](#requirements)
- [Getting started](#getting-started)
- [Using the app](#using-the-app)
- [Results and statuses](#results-and-statuses)
- [How a lookup works](#how-a-lookup-works)
- [Command line and PowerShell module](#command-line-and-powershell-module)
- [Web page, system tray and API mode](#web-page-system-tray-and-api-mode)
- [Local REST API](#local-rest-api)
- [Privacy](#privacy)
- [Files and settings](#files-and-settings)
- [Keyboard shortcuts and accessibility](#keyboard-shortcuts-and-accessibility)
- [Updates](#updates)
- [Tests](#tests)
- [Troubleshooting](#troubleshooting)
- [Upgrading from 1.x](#upgrading-from-1x)
- [License](#license)

---

## Features

- **Single check** with clickable sign-in URLs, copy buttons for the URLs and tenant ID, and the raw response from Microsoft.
- **Clear certainty.** Every result is marked **Definitive**, **Inconclusive**, **Throttled**, **Temporarily unavailable** or **Invalid address**. Anything that is not definitive gets a recommended next step.
- **Tenant details.** Organization name, tenant ID, tenant region and federation provider (GoDaddy, Okta, Ping, AD FS and others).
- **Reliable federation detection.** Confirmed with GetUserRealm, the redirect URL and the domain type, not one response field.
- **Throttling protection.** Microsoft's rate limits are detected and retried with back-off. Results that stay throttled are flagged and never cached.
- **Batch checks.**
  - Paste a list or import `.txt` or `.csv` files.
  - Live progress, with Pause, Resume and Stop.
  - **Retry failed** re-checks only throttled and unavailable results.
  - Optionally skip cached results.
  - Choose the delay between lookups and run up to 4 lookups in parallel.
  - Export to CSV, JSON or an HTML report.
- **Change detection.** An address whose account type, cloud, tenant or federation changed since its previous check is flagged in the result, the tables, the reports, the API and the tray.
- **Domain summary.** Group batch results by domain to see each domain's organization, tenant, cloud, federation provider and account-type mix. Useful for migration and audit work.
- **CSV column mapping.** Choose the email column and keep columns such as employee ID, name or department alongside the results and in exports. A template is included.
- **Automatic inputs.** A watched folder checks new CSV and TXT files and saves the results beside them. A clipboard collector adds every address you copy to the batch list.
- **Command line and PowerShell module.** Structured output (table, objects, JSON or CSV) and exit codes for scripts and CI.
- **Executive report.** HTML or PDF with totals, certainty, account-type, cloud, provider and tenant breakdowns, a daily trend, exceptions and domains.
- **History.** Statistics; search; filters by account type, cloud, status (including "changed since previous check"), federation, source and date; and export of just the filtered rows.
- **Diagnostics and endpoint health.**
  - Availability, median and 95th-percentile response time, and throttling rate over the last 5 minutes.
  - A 30-minute response-time chart.
  - A suggested safe batch speed that you can apply with one click.
  - Per-endpoint counters, the last success and the last problem.
- **Configurable cache.** Separate lifetimes for definitive, inconclusive and failed answers, plus a "refresh stale only" batch option.
- **Privacy controls.**
  - A storage notice, plus history and log retention periods.
  - Encryption of history and the cache with Windows data protection.
  - A private session that saves nothing, and clear-on-exit.
  - Redacted exports, and one-click deletion of history, cache, logs and exported files.
- **Local REST API and web page.**
  - Cross-site requests are refused, with an optional API token and per-program rate limits.
  - OpenAPI documentation, `/health` and Prometheus `/metrics`.
  - Background jobs for large lists.
  - Can run from the system tray.
- **Update checker.** Optional. It reads the latest GitHub release, shows the release notes in the app, and can download the zip and verify it against the published SHA-256 checksum and the script's Authenticode signature. It never installs anything itself.
- **Accessibility.**
  - Light, dark and Windows high-contrast themes. Every colour pair meets the WCAG 4.5:1 contrast minimum, and a test checks this.
  - Keyboard-first batch control.
  - Screen-reader names and live status announcements.
  - Visible focus rings, tooltips on technical fields, and a compact table view.
- **Automated tests** with Pester, run on every push by GitHub Actions.

All lookups run in the background, so the window stays responsive.

## Requirements

- Windows 10 or Windows 11
- Windows PowerShell 5.1, or PowerShell 7 or later
- Internet access to `login.microsoftonline.com`, `login.microsoftonline.us` and `login.partner.microsoftonline.cn`

No modules need to be installed to run the app. Pester 5 is needed only to run the tests.

## Getting started

1. Download the latest release from the [Releases page](https://github.com/imumeshk/MS-Account-Checker/releases) and extract it.
2. Double-click **`launcher.vbs`**.

To start it from PowerShell instead:

```powershell
powershell.exe -STA -ExecutionPolicy Bypass -File .\script\MS_Account_Checker.ps1
```

The app never keeps a PowerShell window open behind it. When started from a visible console, it restarts itself in the background, and the console returns or closes straight away. Add `-Console` to keep it in the console and see error messages there.

If Windows blocks the files after download, right-click the zip, choose **Properties**, tick **Unblock**, then extract again. Alternatively, run:

```powershell
Get-ChildItem -Recurse | Unblock-File
```

## Using the app

The window has six tabs. Ctrl+1 to Ctrl+6 switch between them. The **Settings** button in the header, before About (or Ctrl+,), opens the Settings tab.

### Single check

Enter an address and press Enter, or select **Check**. The result shows:

- The four-line result, which **Copy result** copies:

  ```
  Username: user@contoso.com
  FoundInCloud: Commercial
  Cloud URL: https://login.microsoftonline.com
  Account Type: Likely Entra Microsoft Account (Work/School).
  ```

  Federated accounts get an extra `FederationRedirectUrl` line.
- An **account-type badge** and a **status badge**.
- A **callout with the recommended next step** whenever the answer is not definitive.
- **Tenant details** underneath, with copy buttons for the URLs, organization and tenant ID. Hover over a field name for an explanation.
- **Check again live**, which appears when the answer came from the cache.
- **A purple "changed" callout** when the account type, cloud, tenant or federation differs from the previous check of the same address.
- **Show raw response**, which shows exactly what Microsoft returned.

### Batch check

Paste addresses into the list, one per line. Commas, semicolons and spaces also separate addresses. You can also select **Import file**:

- For a `.csv` file, a dialog shows its columns. Choose which one holds the addresses; the app suggests `Email`, `Mail`, `UserPrincipalName`, `UPN` or `Username`, or the column that contains addresses. Then tick the other columns to keep, such as employee ID, name or department. Kept columns appear in the results table and in CSV, JSON and watched-folder exports. **Save a template** writes an example CSV.
- For any other file, it picks up every address it finds.

Duplicates are skipped and counted, and invalid entries are listed as errors. Turn on **Collect from clipboard** to add every address you copy anywhere to the list.

| Control | What it does |
| --- | --- |
| **Start** (Ctrl+Enter) | Check every address in the list |
| **Pause / Resume** | Pause after the current lookup, then carry on |
| **Stop** | Stop after the current lookup |
| **Retry failed** | Check again only the Throttled and Temporarily unavailable results, updating their rows in place |
| **Skip cached results** | Ask Microsoft for every address, even ones in the cache |
| **Refresh stale only** | Reuse recent answers, but ask again for any older than the **Stale after** setting (default 12 hours) |
| **Group by domain** (Ctrl+G) | One row per domain: addresses, organization, tenant ID, cloud, federation provider, account-type mix, not-found, need-attention and changed counts |
| **Export CSV / JSON** | Save all results, or the domain summary when grouped |
| **HTML / PDF report** | The executive report (see below). PDF uses Microsoft Edge or Google Chrome in the background. |

The coloured counters show how many of each account type and status there are. Right-click rows to copy usernames, tenant IDs or whole rows, or to check an address again.

The delay between lookups (default 200 ms) and the number of parallel lookups (default 1, at most 4) are set on the Settings tab. More parallel lookups finish sooner but make Microsoft throttling more likely. The Diagnostics tab suggests safe values from recent throttling.

Keyboard: Ctrl+Enter starts, Ctrl+P pauses or resumes, Ctrl+Shift+S stops, Ctrl+R retries failed lookups, Ctrl+E exports.

#### Executive report

The HTML and PDF reports contain:

- key figures: checks, unique addresses, the share of definitive answers, items needing attention, changes, distinct tenants and the share federated
- breakdowns by account type, certainty, cloud and federation provider
- the top organizations and tenants
- a checks-per-day trend
- an exceptions list with each item's next step or change
- a domain table
- every result

The HTML version has a **Print or save as PDF** button and prints cleanly.

### History

History lists every check, newest first, from the window, batches, the API and the web page. Statistic cards at the top count each account type.

- **Filter** by account type, cloud, status, federation, source (live or cache) and date (today, or the last 7, 30 or 90 days).
- **Search** by username, account type, cloud, organization or origin (Ctrl+F).
- **Export CSV**, **HTML report** and **PDF report** save only the rows currently shown (Ctrl+E). Choose **Changed since previous check** in the status filter to see every address whose answer changed.
- **Open log** and **Export log** open today's log file or save it as CSV.

### API server

Shows the port, the address, an endpoint list with ready-to-run examples, a request log and a **Send** button for test requests. See [Local REST API](#local-rest-api).

### Diagnostics

- **Counters for the current session:**
  - live lookups, cache hits and cache hit rate
  - retries, throttling events, unavailable and inconclusive results
  - API requests and refused requests
- **Microsoft endpoint health, last 5 minutes:**
  - availability
  - median and 95th-percentile lookup response time
  - throttling rate
  - a **suggested safe batch speed**. **Apply to batch settings** copies it to the Settings tab.
- **Lookup response time, last 30 minutes:** a per-minute chart. Red dots mark failed calls.
- **Microsoft endpoints:** calls, failures, average and last response time, and availability per endpoint for the whole session.
- **Status:** the last successful lookup, the last problem, the API server state (including background jobs) and refused requests.

**Reset counters** starts again.

### Settings

- **Privacy and stored data.**
  - the storage notice and the data folder
  - encryption, private session and redacted exports
  - history and log retention, and clear-on-exit
  - buttons to clear history, clear the cache, delete all logs, or delete every file the app has exported
- **Cache.** How long to reuse definitive answers (hours), inconclusive answers (hours), and throttled or unavailable results (minutes, default 0). Also when answers count as stale.
- **Batch lookups.** The delay between lookups and the number of parallel lookups.
- **Watched folder.** Choose a folder and turn on automatic processing (see [Automatic inputs](#automatic-inputs)).
- **API security and limits.** The token, and the rate limit per program.
- **Updates.** Check at start-up, check now, download and verify, and **What's new**.
- **Display and accessibility.** Compact tables.
- **Keyboard shortcuts.**

#### Automatic inputs

When **Process new files in this folder automatically** is on, the app checks the watched folder every 5 seconds:

1. A new `.csv` or `.txt` file is moved to `processing\`.
2. It is checked like a batch; for CSV files, every column is kept.
3. The results are saved to `results\<name>_results_<time>.csv`.
4. The input file is moved to `processed\`.

Results also go to History, marked with the origin *Folder*. Each file replaces the Batch check list while it runs, and files wait while another batch is running. Programs can drop files into the folder to submit work. For an HTTP interface, see the [background jobs](#background-jobs) API.

### Header switches

| Switch | What it does |
| --- | --- |
| Dark mode | Switches the whole window between light and dark immediately (Ctrl+D). Turned off while Windows high contrast is on. |
| API server | Starts or stops the local REST API and web page. |
| Logging | Writes each check to the daily log file. |

## Results and statuses

### Account types

| Account type | Meaning |
| --- | --- |
| Likely Entra Microsoft Account (Work/School). | Managed work or school account |
| Federated Entra ID (Work/School) Account. | Work or school account that signs in through another identity provider |
| Likely Personal Microsoft Account (MSA). | Personal account on outlook.com, hotmail.com, live.com or msn.com |
| Dual Identity (Work/School & Personal) Account. | Both a personal and a work or school account use this address |
| Likely a GCC High Account. | Found in the US Government (GCC High) cloud |
| Likely an M365 China (VNET) Account. | Found in the Microsoft 365 China cloud |
| Not a Microsoft Account. | No Microsoft account found in any cloud |
| Unknown Result Code | Microsoft returned a code this tool does not recognise |

### Statuses

Microsoft's discovery endpoints are not always definitive, so every result also carries a status.

| Status | Meaning | Recommended next step |
| --- | --- | --- |
| **Definitive** | Microsoft gave a clear answer. | None. |
| **Inconclusive** | The address was not found, but its domain belongs to a Microsoft tenant; or Microsoft returned an unknown code. | Confirm the user's primary sign-in address (UPN) with the organization's administrator. The address may be an alias, a guest or a mail contact. |
| **Throttled** | Microsoft rate-limited every attempt, so the answer may be wrong. | Wait a few minutes and check again, or use **Retry failed**. |
| **Temporarily unavailable** | Microsoft could not be reached. | Check the network connection, proxy or firewall, then check again. |
| **Invalid address** | The text is not a valid email address. | Correct it and check again. |

Throttled, unavailable and invalid results are never cached, so the next check always asks Microsoft again.

## How a lookup works

1. **GetCredentialType, Commercial cloud.** Always called first.
2. **GetCredentialType, GCC High then China.** Called only when Commercial returns `IfExistsResult` 1 and the domain is not a personal domain. A sovereign result of 0 or 5 is accepted. Throttled sovereign answers are ignored, because a throttled reply can claim an account exists.
3. **GetUserRealm** (Commercial only). Confirms whether the domain is managed or federated, and returns the organization name.
4. **OpenID discovery document** for the domain. Returns the tenant ID and tenant region.

Steps 3 and 4 run once per domain, and their answers are cached for 24 hours.

| Purpose | Endpoint |
| --- | --- |
| Credential type (Commercial) | `https://login.microsoftonline.com/common/GetCredentialType` |
| Credential type (GCC High) | `https://login.microsoftonline.us/common/GetCredentialType` |
| Credential type (China) | `https://login.partner.microsoftonline.cn/common/GetCredentialType` |
| User realm | `https://login.microsoftonline.com/getuserrealm.srf?login=<email>&json=1` |
| Tenant discovery | `https://<cloud login host>/<domain>/.well-known/openid-configuration` |

`IfExistsResult` codes:

| Code | Meaning |
| --- | --- |
| 0 | Account exists |
| 1 | Account not found |
| 5 | Account exists in another identity provider (federated) |
| 6 | Dual identity: both personal and work/school |

An account counts as **federated** when any one of these is true:

- GetUserRealm reports `Federated`.
- The response contains a `FederationRedirectUrl`.
- `EstsProperties.DomainType` is 4.
- `IsFederated` is true.

**Throttling** is detected through HTTP 429 or 503, or a `ThrottleStatus` other than 0. Each call is tried up to three times, with pauses of about 2 and then 4 seconds.

## Command line and PowerShell module

### Command line

Pass `-Email` or `-InputFile` and the script runs without a window. It prints the results and sets an exit code.

```powershell
.\script\MS_Account_Checker.ps1 -Email user@contoso.com, other@fabrikam.com
.\script\MS_Account_Checker.ps1 -InputFile .\users.csv -OutFile .\results.csv -Format Csv
.\script\MS_Account_Checker.ps1 -Email user@contoso.com -Format Json
$results = .\script\MS_Account_Checker.ps1 -InputFile .\users.csv -Format Object
```

| Parameter | Description |
| --- | --- |
| `-Email <addresses>` | Addresses to check (commas, semicolons and spaces also separate them) |
| `-InputFile <path>` | A `.csv` or `.txt` file. CSV columns other than the email column are kept in the output. |
| `-Format` | `Table` (default), `Object` (for the pipeline), `Json` or `Csv` |
| `-OutFile <path>` | Write CSV or JSON to a file |
| `-NoCache` | Do not reuse answers within the run |

| Exit code | Meaning |
| --- | --- |
| 0 | Every answer was definitive |
| 1 | At least one result was Inconclusive, Throttled or Invalid |
| 2 | At least one lookup was Temporarily unavailable |
| 3 | Bad input: no addresses, or the input file was not found |

### PowerShell module

```powershell
Import-Module .\module\MSAccountChecker.psd1
Test-MSAccount -Email user@contoso.com
Import-Csv .\users.csv | Test-MSAccount | Export-Csv .\results.csv -NoTypeInformation
$results = Get-Content .\list.txt | Test-MSAccount
exit (Get-MSAccountExitCode -Results $results)
```

`Test-MSAccount` accepts strings from the pipeline, or objects with an `Email`, `Mail`, `UserPrincipalName`, `UPN` or `EmailAddress` property. It returns the same fields as the REST API. `-NoCache` asks Microsoft every time, and `-DelayMs` sets the pause between live lookups (default 200).

`Clear-MSAccountCache` forgets cached answers. `Get-MSAccountExitCode` turns results into the exit codes above.

The module reads its lookup rules from `script\MS_Account_Checker.ps1`, so keep the `module` and `script` folders side by side.

## Web page, system tray and API mode

### Web page

When the API server is running, `http://localhost:8080/` (or your own port) is a page for checking accounts in any browser. It opens automatically when the server starts. To stop that, turn off **Open the web page when the API server starts** on the API server tab.

The page has three tabs:

- **Single check** shows the result, the account-type and status badges, the next step and the tenant details.
- **Batch check** checks a pasted list, with progress, Stop and CSV export.
- **History** shows the desktop app's history with statistics, search and the same filters as the desktop app: account type, cloud, status, federation, source and date. **Export CSV** saves the rows shown. History can only be cleared from the desktop app.

The page's theme control has three settings:

- **Match app** (the default) follows the desktop app's Dark mode switch.
- **Light** and **Dark** override it for that browser.

### System tray

While the API server is running, closing or minimizing the window hides the app to the system tray, and the API and web page keep working. You can turn this off on the API server tab.

- **Left-click** the tray icon to show the window.
- **Right-click** the tray icon for these options:
  - Open web page
  - Show window
  - Start or stop the API server
  - Exit

### API mode

```powershell
powershell.exe -STA -WindowStyle Hidden -ExecutionPolicy Bypass -File .\script\MS_Account_Checker.ps1 -ApiMode -Port 8080
```

This starts the API server, opens the web page and runs from the tray without showing the window.

| Parameter | Description |
| --- | --- |
| `-ApiMode` | Start the API server and run from the system tray |
| `-Port <number>` | API port, overriding the saved setting |
| `-NoBrowser` | Do not open the web page when the server starts |
| `-Console` | Stay attached to the console it was started from (for troubleshooting) |

## Local REST API

Turn on the **API server** switch. The server only accepts connections from the same computer, and needs no administrator rights. API results appear in History and share the cache.

| Method | Path | Description |
| --- | --- | --- |
| GET | `/` | Web page |
| GET | `/api/check?email=user@domain.com` | Check one address |
| POST | `/api/batch` | Check up to 200 addresses. Body: `["a@x.com","b@y.com"]` or `{"emails":[...]}` |
| GET | `/api/history?limit=200&search=text` | Recent checks from History, newest first, with counts. Optional filters below |
| POST | `/api/jobs` | Submit up to 5,000 addresses as a background job. Returns 202 with the job ID |
| GET | `/api/jobs` | List background jobs |
| GET | `/api/jobs/{id}` | A job's status, progress and results |
| DELETE | `/api/jobs/{id}` | Cancel a job |
| GET | `/api/health` | Server status, uptime, theme and whether a token is required |
| GET | `/api/stats` | Request counters and lookup diagnostics, including recent endpoint health |
| GET | `/api/openapi.json` | OpenAPI 3 description of the API |
| GET | `/api/endpoints` | List of endpoints |
| GET | `/health` | Liveness check. Never needs a token and is never rate-limited |
| GET | `/metrics` | Prometheus metrics: requests, refusals, live and cached lookups, retries, throttling, per-endpoint calls, failures and latency, and recent availability |

### History filters

`/api/history` accepts the same filters as the History tab. Combine any of them with `search` and `limit` (1 to 5,000, default 200):

| Parameter | Values |
| --- | --- |
| `type` | `Entra`, `MSA`, `Federated`, `Dual`, `Sovereign` (GCC High or China), `NotMicrosoft`, `Error` |
| `cloud` | `Commercial`, `GCCHigh`, `China`, `N/A` |
| `status` | `OK`, `Inconclusive`, `Throttled`, `Unavailable`, `Invalid`, or `Changed` for addresses that changed since their previous check |
| `federation` | `yes`, `no` |
| `source` | `Live`, `Cache` |
| `days` | `0` for today, or a number of days back |

For example, `/api/history?status=Throttled&days=7` returns this week's throttled checks. The `counts` in the response always cover the whole history.

### Background jobs

For large lists, submit a job and poll it, instead of holding one long `/api/batch` request:

```powershell
$job = Invoke-RestMethod http://localhost:8080/api/jobs -Method Post -ContentType 'application/json' -Body (ConvertTo-Json @(Get-Content .\list.txt))
do { Start-Sleep 5; $status = Invoke-RestMethod "http://localhost:8080$($job.statusUrl)" } while ($status.status -in 'queued', 'running')
$status.results | Export-Csv .\results.csv -NoTypeInformation
```

Jobs run one at a time in the background, using the batch delay from the Settings tab. Their results also appear in History. The newest 100 jobs are kept until the app closes.

### Security

- **Requests from other websites are refused.** Any request that a browser marks as coming from another site (`Origin` or `Sec-Fetch-Site`) gets HTTP 403. A web page you happen to visit therefore cannot send lookups to your API or read its answers.
  - No CORS headers are sent.
  - Following a link to the web page from another site still opens it, because that site cannot read the page.
- **Optional token.** On the Settings tab, turn on **Require an API token**. Every `/api` request then needs the token, and gets HTTP 401 without it. The web page includes the token automatically when you open it from this computer. Scripts send it like this:

  ```powershell
  Invoke-RestMethod "http://localhost:8080/api/check?email=user@contoso.com" -Headers @{ 'X-API-Key' = '<token>' }
  ```

  `Authorization: Bearer <token>` also works.
- **Rate limit.** Each program may send up to 120 requests a minute by default (Settings tab; 0 turns it off). Beyond that it gets HTTP 429 with a `Retry-After` header. The web page's own refreshes and `/health` are exempt.
- **Protective headers.** The web page is served with a Content Security Policy, `X-Frame-Options: DENY`, `X-Content-Type-Options: nosniff` and `Referrer-Policy: no-referrer`.
- **No IP allow-list.** The server only accepts connections from this computer, where every client's address is `127.0.0.1` or `::1`, so an allow-list could not tell programs apart. Use the token for that instead.

### Example response

```json
{
  "username": "user@contoso.com",
  "foundInCloud": "Commercial",
  "cloudUrl": "https://login.microsoftonline.com",
  "federationRedirectUrl": "",
  "accountType": "Likely Entra Microsoft Account (Work/School).",
  "category": "Entra",
  "organization": "Contoso",
  "tenantId": "00000000-0000-0000-0000-000000000000",
  "tenantRegion": "North America (NA)",
  "domainType": "Managed",
  "federationProvider": "",
  "status": "OK",
  "statusLabel": "Definitive",
  "nextAction": "",
  "throttled": false,
  "note": "",
  "source": "Live",
  "checkedAt": "2026-11-17 10:30:45",
  "output": "Username: user@contoso.com\r\nFoundInCloud: Commercial\r\n...",
  "error": ""
}
```

HTTP status codes:

- **400** for a missing or invalid address
- **401** for a missing or wrong token
- **429** when the rate limit is reached
- **403** for a cross-site request
- **413** for more than 200 addresses
- **502** when Microsoft could not be reached

URL-encode the address in the query string, so a `+` becomes `%2B`.

## Privacy

- To check an address, the app sends it to Microsoft's public sign-in endpoints listed above. Nothing is sent anywhere else.
- **Email addresses, account types and tenant details (organization, tenant ID and region) are stored only on this computer**, in `script/config` (history and cache) and `script/logs` (logs, when logging is on). The app shows this on the History and Settings tabs.
- On the Settings tab you can:
  - keep history for a set number of days,
  - delete log files older than a set number of days,
  - clear history and the cache every time the app closes,
  - clear history, clear the cache or delete all logs at any time.
  
  Retention is applied at start-up and every hour.
- **Encrypt history and the cache on disk** uses Windows data protection (DPAPI) for your account. Only your Windows account on this computer can read the files. Logs stay plain text, so turn logging off if that matters.
- **Private session** keeps new checks in memory only. They are never written to history, the cache or logs, and the setting resets when the app restarts.
- **Redact email addresses and imported columns in exports** writes `j***@contoso.com` instead of the full address, and blanks imported columns. This applies to CSV, JSON, HTML, PDF and watched-folder results.
- **Delete exported files** removes every file the app has exported. They are listed in `script/config/exported_files.json`.
- These folders contain email addresses. Do not commit them to source control; the included `.gitignore` excludes them.

## Files and settings

```
MS-Account-Checker/
|-- launcher.vbs                       Starts the app without a console window
|-- script/
|   |-- MS_Account_Checker.ps1         The application
|   |-- config/
|   |   |-- AccChecker.config          Settings (Key=Value)
|   |   |-- check_history.json         History (last 5,000 checks)
|   |   |-- lookup_cache.json          Lookup cache
|   |   `-- exported_files.json        Files the app has exported (for one-click deletion)
|   `-- logs/
|       `-- MS_Account_Checker_Logs_YYYY-MM-DD.txt
|-- module/MSAccountChecker.psd1       PowerShell module (Test-MSAccount)
|-- module/MSAccountChecker.psm1
|-- tests/MS_Account_Checker.Tests.ps1 Pester tests
|-- .github/workflows/tests.yml        Runs the tests on every push
|-- README.md
|-- CHANGELOG.md
`-- LICENSE
```

**`AccChecker.config` is the only settings file.** Files named `config.json` or `settings.json` in the config folder come from pre-release builds. The app ignores them, and you can delete them.

### Settings (`script/config/AccChecker.config`)

All of these can be changed in the app.

| Key | Default | Description |
| --- | --- | --- |
| `EnableLogging` | `False` | Write checks to the daily log |
| `DarkMode` | `False` | Dark theme |
| `ApiPort` | `8080` | Port for the local API |
| `OpenBrowserOnApiStart` | `True` | Open the web page when the API server starts |
| `TrayOnClose` | `True` | Keep running in the system tray while the API server runs |
| `ApiRequireToken` | `False` | Require an API token |
| `ApiToken` | (generated) | The token, created when first required |
| `HistoryRetentionDays` | `0` | Delete history older than this many days (0 keeps everything, up to 5,000 checks) |
| `LogRetentionDays` | `0` | Delete log files older than this many days (0 keeps all) |
| `ClearOnExit` | `False` | Clear history and the cache when the app closes |
| `BatchDelayMs` | `200` | Pause between batch lookups, 0 to 10,000 ms |
| `BatchConcurrency` | `1` | Parallel batch lookups, 1 to 4 |
| `BatchSkipCache` | `False` | Skip cached results in batches |
| `StableCacheHours` | `24` | Reuse definitive answers for this many hours (0 to 720) |
| `InconclusiveCacheHours` | `24` | Reuse inconclusive answers for this many hours (0 to 720) |
| `FailureCacheMinutes` | `0` | Reuse throttled or unavailable results for this many minutes (0 to 60) |
| `StaleAfterHours` | `12` | Answers older than this are refreshed by **Refresh stale only** |
| `ApiRateLimitPerMinute` | `120` | Requests a minute per program (0 for no limit) |
| `EncryptStoredData` | `False` | Encrypt history and the cache with Windows data protection |
| `RedactExports` | `False` | Redact addresses and imported columns in exports |
| `WatchFolder`, `WatchFolderEnabled` | (none), `False` | The watched folder and whether it is processed |
| `CheckUpdatesOnStart` | `False` | Check GitHub for a new release at start-up |
| `CompactTables` | `False` | Smaller table rows |
| `WindowLeft`, `WindowTop`, `WindowWidth`, `WindowHeight`, `WindowMaximized` | (auto) | Last window position |

### Log format

The log is a tab-separated file, one per day, with these columns:

```
Timestamp  Username  FoundInCloud  CloudURL  FederationRedirectUrl  AccountType
Organization  TenantId  TenantRegion  FederationProvider  Status
```

## Keyboard shortcuts and accessibility

| Shortcut | Action |
| --- | --- |
| Ctrl+1 ... Ctrl+6 | Switch tabs |
| Ctrl+, | Settings |
| Ctrl+L | Go to the email box |
| Enter | Check the address |
| Ctrl+Shift+C | Copy the result |
| Ctrl+Enter | Start the batch (Batch check tab) |
| Ctrl+P | Pause or resume the batch |
| Ctrl+Shift+S | Stop the batch |
| Ctrl+R | Retry failed lookups |
| Ctrl+G | Group batch results by domain |
| Ctrl+F | Search history |
| Ctrl+E | Export history (filtered) or batch results |
| Ctrl+D | Dark mode on or off |
| F1 | About |
| Esc | Close About, What's new or the import dialog |

The app also supports accessibility in these ways:

- Every button, switch, tab and filter shows a focus ring when reached with the keyboard.
- Controls carry screen-reader names and help text, and status messages are announced as they change, in the window and on the web page.
- Every text colour in the light and dark themes meets the WCAG 2.1 contrast minimum of 4.5:1, and a test checks this.
- **Compact tables** on the Settings tab fits more results on screen.
- Technical fields and buttons have tooltips.
- Windows high contrast themes are followed automatically, including when switched while the app is open.
- Table columns resize with the window and scroll sideways when space runs out.

## Updates

On the Settings tab, **Check for updates** asks GitHub for the latest release. You can also have it check at start-up; it is off by default, and nothing else is sent.

When a newer version exists, **Download and verify**:

1. saves the release zip to your Downloads folder,
2. compares it with the release's published `SHA256SUMS.txt` checksum, and deletes it if they differ,
3. reports the Authenticode signature of the script inside.

It never installs anything; extract the zip and replace your files to update.

**What's new** shows the release notes and the bundled changelog in the app. The About panel shows whether the running script is signed.

## Tests

The Pester tests read the lookup logic, the API server and the helper functions straight from the script. Microsoft is simulated, so the tests never use the network and the app window never opens. They cover:

- email validation
- every account type and cloud
- federation signals
- Inconclusive, Throttled and Unavailable results
- retry and back-off
- cache expiry and skip-cache
- batch list parsing
- history filters
- CSV round-trips and HTML encoding
- the API's responses, cross-site refusal and token checks
- cache lifetimes by outcome and refresh-stale-only
- change detection and the domain summary
- CSV column mapping, extra columns and redaction in exports and reports
- endpoint-health figures and the safe-batch advice
- `/health`, `/metrics`, the OpenAPI document, rate limiting and background jobs
- command-line exit codes and the PowerShell module
- Windows data protection round-trips
- WCAG contrast of every theme colour

```powershell
Install-Module Pester -MinimumVersion 5.5.0 -Scope CurrentUser -Force -SkipPublisherCheck
Invoke-Pester -Path .\tests -Output Detailed
```

GitHub Actions runs the same tests on Windows PowerShell 5.1 and PowerShell 7 for every push and pull request.

## Troubleshooting

| Problem | What to do |
| --- | --- |
| Nothing happens when launching | Unblock the downloaded files (see [Getting started](#getting-started)). Then run the PowerShell command shown there with `-Console` added, to see any error. |
| "needs an STA thread" | Start it with `launcher.vbs`, or add `-STA` to the PowerShell command. |
| Many results are Throttled | Microsoft is rate-limiting your network. Wait a few minutes, then select **Retry failed**. For large batches, use 1 parallel lookup and a longer delay on the Settings tab. |
| Results are Temporarily unavailable | Microsoft could not be reached. Check the proxy or firewall; the Diagnostics tab shows the last problem and which endpoint is failing. |
| Result is Inconclusive | Microsoft's answer is not definitive. Follow the recommended next step shown with the result. |
| API server will not start | The port is in use by another program. Choose a different port on the API server tab. |
| Scripts get HTTP 401 | A token is required. Copy it from the Settings tab and send it in the `X-API-Key` header. |
| Web page or scripts get HTTP 403 | The request came from another website. Open the page directly at `http://localhost:<port>/`, or call the API from a script. |
| The window disappeared but the app is still running | The API server is on, so the app moved to the system tray. Left-click the tray icon to show it, or right-click and choose **Exit**. The icon may be in the taskbar's overflow area (^). |
| Settings not saved | Make sure the `script` folder is writable. Do not run the app from inside a zip file. |

## Upgrading from 1.x

- Replace the repository files with the 2.0.0 release. `launcher.vbs` still starts the app the same way.
- `AccChecker.config` is read as before, so your `EnableLogging` and `DarkMode` choices carry over.
- The log now has new columns (see [Log format](#log-format)). Keep any 1.x log files separately if you still need them.
- The Settings menu has been replaced by the header switches and the Settings tab.

## License

Released under the [CC0 1.0 Universal](LICENSE) public-domain dedication.

This project is not affiliated with or endorsed by Microsoft. It uses Microsoft's public sign-in discovery endpoints, which may change without notice.
