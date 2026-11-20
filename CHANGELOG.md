# Changelog

All notable changes to this project are documented in this file.
The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project uses [Semantic Versioning](https://semver.org/).

## [2.0.0] - 2026-11-19

A full rewrite. The tool moves from a fixed-size WinForms window to a resizable WPF interface, and gains batch checks, history, tenant details, diagnostics, a secured local API and automated tests.

### Added

#### Results
- **Statuses with next steps.** Every result is Definitive, Inconclusive, Throttled, Temporarily unavailable or Invalid address. Anything not definitive shows a prominent callout with the recommended next step, in the window, the web page, the API and exports.
- **Inconclusive status** for addresses that are not found although their domain belongs to a Microsoft tenant (usually an alias, guest or contact), and for unknown result codes.
- **Tenant details** for work and school domains: organization name, tenant ID, tenant region and domain type. These come from GetUserRealm and the tenant's OpenID discovery document.
- **Federation provider** identification from the sign-in URL, covering GoDaddy, Okta, OneLogin, Ping, Duo, Auth0, AD FS and others.
- **Throttling protection.** HTTP 429/503 and `ThrottleStatus` are detected, with up to three attempts and back-off. Results that stay throttled are flagged and not cached.
- Copy buttons for the Cloud URL, FederationRedirectUrl, organization and tenant ID; **Check again live** for cached answers; and a raw response view.

#### Change detection and audit
- **Change detection.** An address whose account type, cloud, tenant or federation state or provider differs from its previous definitive check is flagged. It shows a purple callout on Single check, a Changes column in Batch check and History, a "Changed since previous check" filter in the app, the web page and `/api/history?status=Changed`, a Changed counter, and tray alerts for API results.
- **Domain summary** (Group by domain). One row per domain: addresses, organization, tenant ID, cloud, federation provider, account-type mix, not-found, need-attention and changed counts. Exportable as CSV or JSON.
- **Executive report** (HTML and PDF):
  - key figures and breakdowns by account type, certainty, cloud and federation provider
  - top organizations and tenants, and a daily trend
  - exceptions with next steps or changes, domains, and every result
  - prints cleanly, and PDF export uses Microsoft Edge or Google Chrome headless

#### Batch checks
- Paste a list or import `.txt`/`.csv` files, with live progress and CSV, JSON or HTML export.
- **Pause and Resume**, and **Stop**.
- **Retry failed**, which re-checks only Throttled and Temporarily unavailable results and updates their rows in place.
- **Skip cached results** option.
- Configurable delay between lookups, and up to 4 parallel lookups.
- Duplicates skipped and counted, with counters per account type and status.
- **Refresh stale only**: reuse recent answers and re-check those older than a configurable age.
- **CSV column mapping**: choose the email column and keep other columns (employee ID, name, department and so on) in the results table and exports. Includes a template.
- **Watched folder**: new CSV/TXT files are checked automatically. Results are saved to `results\`, and inputs moved to `processed\`.
- **Clipboard collector**: adds addresses you copy anywhere to the batch list.
- Keyboard control: Ctrl+P pause/resume, Ctrl+Shift+S stop, Ctrl+R retry failed, Ctrl+G group by domain.

#### History
- Statistics, search, and **filters by account type, cloud, status, federation, source and date**.
- **Export only the filtered rows**, as CSV or an HTML report.
- Saved between sessions (last 5,000 checks).

#### Diagnostics
- A Diagnostics tab showing:
  - live lookups, cache hits and cache hit rate
  - retries, throttling events, unavailable and inconclusive results
  - per-endpoint calls, failures, average and last response time, and availability
  - the last successful lookup, the last problem
  - the API server state, its request count and refused requests

#### Cache
- Cache lifetimes per outcome: definitive and inconclusive answers in hours, throttled or unavailable results in minutes (off by default).

#### Privacy
- A visible notice that email addresses and tenant details are stored only on this computer.
- **History retention** and **log retention** in days, applied at start-up and every hour.
- **Clear history and the cache on exit** option.
- Buttons to clear history, clear the cache and delete all logs.
- **Encryption at rest** for history and the cache, with Windows data protection (DPAPI) for the current user.
- **Private session**: checks made while it is on are never saved to history, the cache or logs.
- **Redacted exports**: addresses shown as `j***@contoso.com`, and imported columns blanked, in every export and report.
- **Delete exported files**: removes every file the app has exported.
- `.gitignore` entries for the `config` and `logs` folders.

#### Local REST API and web page
- REST API on `http://localhost:<port>/` with `/api/check`, `/api/batch`, `/api/history` (with filters), `/api/health`, `/api/stats` and `/api/endpoints`. Includes a request log and a built-in test request. No administrator rights are needed.
- **Web page** for single checks, batch checks and history in any browser. Its History tab has the same filters as the desktop app. Its theme follows the desktop app, with its own Match app / Light / Dark control.
- **System tray mode**, plus `-ApiMode`, `-Port` and `-NoBrowser` parameters.

#### Security
- Cross-site requests (`Origin` and `Sec-Fetch-Site`) are refused, so websites you visit cannot use the API.
- **Optional API token**, sent as `X-API-Key` or `Authorization: Bearer`. The web page includes it automatically when opened from this computer.
- Content Security Policy, `X-Frame-Options`, `nosniff` and `no-referrer` headers on the web page.
- **Per-program rate limit** (default 120 requests a minute) with HTTP 429 and `Retry-After`.

#### API operations
- `/api/openapi.json` (OpenAPI 3), `/health` (liveness, no token), and `/metrics` (Prometheus).
- **Background jobs**: `POST /api/jobs` for up to 5,000 addresses, `GET /api/jobs/{id}` for progress and results, `DELETE` to cancel.

#### Automation
- **Command-line mode**: `-Email` or `-InputFile`, with `-Format Table|Object|Json|Csv`, `-OutFile`, and exit codes 0 to 3.
- **PowerShell module** `MSAccountChecker`, with `Test-MSAccount` (pipeline input), `Clear-MSAccountCache` and `Get-MSAccountExitCode`.

#### Endpoint health and updates
- **Recent endpoint health** on the Diagnostics tab and in `/api/stats`:
  - availability, median and 95th-percentile latency, and throttling rate over 5 minutes
  - a 30-minute latency chart
  - a suggested safe batch speed with one-click Apply
- **Update checker** (optional, off at start-up by default). It reads the latest GitHub release, shows the release notes in-app, and downloads the zip and verifies it against the published `SHA256SUMS.txt` and the script's Authenticode signature. It never installs anything.
- An in-app **What's new** viewer, and signature status in About.

#### Usability and accessibility
- Keyboard shortcuts: Ctrl+1 to Ctrl+6, Ctrl+comma (Settings), Ctrl+L, Ctrl+Shift+C, Ctrl+Enter, Ctrl+P, Ctrl+Shift+S, Ctrl+R, Ctrl+G, Ctrl+F, Ctrl+E, Ctrl+D, F1 and Esc.
- A **Settings** button in the header, before About.
- Screen-reader names, help text and live status announcements in the window and on the web page.
- **Compact tables** option.
- **WCAG 2.1 contrast**: three light-theme colours adjusted so every text colour meets 4.5:1. A test checks every pair.
- Visible focus rings on buttons, switches, tabs and filters.
- Tooltips on technical fields.
- Right-click menus on result tables.
- **Windows high contrast support**, including changes while the app is open.
- Table columns that resize with the window.

#### Quality
- **Automated tests** (Pester 5). They cover:
  - email validation, every classification path, federation signals, statuses, and retry and back-off
  - cache expiry and policies, change detection, the domain summary, CSV mapping, redaction, history filters and import/export
  - endpoint-health advice, command-line exit codes and the module
  - DPAPI round-trips and theme contrast
  - the API's responses, security checks, rate limits, metrics, OpenAPI and background jobs
- A GitHub Actions workflow that runs the tests on Windows PowerShell 5.1 and PowerShell 7 for every push.
- An app icon, an About panel, and remembered window size and position.

### Changed
- **Federation detection** now uses GetUserRealm `NameSpaceType`, `FederationRedirectUrl`, `EstsProperties.DomainType` and `IsFederated`, instead of `IsFederated` alone.
- Throttled answers from GCC High and China are no longer trusted when checking sovereign clouds.
- Dark mode applies instantly to the whole window, including the title bar.
- The Settings menu is replaced by header switches and a Settings tab.
- Lookups run in the background, so the window never freezes.
- The app no longer keeps a PowerShell window open behind it. `-Console` keeps it in the console for troubleshooting.
- Email validation accepts `+` addressing and other valid characters.
- Log files have new columns: Organization, TenantId, TenantRegion, FederationProvider and Status.
- Logging errors are reported in the status bar instead of being ignored silently.

### Removed
- The Litera One compatibility assessment, from the result, the log and the documentation.

### Unchanged
- The lookup order: Commercial first, then GCC High and China only when Commercial returns code 1 for a non-personal domain.
- The account-type wording and the four-line result format.
- The settings file (`script/config/AccChecker.config`, Key=Value). It is the only settings file.
- The log file location and name (`script/logs/MS_Account_Checker_Logs_YYYY-MM-DD.txt`).

## [1.1.0]

- WinForms GUI to check whether an email address is a Microsoft account across the Commercial, GCC High and China clouds.
- Colour-coded results, dark mode and optional daily TSV logging with CSV export.
- `launcher.vbs` to start the GUI without a console window.

[2.0.0]: https://github.com/imumeshk/MS-Account-Checker/releases/tag/v2.0.0
