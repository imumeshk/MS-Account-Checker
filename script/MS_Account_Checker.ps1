#region --- Script Parameters ---
#region --- Assembly Loading & Initial Setup ---
# Load required .NET assemblies for GUI creation.
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Define a custom WinForms renderer class to override default menu colors for dark mode.
# This is defined once at script start to prevent type re-definition errors.
Add-Type -ReferencedAssemblies 'System.Windows.Forms', 'System.Drawing', 'System.Drawing.Primitives' -TypeDefinition @"
    public class ProfessionalColorTable : System.Windows.Forms.ProfessionalColorTable {
        public override System.Drawing.Color ToolStripDropDownBackground { get { return System.Drawing.Color.FromArgb(30,30,30); } }
        public override System.Drawing.Color MenuItemSelected { get { return System.Drawing.Color.FromArgb(70,70,70); } }
        public override System.Drawing.Color MenuStripGradientBegin { get { return System.Drawing.Color.FromArgb(30,30,30); } }
        public override System.Drawing.Color MenuStripGradientEnd { get { return System.Drawing.Color.FromArgb(30,30,30); } }
    }
"@

# Determine the application's execution directory. This method is reliable for both .ps1 scripts
# and compiled .exe files, where $MyInvocation can be null.
if ($PSScriptRoot) {
    # $PSScriptRoot is reliable for .ps1 files, whether run directly or dot-sourced.
    $scriptDir = $PSScriptRoot
} else {
    # Fallback for compiled .exe files where $PSScriptRoot is not available.
    $scriptDir = Split-Path -Parent ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName)
}
# Define paths for required directories.
$configDir  = Join-Path $scriptDir "config"
$logsDir    = Join-Path $scriptDir "logs"

# --- Application State & Configuration Variables ---
# These variables hold the application's state and are modified by user settings.
$global:EnableLogging = $false
$global:DarkMode = $false

# Define the full path to the configuration file.
$global:ConfigFile = Join-Path $configDir "AccChecker.config"

# Define a dynamic name for the log file, creating a new one each day.
$global:AccountCheckerLogFile = Join-Path $logsDir "MS_Account_Checker_Logs_$(Get-Date -Format 'yyyy-MM-dd').txt"

# Pre-compile the email validation regex for improved performance on repeated checks.
$global:EmailRegex = [regex]'^[\w.-]+@[\w.-]+\.[a-zA-Z]{2,}$'

#endregion

#region --- Core Functions: Config, Logging & Validation ---

<#
    .SYNOPSIS
        Loads application settings from the configuration file.
    .DESCRIPTION
        Reads the 'AccChecker.config' file, which uses a "Key=Value" format. It parses the
        settings for logging and dark mode, converting them to boolean values. If the file
        doesn't exist or an error occurs, it uses default settings and shows a warning.
    .NOTES
        This function is called once at script startup.
        It uses ConvertFrom-StringData for simple key-value parsing.
#>
function Load-Config {
    if (Test-Path $global:ConfigFile) {
        try {
            $config = Get-Content $global:ConfigFile | ConvertFrom-StringData
            if ($config.ContainsKey('EnableLogging')) {
                $global:EnableLogging = [System.Convert]::ToBoolean($config.EnableLogging)
            }
            if ($config.ContainsKey('DarkMode')) {
                $global:DarkMode = [System.Convert]::ToBoolean($config.DarkMode)
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to read config file at '$($global:ConfigFile)'. Using default settings.", "Config Error", "OK", "Warning")
        }
    }
}

<#
    .SYNOPSIS
        Saves the current application settings to the configuration file.
    .DESCRIPTION
        Writes the current state of $global:EnableLogging and $global:DarkMode to the
        'AccChecker.config' file in a "Key=Value" format. This function is called whenever
        a setting is changed in the UI.
    .NOTES
        The file is saved with UTF8 encoding to ensure compatibility.
        It will overwrite the existing file.
#>
function Save-Config {
    $configContent = @"
EnableLogging=$($global:EnableLogging)
DarkMode=$($global:DarkMode)
"@
    try {
        # Ensure the parent directory exists before writing the file. This is created on-demand.
        $parentDir = Split-Path -Parent $global:ConfigFile
        if (-not (Test-Path $parentDir)) {
            New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
        }
        $configContent | Set-Content -Path $global:ConfigFile -Encoding UTF8
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error saving configuration: $($_.Exception.Message)", "Config Save Error", "OK", "Error")
    }
}

<#
    .SYNOPSIS
        Writes the result of an account check to a daily log file.
    .DESCRIPTION
        If logging is enabled via the UI, this function appends a timestamped entry to a
        tab-separated value (TSV) file. If the log file for the current day does not exist,
        it creates the file and adds a header row first.
    .PARAMETER Email
        The email address that was checked.
#>
function Log-Result {
    param ([hashtable]$LogData)
    if ($global:EnableLogging) {
        try {
            # Ensure the logs directory exists before writing. This is created on-demand.
            $parentDir = Split-Path -Parent $global:AccountCheckerLogFile
            if (-not (Test-Path $parentDir)) {
                New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
            }

            if (-not (Test-Path $global:AccountCheckerLogFile)) {
                # Create the header row if the file doesn't exist.
                $headers = "Timestamp`tUsername`tFoundInCloud`tCloudURL`tFederationRedirectUrl`tAccountType`tLiteraOneCompatibility"
                $headers | Out-File -FilePath $global:AccountCheckerLogFile -Encoding UTF8
            }

            # Build the tab-separated log entry.
            $logEntry = "$($LogData.Timestamp)`t$($LogData.Username)`t$($LogData.FoundInCloud)`t$($LogData.CloudURL)`t$($LogData.FederationRedirectUrl)`t$($LogData.AccountType)`t$($LogData.LiteraOneCompatibility)"
            $logEntry | Out-File -Append -FilePath $global:AccountCheckerLogFile -Encoding UTF8
        } catch {
            # Silently ignore logging errors to not interrupt the user. The error could be logged elsewhere if needed.
        }
    }
}

<#
    .SYNOPSIS
        Validates if a string matches the expected format of an email address.
    .PARAMETER Email
        The string to validate.
    .RETURN
        $true if the format is valid, otherwise $false.
#>
function IsValidEmail {
    param([string]$Email)
    return $global:EmailRegex.IsMatch($Email)
}

#endregion

#region --- UI Helper Functions ---

<#
    .SYNOPSIS
        Applies the selected visual theme (dark or light) to all GUI controls.
    .DESCRIPTION
        This function centralizes all theme-related style changes. It sets the BackColor and
        ForeColor for the form and all its child controls based on the chosen theme.
    .PARAMETER dark
        If $true, applies the dark theme. If $false, applies the light theme.
#>
function Apply-Theme {
    param([bool]$dark)

    $formBg = if ($dark) { [System.Drawing.Color]::FromArgb(30,30,30) } else { [System.Drawing.SystemColors]::Control }
    $formFg = if ($dark) { [System.Drawing.Color]::White } else { [System.Drawing.Color]::Black }
    $textBoxBg = if ($dark) { [System.Drawing.Color]::FromArgb(60,60,60) } else { [System.Drawing.Color]::White } # Darker textbox for dark mode
    $textBoxFg = $formFg
    $buttonBg = if ($dark) { [System.Drawing.Color]::FromArgb(70,70,70) } else { [System.Drawing.SystemColors]::Control } # Darker button for dark mode
    $buttonFg = $formFg

    # For dark mode, apply the custom renderer to correctly style the menu bar.
    if ($dark) {
        $renderer = New-Object System.Windows.Forms.ToolStripProfessionalRenderer(New-Object ProfessionalColorTable)
        $global:UI.MenuStrip.Renderer = $renderer
    } else {
        $global:UI.MenuStrip.Renderer = $null # Revert to the default Windows renderer for light mode.
    }

    # Apply colors to the main form and its menu strip.
    $global:UI.Form.BackColor = $formBg
    $global:UI.Form.ForeColor = $formFg
    $global:UI.MenuStrip.BackColor = $formBg
    $global:UI.MenuStrip.ForeColor = $formFg

    # Apply colors to labels and text boxes.
    $global:UI.EmailLabel.ForeColor = $formFg
    $global:UI.ResultLabel.ForeColor = $formFg

    $global:UI.EmailTextBox.BackColor = $textBoxBg
    $global:UI.EmailTextBox.ForeColor = $textBoxFg

    # Apply colors to the rich text box for API responses.
    if ($global:UI.ApiResponseBox) {
        $global:UI.ApiResponseBox.BackColor = $textBoxBg
        $global:UI.ApiResponseBox.ForeColor = $textBoxFg
    }

    # Apply colors to all buttons.
    $global:UI.CheckButton.BackColor = $buttonBg
    $global:UI.CheckButton.ForeColor = $buttonFg
    $global:UI.PasteButton.BackColor = $buttonBg
    $global:UI.PasteButton.ForeColor = $buttonFg
    $global:UI.ClearButton.BackColor = $buttonBg
    $global:UI.ClearButton.ForeColor = $buttonFg
    $global:UI.CopyButton.BackColor = $buttonBg
    $global:UI.CopyButton.ForeColor = $buttonFg

    # Recursively apply theme colors to all menu and submenu items.
    foreach ($item in $global:UI.MenuStrip.Items) {
        $item.BackColor = $formBg
        $item.ForeColor = $formFg

        foreach ($subItem in $item.DropDownItems) {
            $subItem.BackColor = $formBg
            $subItem.ForeColor = $formFg
            if ($subItem -is [System.Windows.Forms.ToolStripMenuItem]) {
                foreach ($childItem in $subItem.DropDownItems) {
                    $childItem.BackColor = $formBg
                    $childItem.ForeColor = $formFg
                }
            }
        }
    }
}

<#
    .SYNOPSIS
        Manages the enabled/disabled state of UI controls during background operations.
    .DESCRIPTION
        This function provides visual feedback during long-running tasks (like API calls).
        It disables interactive controls to prevent concurrent operations and can display
        a status message in the result label.
    .PARAMETER IsBusy
        If $true, controls are disabled. If $false, they are re-enabled.
    .PARAMETER Message An optional message to display while the UI is busy.
#>
function Set-UIState {
    param([bool]$IsBusy, [string]$Message = "")
    $global:UI.CheckButton.Enabled = -not $IsBusy
    $global:UI.EmailTextBox.Enabled = -not $IsBusy
    $global:UI.PasteButton.Enabled = -not $IsBusy
    $global:UI.ClearButton.Enabled = -not $IsBusy
    $global:UI.CopyButton.Enabled = -not $IsBusy

    # If busy, update the result label with the provided status message.
    if ($IsBusy) {
        $global:UI.ResultLabel.Text = $Message
        $global:UI.ResultLabel.ForeColor = $global:UI.Form.ForeColor
    }
}

<#
    .SYNOPSIS
        Formats and displays a detailed, color-coded summary in the response text box.
    .DESCRIPTION
        This function populates the RichTextBox with a detailed breakdown of the check result.
        It builds a complete string, sets it as the text (which auto-detects URLs), and then
        applies color and bold formatting to specific keywords for readability. It handles
        three main scenarios: a successful result, an account not found, or an API error.
    .PARAMETER Email
        The email address that was checked.
#>
function Update-ApiResponseDisplay {
    param(
        $FinalResult,
        $ResultColor,
        $Message,
        $Email
    )

    $rtb = $global:UI.ApiResponseBox
    $rtb.Clear()

    # Helper scriptblock to find and format specific text within the RichTextBox.
    # This is more reliable than appending pre-formatted text segments.
    $formatText = {
        param($TextToFind, $Color, $IsBold = $false)
        $startIndex = 0
        while ($startIndex -lt $rtb.TextLength) {
            $foundIndex = $rtb.Find($TextToFind, $startIndex, [System.Windows.Forms.RichTextBoxFinds]::None)
            if ($foundIndex -eq -1) { break }

            $rtb.Select($foundIndex, $TextToFind.Length)
            if ($null -ne $Color) { $rtb.SelectionColor = $Color }
            if ($IsBold) { $rtb.SelectionFont = New-Object System.Drawing.Font($rtb.Font, [System.Drawing.FontStyle]::Bold) }

            $startIndex = $foundIndex + $TextToFind.Length
        }
        $rtb.SelectionLength = 0 # Deselect text.
        $rtb.SelectionFont = $rtb.Font # Reset font style.
    }

    if ($null -ne $FinalResult) {
        # Scenario 1: A valid account result was found.
        $response = $FinalResult.Response
        $foundInCloud = $FinalResult.Cloud
        $cloudUrl = ""

        # Determine the correct login URL based on the account type and cloud.
        if ($Message -match "Personal") {
            $cloudUrl = "https://login.live.com"
        } else {
            $endpoints = @{
                "Commercial" = "https://login.microsoftonline.com"
                "China"      = "https://login.partner.microsoftonline.cn"
                "GCCHigh"    = "https://login.microsoftonline.us"
            }
            $cloudUrl = $endpoints[$foundInCloud]
        }

        # Build the full output string before setting it in the text box.
        $output = "Username: $Email`n"
        $output += "FoundInCloud: $foundInCloud`n"
        $output += "Cloud URL: $cloudUrl`n"

        $federationUrl = $response.Credentials.FederationRedirectUrl
        if ($federationUrl) { # Only show this line if the account is federated.
            $output += "FederationRedirectUrl: $federationUrl`n" # This will also be a link.
        }
        $output += "Account Type: $Message`n"
        $output += "Litera One Compatibility: "

        if ($federationUrl -and $federationUrl -match "sso\.godaddy\.com") { # GoDaddy check has top priority.
            $output += "Incompatible with Litera One (due to GoDaddy federation)"
        }       
        elseif ($foundInCloud -eq 'GCCHigh') {
            $output += "Incompatible with Litera One (due to GCC High)"
        }
        elseif ($foundInCloud -eq 'Commercial' -and $response.IfExistsResult -in @(0, 1, 5, 6)) {
            $output += "Compatible with Litera One"
        }
        else {
            $output += "May not be Compatible with Litera One"
        }

        # Set the text box content. The RichTextBox will automatically hyperlink URLs.
        $rtb.Text = $output

        # Apply color and style formatting to key parts of the text.
        $formatText.Invoke($Email, $ResultColor)
        $formatText.Invoke($foundInCloud, $null, $true)
        $rtb.DeselectAll()

    # Scenario 2: The account was confirmed not to be a Microsoft account.
    } elseif ($Message -match "Not a Microsoft Account") {
        $rtb.Text = "If you believe this result is incorrect, please confirm with your administrator that the email address is a valid Microsoft account and is set as the primary alias."
    # Scenario 3: An error occurred, preventing a valid result.
    } else {
        $rtb.Text = "No valid API response. Error: $Message"
    }

    # Return the generated text so it can be used for logging.
    return $rtb.Text
}


#endregion

#region --- Core Logic: API Interaction ---

<#
    .SYNOPSIS
        Queries a specific Microsoft cloud endpoint to get an account's credential type.
    .DESCRIPTION
        This function sends a POST request to the 'GetCredentialType' endpoint for a given
        Microsoft cloud environment. It includes a 15-second timeout and validates that the
        API response is not null and contains the expected 'IfExistsResult' property.
    .PARAMETER Email
        The email address to check.
    .RETURN A hashtable containing the API response object and the cloud it was found in.
        Throws an exception on network errors or invalid API responses.
#>
function Check-MicrosoftAccount {
    param(
        [string]$Email,
        [string]$CloudEnvironment = "Commercial"
    )

    # Map cloud names to their respective API endpoints.
    $endpoints = @{
        "Commercial" = "https://login.microsoftonline.com/common/GetCredentialType"
        "China"      = "https://login.partner.microsoftonline.cn/common/GetCredentialType"
        "GCCHigh"    = "https://login.microsoftonline.us/common/GetCredentialType"
    }

    $url = $endpoints[$CloudEnvironment]
    if ([string]::IsNullOrEmpty($url)) { throw "Invalid Cloud Environment specified: $CloudEnvironment" }

    # Construct the JSON payload for the POST request.
    $body = @{
        Username            = $Email
        isOtherIdpSupported = $true
    } | ConvertTo-Json

    try {
        $response = Invoke-RestMethod -Uri $url -Method Post -Body $body -ContentType "application/json" -TimeoutSec 15

        # The API can sometimes return null or an unexpected object. Validate the response.
        if ($null -eq $response) {
            throw "API Response Error: Received null response from Microsoft API for $CloudEnvironment cloud."
        }
        # The 'IfExistsResult' property is critical for determining the account state.
        if (-not ($response | Get-Member -MemberType NoteProperty -Name IfExistsResult -ErrorAction SilentlyContinue)) {
            throw "API Response Error: 'IfExistsResult' property missing from Microsoft API response for $CloudEnvironment cloud."
        }

        return @{ Response = $response; Cloud = $CloudEnvironment }
    } catch {
        # Wrap any exception in a more descriptive message for easier debugging.
        throw "Network/API Error: $($_.Exception.Message)"
    }
}

#endregion

#region --- GUI Event Handlers ---

<#
    .SYNOPSIS
        Handles the "Check" button click event to orchestrate the account lookup process.
    .DESCRIPTION
        This is the core logic function. It validates the user's input, sets the UI to a busy
        state, sequentially queries the Microsoft cloud environments, processes the results to
        determine the account type, and finally updates the UI with the findings.
#>
function Handle-CheckButtonClick {
    $email = $global:UI.EmailTextBox.Text.Trim()

    # 1. Validate the user's input.
    if ([string]::IsNullOrWhiteSpace($email)) {
        $global:UI.ResultLabel.Text = "Please enter an email address."
        $global:UI.ResultLabel.ForeColor = [System.Drawing.Color]::Red
        $global:UI.EmailTextBox.ForeColor = [System.Drawing.Color]::Red # Visual feedback on textbox
        return
    }
    if (-not (IsValidEmail $email)) {
        $global:UI.ResultLabel.Text = "Invalid email format. Please check and try again."
        $global:UI.ResultLabel.ForeColor = [System.Drawing.Color]::Red
        $global:UI.EmailTextBox.ForeColor = [System.Drawing.Color]::Red # Visual feedback on textbox
        return
    }

    # 2. Reset UI colors and set the state to "busy".
    Apply-Theme -dark:$global:DarkMode # Reset colors to the current theme defaults.
    Set-UIState -IsBusy $true -Message "Contacting Microsoft..."

    # 3. Sequentially query clouds to find the account's home environment.
    $response = $null
    $errorMsg = $null
    $finalResult = $null

    # --- Stricter Sequential Cloud Checking Logic ---
    $commercialResult = $null
    try {
        # Pass 1: Always check Commercial first.
        $commercialResult = Check-MicrosoftAccount -Email $email -CloudEnvironment "Commercial"
    } catch {
        $errorMsg = $_.Exception.Message
    }

    # Define common personal account domains to aid in classification.
    $personalDomains = @("outlook.com", "hotmail.com", "live.com", "msn.com")

    # Pass 2: Check sovereign clouds (GCCHigh, China) only if the commercial check
    # explicitly indicates the account doesn't exist there (code 1). This prevents
    # incorrectly identifying a commercial account as being in a sovereign cloud.
    if ($null -ne $commercialResult -and $commercialResult.Response.IfExistsResult -eq 1 -and $null -eq $errorMsg -and ($email.Split('@')[1].ToLower() -notin $personalDomains)) {
        foreach ($cloud in @("GCCHigh", "China")) {
            # A managed (0) or federated (5) result from a sovereign cloud is a strong signal.
            try {
                $sovereignResult = Check-MicrosoftAccount -Email $email -CloudEnvironment $cloud
                if ($sovereignResult.Response.IfExistsResult -in @(0, 5)) { # Trust managed (0) and federated (5) results.
                    $finalResult = $sovereignResult
                    break # Found it, stop checking.
                }
            } catch {
                # Ignore errors from sovereign clouds; we can fall back to the commercial result.
            }
        }
    }


    # Fallback: If no definitive result was found in sovereign clouds, use the original
    # commercial cloud result. This is critical for correctly identifying commercial accounts.
    if ($null -eq $finalResult -and $null -ne $commercialResult) {
        $finalResult = $commercialResult
    }

    # 4. Process the results of the API calls.
    $msg = ""
    $resultColor = $null
    $textBoxColor = $null

    # Prepare a hashtable to hold structured data for logging.
    $logData = @{ Timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss"); Username = $email }

    # Scenario A: An error occurred during the API calls.
    if ($null -ne $errorMsg) {
        $msg = $errorMsg
        $resultColor = [System.Drawing.Color]::DarkOrange
        $textBoxColor = [System.Drawing.Color]::DarkOrange
    # Scenario B: The account was not found in any cloud.
    } elseif ($null -eq $finalResult) {
        $msg = "Not a Microsoft Account in Commercial, GCC High, or China clouds."
        $resultColor = [System.Drawing.Color]::Red
        $textBoxColor = [System.Drawing.Color]::Red

        # Populate log data for a non-Microsoft account.
        $logData.FoundInCloud = "N/A"
        $logData.CloudURL = "N/A"
        $logData.FederationRedirectUrl = "N/A"
        $logData.AccountType = "Not a Microsoft Account"
        $logData.LiteraOneCompatibility = "Not Applicable"

    # Case C: A valid response was received from one of the clouds.
    } else {
        $response = $finalResult.Response
        $foundInCloud = $finalResult.Cloud
        $domain = $email.Split('@')[1].ToLower()

        # Populate base log data from the successful result.
        $logData.FoundInCloud = $foundInCloud
        $logData.FederationRedirectUrl = $response.Credentials.FederationRedirectUrl

        # Classify the account type based on its domain, the cloud it was found in, and the API response code.
        # --- Sovereign Cloud Logic ---
        if ($foundInCloud -eq "GCCHigh") {
            $msg = "Likely a GCC High Account."
            $resultColor = [System.Drawing.Color]::DarkCyan
        } elseif ($foundInCloud -eq "China") {
            $msg = "Likely an M365 China (VNET) Account."
            $resultColor = [System.Drawing.Color]::Purple
        # --- Commercial Cloud Logic ---
        } else {
            $code = $response.IfExistsResult
            $isFederated = $response.IsFederated

            switch ($code) {
                0 {
                    if ($personalDomains -contains $domain) {
                        $msg = "Likely Personal Microsoft Account (MSA)."
                        $resultColor = [System.Drawing.Color]::Blue
                    } elseif ($isFederated) {
                        $msg = "Federated Entra ID (Work/School) Account."
                        $resultColor = [System.Drawing.Color]::DarkGray 
                    } else {
                        $msg = "Likely Entra Microsoft Account (Work/School)."
                        $resultColor = [System.Drawing.Color]::Green
                    }
                }
                1 {
                    $msg = "Not a Microsoft Account."
                    $resultColor = [System.Drawing.Color]::Red
                }
                5 {
                    if ($personalDomains -contains $domain) {
                        $msg = "Likely Personal Microsoft Account (MSA)."
                        $resultColor = [System.Drawing.Color]::Blue
                    } else {
                        $msg = "Federated Entra ID (Work/School) Account."
                        $resultColor = [System.Drawing.Color]::DarkGray 
                    }
                }
                6 {
                    # Code 6 indicates a "Dual Identity" account (both Personal and Work/School).
                    # We prioritize the Work/School identity for classification.
                    $msg = "Dual Identity (Work/School & Personal) Account."
                    $resultColor = [System.Drawing.Color]::DarkOrchid
                }
                default {
                    $msg = "Unknown Result Code: $code. Please try again or check logs."
                    $resultColor = [System.Drawing.Color]::Gray
                }
            }
        }
        $textBoxColor = $resultColor

        # Populate the remaining log data fields after classification.
        $logData.AccountType = $msg
        if ($msg -match "Not a Microsoft Account") {
            $logData.FoundInCloud = "N/A" # Clear the cloud value for non-MS accounts
            $logData.CloudURL = ""
            $logData.LiteraOneCompatibility = ""
        } else {
            if ($msg -match "Personal") {
                $logData.CloudURL = "https://login.live.com"
            } else {
                $endpoints = @{ "Commercial" = "https://login.microsoftonline.com"; "China" = "https://login.partner.microsoftonline.cn"; "GCCHigh" = "https://login.microsoftonline.us" }
                $logData.CloudURL = $endpoints[$foundInCloud]
            }
    
            # Determine Litera One compatibility for logging.
            if ($logData.FederationRedirectUrl -and $logData.FederationRedirectUrl -match "sso\.godaddy\.com") {
                $logData.LiteraOneCompatibility = "Incompatible with Litera One (due to GoDaddy federation)"
            }       
            elseif ($foundInCloud -eq 'GCCHigh') {
                $logData.LiteraOneCompatibility = "Incompatible with Litera One (due to GCC High)"
            }
            elseif ($foundInCloud -eq 'Commercial' -and $response.IfExistsResult -in @(0, 5, 6)) { # Removed code 1 from compatibility check
                $logData.LiteraOneCompatibility = "Compatible with Litera One"
            } else {
                $logData.LiteraOneCompatibility = "May not be Compatible with Litera One"
            }
        }
    }

    # 5. Update the UI with the final results.
    Set-UIState -IsBusy $false # Re-enable all controls.

    # Display the summary message and apply color coding.
    $global:UI.ResultLabel.Text = $msg
    $global:UI.ResultLabel.ForeColor = $resultColor
    $global:UI.EmailTextBox.ForeColor = $textBoxColor

    # Determine the message and result object to pass to the detailed display function.
    $finalMessage = if ($errorMsg) { $errorMsg } else { $msg }
    
    # Don't pass the result object if the account wasn't found, to avoid errors.
    $resultForDisplay = if ($finalMessage -match "Not a Microsoft Account") { $null } else { $finalResult } 

    $formattedResult = Update-ApiResponseDisplay -FinalResult $resultForDisplay -ResultColor $resultColor -Message $msg -Email $email
    Log-Result -LogData $logData
}


#endregion

#region --- GUI Construction ---

<#
    .SYNOPSIS
        Constructs and configures all Windows Forms controls for the main application window.
    .DESCRIPTION
        This function programmatically creates the entire GUI, including the main form, menu bar,
        labels, buttons, and text boxes. It assigns properties, wires up event handlers, and
        returns a hashtable containing all UI elements for easy access throughout the script.
#>
function Build-GUI {
    $ui = @{}

    # --- Main Form ---
    $ui.Form = New-Object System.Windows.Forms.Form
    $ui.Form.Name = "formMain"
    $ui.Form.Text = "Microsoft Account Checker"
    $ui.Form.Size = New-Object System.Drawing.Size(500, 460)
    $ui.Form.StartPosition = "CenterScreen"
    $ui.Form.FormBorderStyle = "FixedSingle"
    $ui.Form.MaximizeBox = $false

    # Set the form's icon to the icon of the PowerShell executable itself.
    $ui.Form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName)

    # --- Menu Strip and Items ---
    $ui.MenuStrip = New-Object System.Windows.Forms.MenuStrip
    $ui.MenuStrip.Name = "menuStripMain"

    # -- Settings Menu --
    $ui.MenuSettings = New-Object System.Windows.Forms.ToolStripMenuItem("&Settings")
    $ui.MenuSettings.Name = "menuSettings"

    $ui.LogToggle = New-Object System.Windows.Forms.ToolStripMenuItem("Enable &Logging (Account Checks)")
    $ui.LogToggle.Name = "logToggle"
    $ui.LogToggle.Checked = $global:EnableLogging
    $ui.LogToggle.CheckOnClick = $true
    $ui.LogToggle.Add_Click({
        $global:EnableLogging = $ui.LogToggle.Checked
        Save-Config
        $ui.MenuViewLog.Visible = $global:EnableLogging
        $ui.MenuExportLog.Visible = $global:EnableLogging
    })
    $ui.MenuSettings.DropDownItems.Add($ui.LogToggle)

    $ui.DarkToggle = New-Object System.Windows.Forms.ToolStripMenuItem("Enable &Dark Mode")
    $ui.DarkToggle.Name = "darkToggle"
    $ui.DarkToggle.Checked = $global:DarkMode
    $ui.DarkToggle.CheckOnClick = $true
    $ui.DarkToggle.Add_Click({
        $global:DarkMode = $ui.DarkToggle.Checked
        Save-Config
        Apply-Theme -dark:$global:DarkMode
    })
    $ui.MenuSettings.DropDownItems.Add($ui.DarkToggle)
    $ui.MenuStrip.Items.Add($ui.MenuSettings)

    # -- View Log Menu Item --
    $ui.MenuViewLog = New-Object System.Windows.Forms.ToolStripMenuItem("&View Account Log")
    $ui.MenuViewLog.Name = "menuViewLog"
    $ui.MenuViewLog.Visible = $global:EnableLogging
    $ui.MenuViewLog.Add_Click({
        if (Test-Path $global:AccountCheckerLogFile) {
            try { Start-Process notepad.exe $global:AccountCheckerLogFile }
            catch { [System.Windows.Forms.MessageBox]::Show("Could not open account log file: $($_.Exception.Message)", "Error", "OK", "Error") }
        } else {
            [System.Windows.Forms.MessageBox]::Show("No account checker log file found at '$($global:AccountCheckerLogFile)'.", "Log File Not Found")
        }
    })
    $ui.MenuStrip.Items.Add($ui.MenuViewLog)

    # -- Export Log Menu Item --
    $ui.MenuExportLog = New-Object System.Windows.Forms.ToolStripMenuItem("&Export Account Log to CSV")
    $ui.MenuExportLog.Name = "menuExportLog"
    $ui.MenuExportLog.Visible = $global:EnableLogging
    $ui.MenuExportLog.Add_Click({
        if (Test-Path $global:AccountCheckerLogFile) {
            $dialog = New-Object System.Windows.Forms.SaveFileDialog
            $dialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
            $dialog.FileName = "MS_Account_Checker_Results_$(Get-Date -Format 'yyyy-MM-dd').csv"
            $dialog.Title = "Save Account Checker Results"
            if ($dialog.ShowDialog($ui.Form) -eq "OK") {
                try {
                    # Import the tab-separated log file and export it as a proper CSV.
                    Import-Csv -Path $global:AccountCheckerLogFile -Delimiter "`t" | Export-Csv -Path $dialog.FileName -NoTypeInformation -Encoding UTF8
                    [System.Windows.Forms.MessageBox]::Show("Exported to CSV successfully to:`n$($dialog.FileName)", "Export Success", "OK", "Information")
                } catch {
                    [System.Windows.Forms.MessageBox]::Show("Error exporting to CSV: $($_.Exception.Message)", "Export Error", "OK", "Error")
                }
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("No account checker log file found to export.", "Log File Not Found")
        }
    })
    $ui.MenuStrip.Items.Add($ui.MenuExportLog)

    # -- About Menu Item --
    $ui.MenuAbout = New-Object System.Windows.Forms.ToolStripMenuItem("&About")
    $ui.MenuAbout.Name = "menuAbout"
    $ui.MenuAbout.Add_Click({
        $aboutBox = New-Object System.Windows.Forms.Form
        $aboutBox.Text = "About Microsoft Account Checker"
        $aboutBox.Size = New-Object System.Drawing.Size(400, 220)
        $aboutBox.StartPosition = "CenterParent"
        $aboutBox.FormBorderStyle = "FixedDialog"
        $aboutBox.MaximizeBox = $false
        $aboutBox.MinimizeBox = $false
        $aboutBox.ShowInTaskbar = $false
        $aboutBox.Topmost = $true

        $aboutLabel = New-Object System.Windows.Forms.Label
        $aboutLabel.Text = "Microsoft Account Checker`nVersion 1.1.0`n`nThis tool helps identify Microsoft account types across Commercial, GCC High, and China clouds."
        $aboutLabel.Location = New-Object System.Drawing.Point(20, 20)
        $aboutLabel.Size = New-Object System.Drawing.Size(350, 60)
        $aboutLabel.TextAlign = "MiddleCenter"

        $okButton = New-Object System.Windows.Forms.Button
        $okButton.Text = "OK"
        $okButton.Location = New-Object System.Drawing.Point(150, 140)
        $okButton.Size = New-Object System.Drawing.Size(100, 30)
        $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $aboutBox.CancelButton = $okButton # Allows closing with Escape key

        # Manually apply the current theme to the dynamically created About box.
        if ($global:DarkMode) {
            $aboutBox.BackColor = [System.Drawing.Color]::FromArgb(30,30,30)
            $aboutLabel.ForeColor = [System.Drawing.Color]::White
            $okButton.BackColor = [System.Drawing.Color]::FromArgb(70,70,70)
            $okButton.ForeColor = [System.Drawing.Color]::White
        }

        $aboutBox.Controls.Add($aboutLabel)
        $aboutBox.Controls.Add($okButton)
        $aboutBox.ShowDialog($ui.Form)
    })
    $ui.MenuStrip.Items.Add($ui.MenuAbout)

    $ui.Form.Controls.Add($ui.MenuStrip)
    $ui.Form.MainMenuStrip = $ui.MenuStrip

    # --- Main Controls ---
    # -- Email Label --
    $ui.EmailLabel = New-Object System.Windows.Forms.Label
    $ui.EmailLabel.Text = "Enter email address:"
    $ui.EmailLabel.Location = New-Object System.Drawing.Point(20, 40)
    $ui.EmailLabel.Size = New-Object System.Drawing.Size(440, 20)
    $ui.Form.Controls.Add($ui.EmailLabel)

    # -- Email Textbox --
    $ui.EmailTextBox = New-Object System.Windows.Forms.TextBox
    $ui.EmailTextBox.Location = New-Object System.Drawing.Point(20, 65)
    $ui.EmailTextBox.Size = New-Object System.Drawing.Size(440, 25)
    $ui.EmailTextBox.Add_KeyUp({ # Reset result colors when the user types a new email.
        $ui.ResultLabel.Text = ""
        $defaultFg = if ($global:DarkMode) { [System.Drawing.Color]::White } else { [System.Drawing.Color]::Black }
        $defaultBg = if ($global:DarkMode) { [System.Drawing.Color]::FromArgb(60,60,60) } else { [System.Drawing.Color]::White }
        $ui.EmailTextBox.ForeColor = $defaultFg
        $ui.EmailTextBox.BackColor = $defaultBg
    })
    $ui.EmailTextBox.Add_KeyDown({ if ($_.KeyCode -eq "Enter") { $ui.CheckButton.PerformClick() } })
    $ui.Form.Controls.Add($ui.EmailTextBox)

    # -- Result Label (for summary messages) --
    $ui.ResultLabel = New-Object System.Windows.Forms.Label
    $ui.ResultLabel.Location = New-Object System.Drawing.Point(20, 100)
    $ui.ResultLabel.Size = New-Object System.Drawing.Size(440, 40)
    $ui.Form.Controls.Add($ui.ResultLabel)

    # -- Check Button --
    $ui.CheckButton = New-Object System.Windows.Forms.Button
    $ui.CheckButton.Text = "Check"
    $ui.CheckButton.Location = New-Object System.Drawing.Point(20, 150)
    $ui.CheckButton.Size = New-Object System.Drawing.Size(100, 30)
    $ui.CheckButton.Add_Click({ Handle-CheckButtonClick })
    $ui.Form.Controls.Add($ui.CheckButton)

    # -- Paste Button --
    $ui.PasteButton = New-Object System.Windows.Forms.Button
    $ui.PasteButton.Text = "Paste"
    $ui.PasteButton.Location = New-Object System.Drawing.Point(130, 150)
    $ui.PasteButton.Size = New-Object System.Drawing.Size(100, 30)
    $ui.PasteButton.Add_Click({
        $ui.EmailTextBox.Text = [System.Windows.Forms.Clipboard]::GetText().Trim()
        $ui.ResultLabel.Text = ""
        $defaultFg = if ($global:DarkMode) { [System.Drawing.Color]::White } else { [System.Drawing.Color]::Black }
        $ui.EmailTextBox.ForeColor = $defaultFg
    })
    $ui.Form.Controls.Add($ui.PasteButton)

    # -- Clear Button --
    $ui.ClearButton = New-Object System.Windows.Forms.Button
    $ui.ClearButton.Text = "Clear"
    $ui.ClearButton.Location = New-Object System.Drawing.Point(240, 150)
    $ui.ClearButton.Size = New-Object System.Drawing.Size(100, 30)
    $ui.ClearButton.Add_Click({
        $ui.EmailTextBox.Clear()
        $ui.ResultLabel.Text = ""
        $ui.ApiResponseBox.Clear()
        Apply-Theme -dark:$global:DarkMode
    })
    $ui.Form.Controls.Add($ui.ClearButton)

    # -- Copy Button (for copying results) --
    $ui.CopyButton = New-Object System.Windows.Forms.Button
    $ui.CopyButton.Text = "Copy"
    $ui.CopyButton.Location = New-Object System.Drawing.Point(350, 150)
    $ui.CopyButton.Size = New-Object System.Drawing.Size(100, 30)
    $ui.CopyButton.Add_Click({
        $textToCopy = $ui.ApiResponseBox.Text

        # Provide feedback if the user tries to copy an empty response.
        if ([string]::IsNullOrWhiteSpace($textToCopy)) {
            [System.Windows.Forms.MessageBox]::Show("There is no text in the response box to copy.", "Nothing to Copy", "OK", "Information")
        } else {
            try {
                Set-Clipboard -Value $textToCopy
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Failed to copy text to clipboard: $($_.Exception.Message)", "Copy Error", "OK", "Error")
            }
        }
    })
    $ui.Form.Controls.Add($ui.CopyButton)

    # -- API Response RichTextBox (for detailed results) --
    $ui.ApiResponseBox = New-Object System.Windows.Forms.RichTextBox
    $ui.ApiResponseBox.Location = New-Object System.Drawing.Point(20, 220)
    $ui.ApiResponseBox.Size = New-Object System.Drawing.Size(440, 180)
    $ui.ApiResponseBox.ReadOnly = $true
    $ui.ApiResponseBox.Font = New-Object System.Drawing.Font("Consolas", 8)
    $ui.ApiResponseBox.DetectUrls = $true # Enable automatic URL detection and hyperlinking.
    $ui.ApiResponseBox.Add_LinkClicked({
        param($sender, $e)
        try {
            # Use Start-Process to open the clicked link in the user's default browser.
            Start-Process $e.LinkText -ErrorAction Stop
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Could not open the link: $($e.LinkText)`n`nError: $($_.Exception.Message)", "Link Error", "OK", "Error")
        }
    })
    $ui.Form.Controls.Add($ui.ApiResponseBox)

    return $ui
}

#endregion

#region --- Application Entry Point ---

# 1. Load user settings from the config file.
Load-Config

# 2. Build the GUI and store control objects in a global variable.
$global:UI = Build-GUI

# 3. Set the initial state of menu items to reflect the loaded configuration.
$global:UI.LogToggle.Checked = $global:EnableLogging
$global:UI.DarkToggle.Checked = $global:DarkMode

# 4. Apply the visual theme based on the loaded configuration.
Apply-Theme -dark:$global:DarkMode

# 5. Set initial focus to the email input box for a better user experience.
$global:UI.Form.Add_Shown({ $global:UI.EmailTextBox.Select() })

# 6. Show the form. The script will wait here until the user closes the window.
[void]$global:UI.Form.ShowDialog()

#endregion