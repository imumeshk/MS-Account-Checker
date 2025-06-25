Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- Global Variables & Paths ---
# Script root directory (where the script resides)
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

# Define data directories relative to the script's location
$dataDir    = Join-Path (Split-Path -Parent $scriptRoot) "Data"
$configDir  = Join-Path $dataDir "config"
$logsDir    = Join-Path $dataDir "logs"

# Ensure required folders exist
foreach ($dir in @($dataDir, $configDir, $logsDir)) {
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir | Out-Null
    }
}
# Global variables for application state
$global:EnableLogging = $false
$global:DarkMode = $false
# Dynamic daily log file name, now consistent with previous versions' intention
$global:AccountCheckerLogFile = Join-Path $logsDir "MS_Account_Checker_Logs_$(Get-Date -Format 'yyyy-MM-dd').txt" 
$global:ConfigFile = Join-Path $configDir "AccChecker.config"

# Pre-compile regex for email validation for slight performance gain
$global:EmailRegex = [regex]'^[\w.-]+@[\w.-]+\.[a-zA-Z]{2,}$' 

# --- Configuration Functions ---

# Load config from simple Key=Value format
function Load-Config {
    if (Test-Path $global:ConfigFile) {
        try {
            $lines = Get-Content $global:ConfigFile
            foreach ($line in $lines) {
                if ($line -match "^EnableLogging\s*=\s*(true|false)$") {
                    $global:EnableLogging = ($Matches[1] -eq "true")
                }
                if ($line -match "^DarkMode\s*=\s*(true|false)$") {
                    $global:DarkMode = ($Matches[1] -eq "true")
                }
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to read config: $_", "Config Error", "OK", "Error")
        }
    }
}
# Save config to simple Key=Value format
function Save-Config {
    try {
        @"
EnableLogging=$($global:EnableLogging)
DarkMode=$($global:DarkMode)
"@ | Set-Content -Path $global:ConfigFile -Encoding UTF8
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error saving configuration: $($_.Exception.Message)", "Config Save Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

# --- Logging Functions ---

# Log result to the daily log file
function Log-Result {
    param ([string]$Email, [string]$Result)
    if ($global:EnableLogging) {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        # Ensure the correct global log file variable is used
        if (-not (Test-Path $global:AccountCheckerLogFile)) { 
            "Timestamp`tEmail`tResult" | Out-File -FilePath $global:AccountCheckerLogFile -Encoding UTF8
        }
        "$timestamp`t$Email`t$Result" | Out-File -Append -FilePath $global:AccountCheckerLogFile -Encoding UTF8
    }
}

# --- Validation Functions ---

# Validate email format using pre-compiled regex
function IsValidEmail {
    param([string]$Email)
    return $global:EmailRegex.IsMatch($Email)
}

# --- UI Helper Functions ---

# Apply theme to the main form and its controls
function Apply-Theme {
    param([bool]$dark)
    
    $formBg = if ($dark) { [System.Drawing.Color]::FromArgb(30,30,30) } else { [System.Drawing.SystemColors]::Control }
    $formFg = if ($dark) { [System.Drawing.Color]::White } else { [System.Drawing.Color]::Black }
    $textBoxBg = if ($dark) { [System.Drawing.Color]::FromArgb(60,60,60) } else { [System.Drawing.Color]::White } # Darker textbox for dark mode
    $textBoxFg = $formFg
    $buttonBg = if ($dark) { [System.Drawing.Color]::FromArgb(70,70,70) } else { [System.Drawing.SystemColors]::Control } # Darker button for dark mode
    $buttonFg = $formFg

    # Apply colors to Form and MenuStrip
    $global:formMain.BackColor = $formBg
    $global:formMain.ForeColor = $formFg
    $global:menuStripMain.BackColor = $formBg
    $global:menuStripMain.ForeColor = $formFg

    # Apply colors to specific controls
    $global:lblEmailLabel.ForeColor = $formFg
    $global:lblResult.ForeColor = $formFg # Default for result label, overridden by API check results

    $global:txtEmail.BackColor = $textBoxBg
    $global:txtEmail.ForeColor = $textBoxFg

    $global:btnCheck.BackColor = $buttonBg
    $global:btnCheck.ForeColor = $buttonFg
    $global:btnPaste.BackColor = $buttonBg
    $global:btnPaste.ForeColor = $buttonFg
    $global:btnClear.BackColor = $buttonBg
    $global:btnClear.ForeColor = $buttonFg

    # Apply to all dropdown items in the menu, adjusting for background
    foreach ($item in $global:menuStripMain.Items) {
        $item.BackColor = $formBg
        $item.ForeColor = $formFg
        foreach ($subItem in $item.DropDownItems) {
            $subItem.BackColor = $formBg
            $subItem.ForeColor = $formFg
            # For menu item text, ensure it matches the forecolor
            if ($subItem -is [System.Windows.Forms.ToolStripMenuItem]) {
                foreach ($childItem in $subItem.DropDownItems) {
                    $childItem.BackColor = $formBg
                    $childItem.ForeColor = $formFg
                }
            }
        }
    }
}

# Set UI state (enable/disable buttons, show messages) during API calls
function Set-UIState {
    param([bool]$IsBusy, [string]$Message = "")
    $global:btnCheck.Enabled = -not $IsBusy
    $global:txtEmail.Enabled = -not $IsBusy
    $global:btnPaste.Enabled = -not $IsBusy
    $global:btnClear.Enabled = -not $IsBusy

    # Optionally display a busy message
    if ($IsBusy) {
        $global:lblResult.Text = $Message
        # Use theme's default foreground color for busy message
        $global:lblResult.ForeColor = $global:formMain.ForeColor 
        $global:txtEmail.ForeColor = $global:txtEmail.ForeColor # Maintain current text color for busy state
    }
}

# --- Core Logic: Check Microsoft Account ---

# Function to call the Microsoft API
function Check-MicrosoftAccount {
    param([string]$Email)
    $url = "https://login.microsoftonline.com/common/GetCredentialType"
    $body = @{
        Username = $Email
        isOtherIdpSupported = $true
        checkPhones = $false
        isRemoteNGCSupported = $true
        isCookieBannerShown = $false
        isFidoSupported = $false
        forceotclogin = $false
        otclogindisallowed = $false
        tx = ""
        loginChallenge = ""
        showOneTimeCode = $false
        isExternalFederationDisallowed = $false
        isRemoteConnectSupported = $false
    } | ConvertTo-Json -Depth 10

    try {
        $response = Invoke-RestMethod -Uri $url -Method Post -Body $body -ContentType "application/json" -TimeoutSec 10

        # --- Enhanced checks for API response validity ---
        if ($null -eq $response) { # If Invoke-RestMethod returns null without throwing
            throw "API Response Error: Received null response from Microsoft API." 
        }
        # Check if the critical 'IfExistsResult' property exists
        if (-not ($response | Get-Member -MemberType NoteProperty -Name IfExistsResult -ErrorAction SilentlyContinue)) {
            throw "API Response Error: 'IfExistsResult' property missing from Microsoft API response."
        }
        # --- End enhanced checks ---

        return $response
    } catch {
        # Catch specific network/API errors and re-throw with a more descriptive message
        throw "Network/API Error: $($_.Exception.Message)"
    }
}

# --- GUI Event Handlers ---

function Handle-CheckButtonClick {
    $email = $global:txtEmail.Text.Trim()

    if ([string]::IsNullOrWhiteSpace($email)) {
        $global:lblResult.Text = "Please enter an email address."
        $global:lblResult.ForeColor = [System.Drawing.Color]::Red
        $global:txtEmail.ForeColor = [System.Drawing.Color]::Red # Visual feedback on textbox
        return
    }
    if (-not (IsValidEmail $email)) {
        $global:lblResult.Text = "Invalid email format. Please check and try again."
        $global:lblResult.ForeColor = [System.Drawing.Color]::Red
        $global:txtEmail.ForeColor = [System.Drawing.Color]::Red # Visual feedback on textbox
        return
    }

    # Reset textbox and label colors before starting new check, using current theme's default foreground
    Apply-Theme -dark:$global:DarkMode # Reapply theme to ensure base colors are set
    
    Set-UIState -IsBusy $true -Message "Contacting Microsoft..."

    # --- Synchronous API Call (reverted from asynchronous) ---
    $response = $null
    $errorMsg = $null

    try {
        $response = Check-MicrosoftAccount -Email $email
    } catch {
        $errorMsg = $_.Exception.Message
    } finally {
        Set-UIState -IsBusy $false # Re-enable UI elements immediately
    }

    $msg = ""
    $resultColor = $null # Will be set based on API result
    $textBoxColor = $null # Will be set based on API result

    if ($null -ne $errorMsg) { # An error occurred during the API call
        $msg = $errorMsg
        $resultColor = [System.Drawing.Color]::DarkOrange
        $textBoxColor = [System.Drawing.Color]::DarkOrange
    } elseif ($null -eq $response) { # Should be rare with enhanced error handling, but as a fallback
        $msg = "Unknown Error: Could not get a valid response (internal)."
        $resultColor = [System.Drawing.Color]::DarkRed
        $textBoxColor = [System.Drawing.Color]::DarkRed
    } else { # Valid response received
        $code = $response.IfExistsResult
        $isFederated = $response.IsFederated
        $domain = $email.Split('@')[1].ToLower()
        $personalDomains = @("outlook.com", "hotmail.com", "live.com", "msn.com", "outlook.in", "live.in", "msn.in") # Extended domains

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
                $textBoxColor = $resultColor
            }
            1 {
                $msg = "Not a Microsoft Account."
                $resultColor = [System.Drawing.Color]::Red
                $textBoxColor = $resultColor
            }
            5 {
                if ($personalDomains -contains $domain) {
                    $msg = "Likely Personal Microsoft Account (MSA)."
                    $resultColor = [System.Drawing.Color]::Blue
                } else {
                    $msg = "Federated Entra ID (Work/School) Account."
                    $resultColor = [System.Drawing.Color]::DarkGray 
                }
                $textBoxColor = $resultColor
            }
            default {
                $msg = "Unknown Result Code: $code. Please try again or check logs."
                $resultColor = [System.Drawing.Color]::Gray
                $textBoxColor = [System.Drawing.Color]::Gray
            }
        }
    }
    $global:lblResult.Text = $msg
    $global:lblResult.ForeColor = $resultColor
    $global:txtEmail.ForeColor = $textBoxColor # Apply color to textbox
    Log-Result -Email $email -Result $msg
}

# --- GUI Setup ---

Load-Config # This will now ensure the config directory exists

# Main Form
$global:formMain = New-Object System.Windows.Forms.Form
$global:formMain.Name = "formMain"
$global:formMain.Text = "Microsoft Account Checker"
$global:formMain.Size = New-Object System.Drawing.Size(500, 400)
$global:formMain.StartPosition = "CenterScreen"
$global:formMain.FormBorderStyle = "FixedSingle"
$global:formMain.MaximizeBox = $false

# Menu Strip
$global:menuStripMain = New-Object System.Windows.Forms.MenuStrip
$global:menuStripMain.Name = "menuStripMain"

# Settings Menu
$global:menuSettings = New-Object System.Windows.Forms.ToolStripMenuItem("&Settings")
$global:menuSettings.Name = "menuSettings"

$global:logToggle = New-Object System.Windows.Forms.ToolStripMenuItem("Enable &Logging (Account Checks)")
$global:logToggle.Name = "logToggle"
$global:logToggle.Checked = $global:EnableLogging
$global:logToggle.CheckOnClick = $true
$global:logToggle.Add_Click({
    $global:EnableLogging = $global:logToggle.Checked
    Save-Config
    $global:menuViewLog.Visible = $global:EnableLogging
    $global:menuExportLog.Visible = $global:EnableLogging
})
$global:menuSettings.DropDownItems.Add($global:logToggle)

$global:darkToggle = New-Object System.Windows.Forms.ToolStripMenuItem("Enable &Dark Mode")
$global:darkToggle.Name = "darkToggle"
$global:darkToggle.Checked = $global:DarkMode
$global:darkToggle.CheckOnClick = $true
$global:darkToggle.Add_Click({
    $global:DarkMode = $global:darkToggle.Checked
    Apply-Theme -dark:$global:DarkMode
    Save-Config
})
$global:menuSettings.DropDownItems.Add($global:darkToggle)

$global:menuStripMain.Items.Add($global:menuSettings)

# View Log Menu Item (for Account Checker Logs)
$global:menuViewLog = New-Object System.Windows.Forms.ToolStripMenuItem("&View Account Log")
$global:menuViewLog.Name = "menuViewLog"
$global:menuViewLog.Visible = $global:EnableLogging
$global:menuViewLog.Add_Click({
    if (Test-Path $global:AccountCheckerLogFile) {
        try { Start-Process notepad.exe $global:AccountCheckerLogFile }
        catch { [System.Windows.Forms.MessageBox]::Show("Could not open account log file: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) }
    } else {
        [System.Windows.Forms.MessageBox]::Show("No account checker log file found at '$($global:AccountCheckerLogFile)'.", "Log File Not Found")
    }
})
$global:menuStripMain.Items.Add($global:menuViewLog)

# Export Log Menu Item (for Account Checker Logs)
$global:menuExportLog = New-Object System.Windows.Forms.ToolStripMenuItem("&Export Account Log to CSV")
$global:menuExportLog.Name = "menuExportLog"
$global:menuExportLog.Visible = $global:EnableLogging
$global:menuExportLog.Add_Click({
    if (Test-Path $global:AccountCheckerLogFile) {
        $dialog = New-Object System.Windows.Forms.SaveFileDialog
        $dialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        $dialog.FileName = "MS_Account_Checker_Results_$(Get-Date -Format 'yyyy-MM-dd').csv"
        $dialog.Title = "Save Account Checker Results"
        if ($dialog.ShowDialog($global:formMain) -eq [System.Windows.Forms.DialogResult]::OK) {
            $csvPath = $dialog.FileName
            try {
                # Read content, skip header, split by tab, convert to PSCustomObject, then export
                $lines = Get-Content $global:AccountCheckerLogFile | Select-Object -Skip 1
                $lines | ForEach-Object {
                    $parts = $_ -split "`t"
                    [PSCustomObject]@{
                        Timestamp = $parts[0]
                        Email = $parts[1]
                        Result = $parts[2]
                    }
                } | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8 -Force
                [System.Windows.Forms.MessageBox]::Show("Exported to CSV successfully to:`n$csvPath", "Export Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Error exporting to CSV: $($_.Exception.Message)", "Export Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("No account checker log file found to export.", "Log File Not Found")
    }
})
$global:menuStripMain.Items.Add($global:menuExportLog)

# About Menu Item
$global:menuAbout = New-Object System.Windows.Forms.ToolStripMenuItem("&About")
$global:menuAbout.Name = "menuAbout"
$global:menuAbout.Add_Click({
    $aboutBox = New-Object System.Windows.Forms.Form
    $aboutBox.Text = "About Microsoft Account Checker"
    $aboutBox.Size = New-Object System.Drawing.Size(400, 250)
    $aboutBox.StartPosition = "CenterParent"
    $aboutBox.FormBorderStyle = "FixedDialog"
    $aboutBox.MaximizeBox = $false
    $aboutBox.MinimizeBox = $false
    $aboutBox.ShowInTaskbar = $false
    Apply-Theme -dark:$global:DarkMode # Apply main theme to about box as well

    $aboutLabel = New-Object System.Windows.Forms.Label
    $aboutLabel.Text = "Microsoft Account Checker`nVersion 1.0`n`nThis tool helps identify Microsoft account types."
    $aboutLabel.Location = New-Object System.Drawing.Point(20, 20)
    $aboutLabel.Size = New-Object System.Drawing.Size(350, 80)
    $aboutLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $aboutBox.Controls.Add($aboutLabel)

    $btnOpenReadme = New-Object System.Windows.Forms.Button
    $btnOpenReadme.Text = "Open Readme"
    $btnOpenReadme.Location = New-Object System.Drawing.Point(140, 120)
    $btnOpenReadme.Size = New-Object System.Drawing.Size(100, 30)
    $btnOpenReadme.Add_Click({
    $rootPath = Split-Path $scriptRoot -Parent
    $readmePath = Join-Path $rootPath "README.md"
        if (Test-Path $readmePath) {
            try {  Start-Process "notepad.exe" -ArgumentList "`"$readmePath`"" }
            catch { [System.Windows.Forms.MessageBox]::Show("Could not open README.md: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) }
        } else {
            [System.Windows.Forms.MessageBox]::Show("README.md not found in the root directory.", "File Not Found")
        }
    })
    $aboutBox.Controls.Add($btnOpenReadme)

    $aboutBox.ShowDialog($global:formMain)
})
$global:menuStripMain.Items.Add($global:menuAbout)


$global:formMain.Controls.Add($global:menuStripMain)
$global:formMain.MainMenuStrip = $global:menuStripMain

# Email Label
$global:lblEmailLabel = New-Object System.Windows.Forms.Label
$global:lblEmailLabel.Name = "lblEmailLabel"
$global:lblEmailLabel.Text = "Enter email address:"
$global:lblEmailLabel.Location = New-Object System.Drawing.Point(20, 40)
$global:lblEmailLabel.Size = New-Object System.Drawing.Size(440, 20)
$global:formMain.Controls.Add($global:lblEmailLabel)

# Email Textbox
$global:txtEmail = New-Object System.Windows.Forms.TextBox
$global:txtEmail.Name = "txtEmail"
$global:txtEmail.Location = New-Object System.Drawing.Point(20, 65)
$global:txtEmail.Size = New-Object System.Drawing.Size(440, 25)
$global:txtEmail.Add_KeyUp({
    if ([string]::IsNullOrWhiteSpace($global:txtEmail.Text.Trim())) {
        $global:lblResult.Text = ""
        Apply-Theme -dark:$global:DarkMode # Reset to default theme colors for all controls, including txtEmail
    }
})
$global:txtEmail.Add_KeyDown({
    if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $global:btnCheck.PerformClick()
    }
})
$global:formMain.Controls.Add($global:txtEmail)


$global:lblResult = New-Object System.Windows.Forms.Label
$global:lblResult.Name = "lblResult"
$global:lblResult.Location = New-Object System.Drawing.Point(20, 100)
$global:lblResult.Size = New-Object System.Drawing.Size(440, 40)
$global:formMain.Controls.Add($global:lblResult)

$global:btnCheck = New-Object System.Windows.Forms.Button
$global:btnCheck.Name = "btnCheck"
$global:btnCheck.Text = "Check"
$global:btnCheck.Location = New-Object System.Drawing.Point(20, 180)
$global:btnCheck.Size = New-Object System.Drawing.Size(100, 30)
$global:btnCheck.Add_Click({ Handle-CheckButtonClick })
$global:formMain.Controls.Add($global:btnCheck)

$global:btnPaste = New-Object System.Windows.Forms.Button
$global:btnPaste.Name = "btnPaste" 
$global:btnPaste.Text = "Paste"
$global:btnPaste.Location = New-Object System.Drawing.Point(130, 180)
$global:btnPaste.Size = New-Object System.Drawing.Size(100, 30)
$global:btnPaste.Add_Click({ $global:txtEmail.Text = [System.Windows.Forms.Clipboard]::GetText().Trim() })
$global:formMain.Controls.Add($global:btnPaste)

$global:btnClear = New-Object System.Windows.Forms.Button
$global:btnClear.Name = "btnClear"
$global:btnClear.Text = "Clear"
$global:btnClear.Location = New-Object System.Drawing.Point(240, 180)
$global:btnClear.Size = New-Object System.Drawing.Size(100, 30)
$global:btnClear.Add_Click({ 
    $global:txtEmail.Clear(); 
    $global:lblResult.Text = ""; 
    Apply-Theme -dark:$global:DarkMode # Reset textbox and label colors to theme defaults
})
$global:formMain.Controls.Add($global:btnClear)

# Initial theme application and form display
Apply-Theme -dark:$global:DarkMode
$global:formMain.Topmost = $true
[void]$global:formMain.ShowDialog()

# Focus cursor in email field on form load
$global:formMain.Add_Shown({ $global:txtEmail.Select() })