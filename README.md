# Microsoft Account Checker GUI Tool (v1.1.0)

This is a PowerShell-based GUI application that allows users to check whether an email address is associated with a Microsoft account. It identifies the account type across multiple Microsoft clouds (**Commercial, GCC High, and China**) and provides a compatibility assessment for **Litera One**.

---

## 🚀 How to Run

To launch the application:

1.  Navigate to the root directory of the project.
2.  Double-click `launcher.vbs` to start the GUI without a PowerShell window.
3.  The GUI will open, allowing you to enter an email address and check its Microsoft account status.

---

## 📁 Folder Structure

The script automatically creates `config` and `logs` directories inside the `script` folder as needed.

```
MS-Account-Checker/
├── script/
│   ├── MS_Account_Checker.ps1    # Main PowerShell script
│   ├── config/                   # Stores configuration file (AccChecker.config)
│   └── logs/                     # Stores daily log files
├── launcher.vbs                  # VBScript launcher for the GUI
└── README.md                     # This file
```

---

## ⚙️ Settings

Settings are stored in `script/config/AccChecker.config` in a simple `Key=Value` format.

| Key             | Description                                          |
| --------------- | ---------------------------------------------------- |
| `EnableLogging` | `true`/`false`: Enables or disables logging.         |
| `DarkMode`      | `true`/`false`: Enables or disables dark mode for the GUI. |

These settings can be toggled from the **Settings** menu in the GUI.

---

## 🧾 Output & Logging

-   When enabled, logs are saved in `script/logs/` with filenames like `MS_Account_Checker_Logs_YYYY-MM-DD.txt`.
-   The log is a tab-separated value (TSV) file.
-   Each log entry includes:
    -   `Timestamp`
    -   `Username`
    -   `FoundInCloud` (e.g., Commercial, GCCHigh)
    -   `CloudURL`
    -   `FederationRedirectUrl` (if applicable)
    -   `AccountType` (e.g., Personal, Entra ID)
    -   `LiteraOneCompatibility`
-   Logs can be **viewed or exported to CSV** from the GUI's menu.

---

## 🎨 Color Code Meanings

The GUI uses color indicators to show the status of the account check.

| Color          | Meaning                                      |
| -------------- | -------------------------------------------- |
| 🔵 Blue         | Likely Personal Microsoft Account (MSA)      |
| 🟢 Green        | Likely Entra Microsoft Account (Work/School) |
| ⚫ Dark Gray    | Federated Entra ID Account                   |
| 🟣 Dark Orchid  | Dual Identity (Work/School Personal)       |
| 🔵 Dark Cyan    | Likely a GCC High Account                    |
| 🟣 Purple       | Likely an M365 China (VNET) Account          |
| 🔴 Red          | Not a Microsoft Account / Invalid Email      |
| 🟠 Dark Orange  | Network or API Error                         |
| ⚫ Gray         | Unknown Result Code                          |

---

## 🔬 Litera One Compatibility

The tool provides a basic compatibility check for Litera One based on the account type:
-   **Incompatible**: Accounts federated with GoDaddy (`sso.godaddy.com`) or located in the `GCCHigh` cloud.
-   **Compatible**: Commercial accounts with specific `IfExistsResult` codes (0, 5, 6).
-   **May not be Compatible**: Other scenarios, including personal accounts.

---

## 👨‍💻 Developer Section

### API Checking Logic

The `Check-MicrosoftAccount` function queries Microsoft's `GetCredentialType` endpoint. It performs a sequential check, starting with the `Commercial` cloud. If the account is not found there, it proceeds to check sovereign clouds (`GCCHigh`, `China`).

**Endpoints:**

| Cloud        | Endpoint URL                                                |
|--------------|-------------------------------------------------------------|
| Commercial   | `https://login.microsoftonline.com/common/GetCredentialType`  |
| China        | `https://login.partner.microsoftonline.cn/common/GetCredentialType` |
| GCCHigh      | `https://login.microsoftonline.us/common/GetCredentialType`   |

**Request Payload:**
```json
{
  "Username": "user@example.com",
  "isOtherIdpSupported": true
}
```

### Response Handling

The response includes `IfExistsResult` and other properties that determine the account type.

| Code | Meaning                             | Notes                                      |
|------|-------------------------------------|--------------------------------------------|
| 0    | Account exists (Managed)            | Standard work/school or personal account.  |
| 1    | Account does not exist              |                                            |
| 5    | Federated account                   | Account authenticates with a third-party IdP.|
| 6    | Dual Identity                       | Both a Personal and Work/School account exist. |

The result is:
-   Displayed in the GUI with details.
-   Color-coded based on the outcome.
-   Logged (if enabled).

---

## 📌 Requirements

-   Windows OS (Tested on Windows 10/11)
-   PowerShell 5.1+
-   Internet connectivity to access Microsoft login APIs.

---

## 🛠 Troubleshooting

-   **Script fails to run?** Right-click the `launcher.vbs` file → **Properties** → **Unblock**.
-   **Dark mode not applying?** Ensure `DarkMode=true` in `script/config/AccChecker.config`.
-   **No logs created?** Ensure `EnableLogging=true` in `script/config/AccChecker.config` and that the script has write permissions to the `script/logs/` directory.

---

## 📜 License

This tool is distributed under the MIT License.