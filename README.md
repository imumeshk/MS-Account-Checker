# Microsoft Account Checker GUI Tool

This is a PowerShell-based GUI application that allows users to check whether an email address is associated with a Microsoft account. Refer to the `README.md` file in the root directory for detailed usage instructions. It uses the Microsoft login API to determine the type of account (Personal, Entra ID, Federated, or not a Microsoft account).

---

## 🚀 How to Run

To launch the application:

1. Use the `launcher.vbs` script provided in the root directory.
2. Double-click `launcher.vbs` to start the GUI.
3. The GUI will open, allowing you to enter an email address and check its Microsoft account status.

---

## 📁 Folder Structure

```
MS Account Checker/
├── Data/
│   ├── config/         # Stores configuration file (AccChecker.config)
│   └── logs/           # Stores daily log files
├── script/             # Contains the main PowerShell script (MS_Account_Checker.ps1)
├── launcher.vbs        # VBScript launcher for the GUI
└── README.md           # This file
```

---

## ⚙️ Settings

Settings are stored in `Data/config/AccChecker.config` in a simple `Key=Value` format.

| Key              | Description                                 |
|------------------|---------------------------------------------|
| `EnableLogging`  | `true`/`false`: Enables or disables logging |
| `DarkMode`       | `true`/`false`: Enables or disables dark mode for the GUI |

These settings can be toggled from the **Settings** menu in the GUI.

---

## 🧾 Output & Logging

- Logs are saved in `Data/logs/` with filenames like `MS_Account_Checker_Logs_YYYY-MM-DD.txt`.
- Each log entry includes:
  - Timestamp
  - Email address
  - Result status
- Logs can be **viewed or exported to CSV** from the GUI's menu.

---

## 🎨 Color Code Meanings

The GUI uses color indicators to show the status of the account check.

| Color      | Meaning                                 |
|------------|-----------------------------------------|
| 🔵 Blue     | Likely Personal Microsoft Account (MSA) |
| 🟢 Green    | Likely Entra Microsoft Account (Work/School) |
| ⚫ Dark Gray| Federated Entra ID Account              |
| 🔴 Red      | Not a Microsoft Account / Invalid Email |
| 🟠 Orange   | Network or API Error                    |
| ⚪ Gray     | Unknown Result Code                     |

---

## 👨‍💻 Developer Section

### API Checking Logic

The `Check-MicrosoftAccount` function is the core of the tool. It interacts with the Microsoft API:

**Endpoint:**
```
https://login.microsoftonline.com/common/GetCredentialType
```

**Request Payload:**
```json
{
  "Username": "user@example.com"
}
```

### Response Handling

The response includes `IfExistsResult` and related fields:

| Code | Meaning                             |
|------|-------------------------------------|
| 0    | Account exists                      |
| 1    | Account does not exist              |
| 5    | Federated account                   |

Classification logic:
- **Personal Microsoft Account (MSA):** Domains like `outlook.com`, `hotmail.com`, etc.
- **Entra ID (Work/School):** Corporate/organizational domains with `IsFederated = false`
- **Federated Account:** `IsFederated = true`

The result is:
- Displayed in the GUI
- Color-coded as per result
- Logged (if enabled)

---

## 📌 Requirements

- Windows OS (Tested on Windows 10/11)
- PowerShell 5.1+
- Internet connectivity to access Microsoft login API

---

## 🛠 Troubleshooting

- **Script fails to run?** Right-click the `.ps1` or `.vbs` file → Properties → Unblock.
- **Dark mode not applying?** Ensure `DarkMode=true` in `AccChecker.config`.
- **No logs created?** Ensure `EnableLogging=true` in `AccChecker.config` and the script has write permissions.

---

## 📜 License

This tool is distributed under the MIT License.