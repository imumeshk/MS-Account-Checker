Set objShell = CreateObject("Wscript.Shell")
objShell.Run "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File ""script\MS_Account_Checker.ps1""", 0