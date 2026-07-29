$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut("$env:USERPROFILE\Desktop\Portal BAU.lnk")
$Shortcut.TargetPath = "C:\Program Files\Python\python.exe"
$Shortcut.Arguments = "`"D:\00bau\Tool\Update_Phone_Process_FileCSV\app.py`""
$Shortcut.WorkingDirectory = "D:\00bau\Tool\Update_Phone_Process_FileCSV"
$Shortcut.Description = "Portal BAU - http://localhost:5000"
$Shortcut.Save()
Write-Host "Da tao shortcut tren Desktop: Portal BAU"
