Set WshShell = CreateObject("WScript.Shell")
WshShell.CurrentDirectory = "D:\00bau\Tool\Update_Phone_Process_FileCSV"
WshShell.Run """C:\Program Files\Python\python.exe"" ""D:\00bau\Tool\Update_Phone_Process_FileCSV\app.py""", 1, False
WScript.Sleep 2000
WshShell.Run "http://localhost:5000", 1, False
