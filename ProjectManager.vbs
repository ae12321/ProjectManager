Dim objShell, strDir
Set objShell = CreateObject("WScript.Shell")
strDir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
objShell.Run "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File """ & strDir & "\ProjectManager.ps1""", 0, False
