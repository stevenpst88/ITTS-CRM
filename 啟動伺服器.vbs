Set objShell = CreateObject("WScript.Shell")
Set objFso = CreateObject("Scripting.FileSystemObject")
' Use the folder this script lives in (no hard-coded path; works after moving or on another PC)
objShell.CurrentDirectory = objFso.GetParentFolderName(WScript.ScriptFullName)
objShell.Run """C:\Program Files\nodejs\node.exe"" server.js", 0, False
