Set objShell = CreateObject("WScript.Shell")
objShell.CurrentDirectory = "C:\Users\steven.lee\新增資料夾\business-card-crm"
objShell.Run """C:\Program Files\nodejs\node.exe"" server.js", 0, False
