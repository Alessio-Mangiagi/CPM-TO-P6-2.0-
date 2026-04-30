Dim oShell, strDir
Set oShell = CreateObject("WScript.Shell")
strDir = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\"))
oShell.Run "pythonw """ & strDir & "main.py""", 0, False
Set oShell = Nothing
