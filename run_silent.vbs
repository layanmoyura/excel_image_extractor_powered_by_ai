Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "python """ & Replace(WScript.ScriptFullName, WScript.ScriptName, "") & "app.py""", 0, False
Set WshShell = Nothing
