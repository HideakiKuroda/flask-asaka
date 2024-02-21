Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "python.exe app.py", 0
Set WshShell = Nothing