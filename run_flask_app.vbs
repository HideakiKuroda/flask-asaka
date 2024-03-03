Set WshShell = CreateObject("WScript.Shell")
Set WshEnv = WshShell.Environment("PROCESS")
WshEnv("FLASK_APP_RUNNING") = "1"
WshShell.Run "python.exe app.py", 0, True
WshEnv.Remove("FLASK_APP_RUNNING")
Set WshShell = Nothing
