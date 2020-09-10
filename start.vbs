Set WshShell = CreateObject("WScript.Shell")
WshShell.Run chr(34) & "start.bat" & Chr(34), 0 'start the .bat file and hide cmd window
Set WshShell = Nothing