Dim oWshShell : Set oWshShell = WScript.CreateObject("WScript.Shell")

oWshShell.Run "calc"
WScript.Sleep 3000
oWshShell.AppActivate "Rechner"		'-> German MUI
oWshShell.AppActivate "Calculator"	'-> English MUI


