Dim oWshShell		: Set oWshShell		= CreateObject("WScript.Shell")
Dim oEnv		: Set oEnv 		= oWshShell.Environment("System")
oEnv("Testvariable") = "myTest"

