Option Explicit

Dim strComputer 	: strComputer = "."
Dim strProcessName 	: strProcessName = "notepad.exe"
Dim oWMIService 	: Set oWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

oWMIService.Security_.ImpersonationLevel = 3 
oWMIService.Security_.privileges.addasstring "SeDebugPrivilege", True

Dim colProcessList 	: Set colProcessList = oWMIService.ExecQuery("Select * from Win32_Process Where Name = " & Chr(34) & strProcessName & Chr(34))
Dim oProcess

		For Each oProcess in colProcessList
			oProcess.Terminate()
		Next

Set oWMIService		= Nothing
Set colProcessList 	= Nothing
