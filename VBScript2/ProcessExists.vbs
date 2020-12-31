'#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#'
'/|									  |\\\\\\\\'
'//|									   |\\\\\\\'
'///|									    |\\\\\\'
'////|			Version 	1.0.0				     |\\\\\'
'/////|			Author:		Boris TOll 			      |\\\\'
'//////|		Last Update:	31.01.2008			       |\\\'
'///////|								        |\\'
'////////|									 |\'
'#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#'

Do while ProcessExists("notepad.exe")
	msgbox "Process is running"
Loop



' --------------------------------------------
Private Function ProcessExists(strProcessName)

Dim strComputer 	: strComputer = "."
Dim oWMIService 	: Set oWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Dim colProcessList 	: Set colProcessList = oWMIService.ExecQuery("Select * from Win32_Process Where Name = " & Chr(34) & strProcessName & Chr(34))
Dim iProcessTrue 	: iProcessTrue = 0

		For Each objProcess in colProcessList
			iProcessTrue = 1
			Exit For	
		Next


	Set oWMIService		= Nothing
	Set colProcessList 	= Nothing


	If iProcessTrue = 1 then
		ProcessExists = true
	Else
		ProcessExists = false
	End if

End Function
