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

Dim strComputer : strComputer = "myComputer"


If Ping(strComputer) then
	MsgBox strComputer & " is reachable!",64
Else
	MsgBox strComputer & " is not reachable!",16
End if



' --------------------------------
Private Function Ping(strComputer)

Dim oPing 	: Set oPing 	= GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery("select * from Win32_PingStatus where address = '" & strComputer & "'")
Dim bReachable 	: bReachable 	= True

	For Each oStatus in oPing
		If IsNull(oStatus.StatusCode) or oStatus.StatusCode <> 0 Then 
			bReachable = False
		End If
	Next

	Set oPing = Nothing
	Ping = bReachable

End Function
