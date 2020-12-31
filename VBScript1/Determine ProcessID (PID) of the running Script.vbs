Dim strComputer		: strComputer		= "."
Dim oWMIService 	: Set oWMIService 	= GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Dim colItems 		: Set colItems 		= oWMIService.ExecQuery("Select * From Win32_Process")
Dim sScriptCommandLine	: sScriptCommandLine	= wscript.ScriptFullName
Dim dStartTime		: dStartTime		= Now
Dim sScriptHostPath	: sScriptHostPath	= WScript.FullName


For Each oItem in colItems
	If InStr(oItem.CommandLine, sScriptCommandLine) > 0 and InStr(oItem.CommandLine, sScriptHostPath) > 0 Then
		'Wscript.Echo oItem.CommandLine
		dProcessStartTime = oItem.CreationDate
		dProcessStartTime = DateSerial(Left(dProcessStartTime, 4), Mid(dProcessStartTime,  5, 2), Mid(dProcessStartTime,  7, 2) ) + TimeSerial(Mid(dProcessStartTime,  9, 2), Mid(dProcessStartTime, 11, 2), Mid(dProcessStartTime, 13, 2))
		dStartDateDiff = DateDiff("s", dStartTime, dProcessStartTime)
		
		Select Case dStartDateDiff

			Case 0, -1, 1
				Wscript.Echo oItem.ProcessID
		End Select

	End If
Next

