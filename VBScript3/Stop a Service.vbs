
Msgbox StartService("Alerter")

' --------------------------------------- - Stop a Service -
Private Function StopService(strService)

	On Error Resume Next

	Dim oWshShell : Set oWshShell = CreateObject("WScript.Shell")

	Dim oService, iTimeOut, strComputername, tt
	strComputername = oWshShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
	iTimeOut 	= 60

	Set oService = GetObject("WinNT://" & strComputername & "/" & strService)
	
	If (Not oService.status = 1) Then
		oService.Stop
		wscript.sleep(250)

		For tt = 0 to iTimeOut
			If (oService.status = 1) Then
				Exit For
			Else
				wscript.sleep(1000)
    			End If
		Next
	End if

	If (IsObject(oService)) Then Set oService = Nothing
	If (IsObject(oWshShell)) Then Set oWshShell = Nothing

	If ( Err.Number <> 0 ) Then
		StopService = "ERROR: " & Err.Number & " " & Err.Description
		Err.Clear
	Else
		StopService = strService & " stopped successful"
	End If

End Function
