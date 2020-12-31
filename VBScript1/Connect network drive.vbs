
Call ConnectND("m:","\\192.168.2.4\Docs",False,"","")

' -----------------------------------------------------------------------
Private Function ConnectND(sDrive, sTarget, bPermanent, sUser, sPassword)

On Error Resume Next

	Dim oWNet : Set oWNet = WScript.CreateObject("WScript.Network")

	If Len(sUser) > 0 then
		oWNet.MapNetworkDrive sDrive, sTarget, bPermanent, sUser, sPassword
	Else
		oWNet.MapNetworkDrive sDrive, sTarget, bPermanent
	End if

	If Err.Number <> 0 then
		msgbox "Error: " & Err.Number & VbCrLf & Err.Description
	Else
		msgbox "Connect network drive " & sDrive & VbCrLf & "Status = OK"
	End if

End Function
