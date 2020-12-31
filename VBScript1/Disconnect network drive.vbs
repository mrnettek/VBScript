
Call DisconnectND("m:",True,True)

' -----------------------------------
Private Function DisconnectND(sDrive, bForce, bPermanent)

On Error Resume Next

	Dim oWNet : Set oWNet = WScript.CreateObject("WScript.Network")

	oWNet.RemoveNetworkDrive sDrive, bForce, bPermanent

	If Err.Number <> 0 then
		msgbox "Error: " & Err.Number & VbCrLf & Err.Description
	Else
		msgbox "Disconnect network drive " & sDrive & VbCrLf & "Status = OK"
	End if

End Function
