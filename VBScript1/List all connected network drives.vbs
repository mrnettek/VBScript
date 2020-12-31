Dim oWshNetwork : Set oWshNetwork = WScript.CreateObject("WScript.Network")
Dim sList
Dim oDrives : Set oDrives = oWshNetwork.EnumNetworkDrives

For iCount = 0 to oDrives.count - 1 step 2
	sList = sList & "Drive: " & oDrives.item(iCount)
	sList = sList & " Source: " & oDrives.item(iCount+1) & vbcr
Next

msgbox(sList)


