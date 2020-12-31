Dim oWshNetwork : Set oWshNetwork = WScript.CreateObject("WScript.Network")
Dim sText, iCount

Dim oPrinters : Set oPrinters = oWshNetwork.EnumPrinterConnections

sText = "Installed Printers: " & vbcr

For iCount = 0 to oPrinters.Count - 1 Step 2
	sText = sText & oPrinters.Item(iCount) & " " & oPrinters.Item(iCount+1) & vbcr
Next

msgbox(sText)
