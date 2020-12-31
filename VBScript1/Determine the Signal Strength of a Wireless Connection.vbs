strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\wmi")

Set colItems = objWMIService.ExecQuery("Select * From MSNdis_80211_ReceivedSignalStrength")

For Each objItem in colItems
    intStrength = objItem.NDIS80211ReceivedSignalStrength

    If intStrength > -57 Then
        strBars = "5 Bars"
    ElseIf intStrength > -68 Then
        strBars = "4 Bars"
    ElseIf intStrength > -72 Then
        strBars = "3 Bars"
    ElseIf intStrength > -80 Then
        strBars = "2 Bars"
    ElseIf intStrength > -90 Then
        strBars = "1 Bar"
    Else
        strBars = "Strength cannot be determined"
    End If
    
    Wscript.Echo objItem.InstanceName & " -- " & strBars
Next
  


