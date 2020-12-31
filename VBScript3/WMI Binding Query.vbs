On Error Resume Next

Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

arrComputers = Array(".")
For Each strComputer In arrComputers
   WScript.Echo
   WScript.Echo "=========================================="
   WScript.Echo "Computer: " & strComputer
   WScript.Echo "=========================================="

   Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\ServiceModel")
   Set colItems = objWMIService.ExecQuery("SELECT * FROM Binding", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      strBindingElements = Join(objItem.BindingElements, ",")
         WScript.Echo "BindingElements: " & strBindingElements
      WScript.Echo "CloseTimeout: " & WMIDateStringToDate(objItem.CloseTimeout)
      WScript.Echo "Name: " & objItem.Name
      WScript.Echo "Namespace: " & objItem.Namespace
      WScript.Echo "OpenTimeout: " & WMIDateStringToDate(objItem.OpenTimeout)
      WScript.Echo "ReceiveTimeout: " & WMIDateStringToDate(objItem.ReceiveTimeout)
      WScript.Echo "Scheme: " & objItem.Scheme
      WScript.Echo "SendTimeout: " & WMIDateStringToDate(objItem.SendTimeout)
      WScript.Echo
   Next
Next


Function WMIDateStringToDate(dtmDate)
WScript.Echo dtm: 
	WMIDateStringToDate = CDate(Mid(dtmDate, 5, 2) & "/" & _
	Mid(dtmDate, 7, 2) & "/" & Left(dtmDate, 4) _
	& " " & Mid (dtmDate, 9, 2) & ":" & Mid(dtmDate, 11, 2) & ":" & Mid(dtmDate,13, 2))
End Function

