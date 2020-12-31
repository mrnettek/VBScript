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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM NamedPipeTransportBindingElement", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "ChannelInitializationTimeout: " & WMIDateStringToDate(objItem.ChannelInitializationTimeout)
      WScript.Echo "ConnectionBufferSize: " & objItem.ConnectionBufferSize
      WScript.Echo "ConnectionPoolSettings: " & objItem.ConnectionPoolSettings
      WScript.Echo "HostNameComparisonMode: " & objItem.HostNameComparisonMode
      WScript.Echo "ManualAddressing: " & objItem.ManualAddressing
      WScript.Echo "MaxBufferPoolSize: " & objItem.MaxBufferPoolSize
      WScript.Echo "MaxBufferSize: " & objItem.MaxBufferSize
      WScript.Echo "MaxOutputDelay: " & WMIDateStringToDate(objItem.MaxOutputDelay)
      WScript.Echo "MaxPendingAccepts: " & objItem.MaxPendingAccepts
      WScript.Echo "MaxPendingConnections: " & objItem.MaxPendingConnections
      WScript.Echo "MaxReceivedMessageSize: " & objItem.MaxReceivedMessageSize
      WScript.Echo "Scheme: " & objItem.Scheme
      WScript.Echo "TransferMode: " & objItem.TransferMode
      WScript.Echo
   Next
Next


Function WMIDateStringToDate(dtmDate)
WScript.Echo dtm: 
	WMIDateStringToDate = CDate(Mid(dtmDate, 5, 2) & "/" & _
	Mid(dtmDate, 7, 2) & "/" & Left(dtmDate, 4) _
	& " " & Mid (dtmDate, 9, 2) & ":" & Mid(dtmDate, 11, 2) & ":" & Mid(dtmDate,13, 2))
End Function

