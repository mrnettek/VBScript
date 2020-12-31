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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM Service", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      strBaseAddresses = Join(objItem.BaseAddresses, ",")
         WScript.Echo "BaseAddresses: " & strBaseAddresses
      strBehaviors = Join(objItem.Behaviors, ",")
         WScript.Echo "Behaviors: " & strBehaviors
      WScript.Echo "ConfigurationName: " & objItem.ConfigurationName
      WScript.Echo "CounterInstanceName: " & objItem.CounterInstanceName
      WScript.Echo "DistinguishedName: " & objItem.DistinguishedName
      strExtensions = Join(objItem.Extensions, ",")
         WScript.Echo "Extensions: " & strExtensions
      strMetadata = Join(objItem.Metadata, ",")
         WScript.Echo "Metadata: " & strMetadata
      WScript.Echo "Name: " & objItem.Name
      WScript.Echo "Namespace: " & objItem.Namespace
      WScript.Echo "Opened: " & WMIDateStringToDate(objItem.Opened)
      strOutgoingChannels = Join(objItem.OutgoingChannels, ",")
         WScript.Echo "OutgoingChannels: " & strOutgoingChannels
      WScript.Echo "ProcessId: " & objItem.ProcessId
      WScript.Echo
   Next
Next


Function WMIDateStringToDate(dtmDate)
WScript.Echo dtm: 
	WMIDateStringToDate = CDate(Mid(dtmDate, 5, 2) & "/" & _
	Mid(dtmDate, 7, 2) & "/" & Left(dtmDate, 4) _
	& " " & Mid (dtmDate, 9, 2) & ":" & Mid(dtmDate, 11, 2) & ":" & Mid(dtmDate,13, 2))
End Function

