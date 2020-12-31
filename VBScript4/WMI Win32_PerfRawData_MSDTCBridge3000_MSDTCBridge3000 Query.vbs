On Error Resume Next

Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

arrComputers = Array(".")
For Each strComputer In arrComputers
   WScript.Echo
   WScript.Echo "=========================================="
   WScript.Echo "Computer: " & strComputer
   WScript.Echo "=========================================="

   Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
   Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_PerfRawData_MSDTCBridge3000_MSDTCBridge3000", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Averageparticipantcommitresponsetime: " & objItem.Averageparticipantcommitresponsetime
      WScript.Echo "Averageparticipantcommitresponsetime_Base: " & objItem.Averageparticipantcommitresponsetime_Base
      WScript.Echo "Averageparticipantprepareresponsetime: " & objItem.Averageparticipantprepareresponsetime
      WScript.Echo "Averageparticipantprepareresponsetime_Base: " & objItem.Averageparticipantprepareresponsetime_Base
      WScript.Echo "Caption: " & objItem.Caption
      WScript.Echo "CommitretrycountPersec: " & objItem.CommitretrycountPersec
      WScript.Echo "Description: " & objItem.Description
      WScript.Echo "FaultsreceivedcountPersec: " & objItem.FaultsreceivedcountPersec
      WScript.Echo "FaultssentcountPersec: " & objItem.FaultssentcountPersec
      WScript.Echo "Frequency_Object: " & objItem.Frequency_Object
      WScript.Echo "Frequency_PerfTime: " & objItem.Frequency_PerfTime
      WScript.Echo "Frequency_Sys100NS: " & objItem.Frequency_Sys100NS
      WScript.Echo "MessagesendfailuresPersec: " & objItem.MessagesendfailuresPersec
      WScript.Echo "Name: " & objItem.Name
      WScript.Echo "PreparedretrycountPersec: " & objItem.PreparedretrycountPersec
      WScript.Echo "PrepareretrycountPersec: " & objItem.PrepareretrycountPersec
      WScript.Echo "ReplayretrycountPersec: " & objItem.ReplayretrycountPersec
      WScript.Echo "Timestamp_Object: " & objItem.Timestamp_Object
      WScript.Echo "Timestamp_PerfTime: " & objItem.Timestamp_PerfTime
      WScript.Echo "Timestamp_Sys100NS: " & objItem.Timestamp_Sys100NS
      WScript.Echo
   Next
Next

