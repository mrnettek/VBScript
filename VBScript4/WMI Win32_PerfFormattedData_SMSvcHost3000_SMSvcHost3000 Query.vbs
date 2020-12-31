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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_PerfFormattedData_SMSvcHost3000_SMSvcHost3000", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Caption: " & objItem.Caption
      WScript.Echo "ConnectionsAcceptedovernetpipe: " & objItem.ConnectionsAcceptedovernetpipe
      WScript.Echo "ConnectionsAcceptedovernettcp: " & objItem.ConnectionsAcceptedovernettcp
      WScript.Echo "ConnectionsDispatchedovernetpipe: " & objItem.ConnectionsDispatchedovernetpipe
      WScript.Echo "ConnectionsDispatchedovernettcp: " & objItem.ConnectionsDispatchedovernettcp
      WScript.Echo "Description: " & objItem.Description
      WScript.Echo "DispatchFailuresovernetpipe: " & objItem.DispatchFailuresovernetpipe
      WScript.Echo "DispatchFailuresovernettcp: " & objItem.DispatchFailuresovernettcp
      WScript.Echo "Frequency_Object: " & objItem.Frequency_Object
      WScript.Echo "Frequency_PerfTime: " & objItem.Frequency_PerfTime
      WScript.Echo "Frequency_Sys100NS: " & objItem.Frequency_Sys100NS
      WScript.Echo "Name: " & objItem.Name
      WScript.Echo "ProtocolFailuresovernetpipe: " & objItem.ProtocolFailuresovernetpipe
      WScript.Echo "ProtocolFailuresovernettcp: " & objItem.ProtocolFailuresovernettcp
      WScript.Echo "RegistrationsActivefornetpipe: " & objItem.RegistrationsActivefornetpipe
      WScript.Echo "RegistrationsActivefornettcp: " & objItem.RegistrationsActivefornettcp
      WScript.Echo "Timestamp_Object: " & objItem.Timestamp_Object
      WScript.Echo "Timestamp_PerfTime: " & objItem.Timestamp_PerfTime
      WScript.Echo "Timestamp_Sys100NS: " & objItem.Timestamp_Sys100NS
      WScript.Echo "UrisRegisteredfornetpipe: " & objItem.UrisRegisteredfornetpipe
      WScript.Echo "UrisRegisteredfornettcp: " & objItem.UrisRegisteredfornettcp
      WScript.Echo "UrisUnregisteredfornetpipe: " & objItem.UrisUnregisteredfornetpipe
      WScript.Echo "UrisUnregisteredfornettcp: " & objItem.UrisUnregisteredfornettcp
      WScript.Echo
   Next
Next

