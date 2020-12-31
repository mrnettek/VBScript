On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PerfFormattedData_Tcpip_IP",,48)
For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "DatagramsForwardedPersec: " & objItem.DatagramsForwardedPersec
    Wscript.Echo "DatagramsOutboundDiscarded: " & objItem.DatagramsOutboundDiscarded
    Wscript.Echo "DatagramsOutboundNoRoute: " & objItem.DatagramsOutboundNoRoute
    Wscript.Echo "DatagramsPersec: " & objItem.DatagramsPersec
    Wscript.Echo "DatagramsReceivedAddressErrors: " & objItem.DatagramsReceivedAddressErrors
    Wscript.Echo "DatagramsReceivedDeliveredPersec: " & objItem.DatagramsReceivedDeliveredPersec
    Wscript.Echo "DatagramsReceivedDiscarded: " & objItem.DatagramsReceivedDiscarded
    Wscript.Echo "DatagramsReceivedHeaderErrors: " & objItem.DatagramsReceivedHeaderErrors
    Wscript.Echo "DatagramsReceivedPersec: " & objItem.DatagramsReceivedPersec
    Wscript.Echo "DatagramsReceivedUnknownProtocol: " & objItem.DatagramsReceivedUnknownProtocol
    Wscript.Echo "DatagramsSentPersec: " & objItem.DatagramsSentPersec
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "FragmentationFailures: " & objItem.FragmentationFailures
    Wscript.Echo "FragmentedDatagramsPersec: " & objItem.FragmentedDatagramsPersec
    Wscript.Echo "FragmentReassemblyFailures: " & objItem.FragmentReassemblyFailures
    Wscript.Echo "FragmentsCreatedPersec: " & objItem.FragmentsCreatedPersec
    Wscript.Echo "FragmentsReassembledPersec: " & objItem.FragmentsReassembledPersec
    Wscript.Echo "FragmentsReceivedPersec: " & objItem.FragmentsReceivedPersec
    Wscript.Echo "Frequency_Object: " & objItem.Frequency_Object
    Wscript.Echo "Frequency_PerfTime: " & objItem.Frequency_PerfTime
    Wscript.Echo "Frequency_Sys100NS: " & objItem.Frequency_Sys100NS
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Timestamp_Object: " & objItem.Timestamp_Object
    Wscript.Echo "Timestamp_PerfTime: " & objItem.Timestamp_PerfTime
    Wscript.Echo "Timestamp_Sys100NS: " & objItem.Timestamp_Sys100NS
Next

