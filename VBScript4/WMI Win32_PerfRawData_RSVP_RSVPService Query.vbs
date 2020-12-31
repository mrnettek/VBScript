On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PerfRawData_RSVP_RSVPService",,48)
For Each objItem in colItems
    Wscript.Echo "BytesinQoSnotifications: " & objItem.BytesinQoSnotifications
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "FailedQoSrequests: " & objItem.FailedQoSrequests
    Wscript.Echo "FailedQoSsends: " & objItem.FailedQoSsends
    Wscript.Echo "Frequency_Object: " & objItem.Frequency_Object
    Wscript.Echo "Frequency_PerfTime: " & objItem.Frequency_PerfTime
    Wscript.Echo "Frequency_Sys100NS: " & objItem.Frequency_Sys100NS
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "NetworkInterfaces: " & objItem.NetworkInterfaces
    Wscript.Echo "Networksockets: " & objItem.Networksockets
    Wscript.Echo "QoSclients: " & objItem.QoSclients
    Wscript.Echo "QoSenabledreceivers: " & objItem.QoSenabledreceivers
    Wscript.Echo "QoSenabledsenders: " & objItem.QoSenabledsenders
    Wscript.Echo "QoSnotifications: " & objItem.QoSnotifications
    Wscript.Echo "RSVPsessions: " & objItem.RSVPsessions
    Wscript.Echo "Timers: " & objItem.Timers
    Wscript.Echo "Timestamp_Object: " & objItem.Timestamp_Object
    Wscript.Echo "Timestamp_PerfTime: " & objItem.Timestamp_PerfTime
    Wscript.Echo "Timestamp_Sys100NS: " & objItem.Timestamp_Sys100NS
Next

