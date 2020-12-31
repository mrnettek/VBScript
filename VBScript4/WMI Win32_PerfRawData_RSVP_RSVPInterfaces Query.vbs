On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PerfRawData_RSVP_RSVPInterfaces",,48)
For Each objItem in colItems
    Wscript.Echo "BlockedRESVs: " & objItem.BlockedRESVs
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Frequency_Object: " & objItem.Frequency_Object
    Wscript.Echo "Frequency_PerfTime: " & objItem.Frequency_PerfTime
    Wscript.Echo "Frequency_Sys100NS: " & objItem.Frequency_Sys100NS
    Wscript.Echo "Generalfailures: " & objItem.Generalfailures
    Wscript.Echo "Maximumadmittedbandwidth: " & objItem.Maximumadmittedbandwidth
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Numberofactiveflows: " & objItem.Numberofactiveflows
    Wscript.Echo "Numberofincomingmessagesdropped: " & objItem.Numberofincomingmessagesdropped
    Wscript.Echo "Numberofoutgoingmessagesdropped: " & objItem.Numberofoutgoingmessagesdropped
    Wscript.Echo "PATHERRmessagesreceived: " & objItem.PATHERRmessagesreceived
    Wscript.Echo "PATHERRmessagessent: " & objItem.PATHERRmessagessent
    Wscript.Echo "PATHmessagesreceived: " & objItem.PATHmessagesreceived
    Wscript.Echo "PATHmessagessent: " & objItem.PATHmessagessent
    Wscript.Echo "PATHstateblocktimeouts: " & objItem.PATHstateblocktimeouts
    Wscript.Echo "PATHTEARmessagesreceived: " & objItem.PATHTEARmessagesreceived
    Wscript.Echo "PATHTEARmessagessent: " & objItem.PATHTEARmessagessent
    Wscript.Echo "Policycontrolfailures: " & objItem.Policycontrolfailures
    Wscript.Echo "ReceivemessageserrorsBigmessages: " & objItem.ReceivemessageserrorsBigmessages
    Wscript.Echo "ReceivemessageserrorsNomemory: " & objItem.ReceivemessageserrorsNomemory
    Wscript.Echo "Reservedbandwidth: " & objItem.Reservedbandwidth
    Wscript.Echo "Resourcecontrolfailures: " & objItem.Resourcecontrolfailures
    Wscript.Echo "RESVCONFIRMmessagesreceived: " & objItem.RESVCONFIRMmessagesreceived
    Wscript.Echo "RESVCONFIRMmessagessent: " & objItem.RESVCONFIRMmessagessent
    Wscript.Echo "RESVERRmessagesreceived: " & objItem.RESVERRmessagesreceived
    Wscript.Echo "RESVERRmessagessent: " & objItem.RESVERRmessagessent
    Wscript.Echo "RESVmessagesreceived: " & objItem.RESVmessagesreceived
    Wscript.Echo "RESVmessagessent: " & objItem.RESVmessagessent
    Wscript.Echo "RESVstateblocktimeouts: " & objItem.RESVstateblocktimeouts
    Wscript.Echo "RESVTEARmessagesreceived: " & objItem.RESVTEARmessagesreceived
    Wscript.Echo "RESVTEARmessagessent: " & objItem.RESVTEARmessagessent
    Wscript.Echo "SendmessageserrorsBigmessages: " & objItem.SendmessageserrorsBigmessages
    Wscript.Echo "SendmessageserrorsNomemory: " & objItem.SendmessageserrorsNomemory
    Wscript.Echo "Signalingbytesreceived: " & objItem.Signalingbytesreceived
    Wscript.Echo "Signalingbytessent: " & objItem.Signalingbytessent
    Wscript.Echo "Timestamp_Object: " & objItem.Timestamp_Object
    Wscript.Echo "Timestamp_PerfTime: " & objItem.Timestamp_PerfTime
    Wscript.Echo "Timestamp_Sys100NS: " & objItem.Timestamp_Sys100NS
Next

