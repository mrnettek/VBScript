On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PingStatus",,48)
For Each objItem in colItems
    Wscript.Echo "Address: " & objItem.Address
    Wscript.Echo "BufferSize: " & objItem.BufferSize
    Wscript.Echo "NoFragmentation: " & objItem.NoFragmentation
    Wscript.Echo "PrimaryAddressResolutionStatus: " & objItem.PrimaryAddressResolutionStatus
    Wscript.Echo "ProtocolAddress: " & objItem.ProtocolAddress
    Wscript.Echo "ProtocolAddressResolved: " & objItem.ProtocolAddressResolved
    Wscript.Echo "RecordRoute: " & objItem.RecordRoute
    Wscript.Echo "ReplyInconsistency: " & objItem.ReplyInconsistency
    Wscript.Echo "ReplySize: " & objItem.ReplySize
    Wscript.Echo "ResolveAddressNames: " & objItem.ResolveAddressNames
    Wscript.Echo "ResponseTime: " & objItem.ResponseTime
    Wscript.Echo "ResponseTimeToLive: " & objItem.ResponseTimeToLive
    Wscript.Echo "RouteRecord: " & objItem.RouteRecord
    Wscript.Echo "RouteRecordResolved: " & objItem.RouteRecordResolved
    Wscript.Echo "SourceRoute: " & objItem.SourceRoute
    Wscript.Echo "SourceRouteType: " & objItem.SourceRouteType
    Wscript.Echo "StatusCode: " & objItem.StatusCode
    Wscript.Echo "Timeout: " & objItem.Timeout
    Wscript.Echo "TimeStampRecord: " & objItem.TimeStampRecord
    Wscript.Echo "TimeStampRecordAddress: " & objItem.TimeStampRecordAddress
    Wscript.Echo "TimeStampRecordAddressResolved: " & objItem.TimeStampRecordAddressResolved
    Wscript.Echo "TimestampRoute: " & objItem.TimestampRoute
    Wscript.Echo "TimeToLive: " & objItem.TimeToLive
    Wscript.Echo "TypeofService: " & objItem.TypeofService
Next

