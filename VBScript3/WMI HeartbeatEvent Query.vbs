On Error Resume Next

Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

arrComputers = Array(".")
For Each strComputer In arrComputers
   WScript.Echo
   WScript.Echo "=========================================="
   WScript.Echo "Computer: " & strComputer
   WScript.Echo "=========================================="

   Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\aspnet")
   Set colItems = objWMIService.ExecQuery("SELECT * FROM HeartbeatEvent", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "AccountName: " & objItem.AccountName
      WScript.Echo "AppdomainCount: " & objItem.AppdomainCount
      WScript.Echo "ApplicationDomain: " & objItem.ApplicationDomain
      WScript.Echo "ApplicationPath: " & objItem.ApplicationPath
      WScript.Echo "ApplicationVirtualPath: " & objItem.ApplicationVirtualPath
      WScript.Echo "CustomEventDetails: " & objItem.CustomEventDetails
      WScript.Echo "EventCode: " & objItem.EventCode
      WScript.Echo "EventDetailCode: " & objItem.EventDetailCode
      WScript.Echo "EventID: " & objItem.EventID
      WScript.Echo "EventMessage: " & objItem.EventMessage
      WScript.Echo "EventTime: " & objItem.EventTime
      WScript.Echo "MachineName: " & objItem.MachineName
      WScript.Echo "ManagedHeapSize: " & objItem.ManagedHeapSize
      WScript.Echo "Occurrence: " & objItem.Occurrence
      WScript.Echo "PeakWorkingSet: " & objItem.PeakWorkingSet
      WScript.Echo "ProcessID: " & objItem.ProcessID
      WScript.Echo "ProcessName: " & objItem.ProcessName
      WScript.Echo "ProcessStartTime: " & WMIDateStringToDate(objItem.ProcessStartTime)
      WScript.Echo "RequestsExecuting: " & objItem.RequestsExecuting
      WScript.Echo "RequestsQueued: " & objItem.RequestsQueued
      WScript.Echo "RequestsRejected: " & objItem.RequestsRejected
      strSECURITY_DESCRIPTOR = Join(objItem.SECURITY_DESCRIPTOR, ",")
         WScript.Echo "SECURITY_DESCRIPTOR: " & strSECURITY_DESCRIPTOR
      WScript.Echo "SecurityDescriptor: " & objItem.SecurityDescriptor
      WScript.Echo "SequenceNumber: " & objItem.SequenceNumber
      WScript.Echo "ThreadCount: " & objItem.ThreadCount
      WScript.Echo "TIME_CREATED: " & objItem.TIME_CREATED
      WScript.Echo "TrustLevel: " & objItem.TrustLevel
      WScript.Echo "WorkingSet: " & objItem.WorkingSet
      WScript.Echo
   Next
Next


Function WMIDateStringToDate(dtmDate)
WScript.Echo dtm: 
	WMIDateStringToDate = CDate(Mid(dtmDate, 5, 2) & "/" & _
	Mid(dtmDate, 7, 2) & "/" & Left(dtmDate, 4) _
	& " " & Mid (dtmDate, 9, 2) & ":" & Mid(dtmDate, 11, 2) & ":" & Mid(dtmDate,13, 2))
End Function

