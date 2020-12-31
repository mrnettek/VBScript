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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM AuthenticationSuccessAuditEvent", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "AccountName: " & objItem.AccountName
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
      WScript.Echo "NameToAuthenticate: " & objItem.NameToAuthenticate
      WScript.Echo "Occurrence: " & objItem.Occurrence
      WScript.Echo "ProcessID: " & objItem.ProcessID
      WScript.Echo "ProcessName: " & objItem.ProcessName
      WScript.Echo "RequestPath: " & objItem.RequestPath
      WScript.Echo "RequestThreadAccountName: " & objItem.RequestThreadAccountName
      WScript.Echo "RequestUrl: " & objItem.RequestUrl
      strSECURITY_DESCRIPTOR = Join(objItem.SECURITY_DESCRIPTOR, ",")
         WScript.Echo "SECURITY_DESCRIPTOR: " & strSECURITY_DESCRIPTOR
      WScript.Echo "SecurityDescriptor: " & objItem.SecurityDescriptor
      WScript.Echo "SequenceNumber: " & objItem.SequenceNumber
      WScript.Echo "TIME_CREATED: " & objItem.TIME_CREATED
      WScript.Echo "TrustLevel: " & objItem.TrustLevel
      WScript.Echo "UserAuthenticated: " & objItem.UserAuthenticated
      WScript.Echo "UserAuthenticationType: " & objItem.UserAuthenticationType
      WScript.Echo "UserHostAddress: " & objItem.UserHostAddress
      WScript.Echo "UserName: " & objItem.UserName
      WScript.Echo
   Next
Next

