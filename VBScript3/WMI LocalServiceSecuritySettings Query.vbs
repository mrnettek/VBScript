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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM LocalServiceSecuritySettings", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "DetectReplays: " & objItem.DetectReplays
      WScript.Echo "InactivityTimeout: " & WMIDateStringToDate(objItem.InactivityTimeout)
      WScript.Echo "IssuedCookieLifetime: " & WMIDateStringToDate(objItem.IssuedCookieLifetime)
      WScript.Echo "MaxCachedCookies: " & objItem.MaxCachedCookies
      WScript.Echo "MaxClockSkew: " & WMIDateStringToDate(objItem.MaxClockSkew)
      WScript.Echo "MaxPendingSessions: " & objItem.MaxPendingSessions
      WScript.Echo "MaxStatefulNegotiations: " & objItem.MaxStatefulNegotiations
      WScript.Echo "NegotiationTimeout: " & WMIDateStringToDate(objItem.NegotiationTimeout)
      WScript.Echo "ReconnectTransportOnFailure: " & objItem.ReconnectTransportOnFailure
      WScript.Echo "ReplayCacheSize: " & objItem.ReplayCacheSize
      WScript.Echo "ReplayWindow: " & WMIDateStringToDate(objItem.ReplayWindow)
      WScript.Echo "SessionKeyRenewalInterval: " & WMIDateStringToDate(objItem.SessionKeyRenewalInterval)
      WScript.Echo "SessionKeyRolloverInterval: " & WMIDateStringToDate(objItem.SessionKeyRolloverInterval)
      WScript.Echo "TimestampValidityDuration: " & WMIDateStringToDate(objItem.TimestampValidityDuration)
      WScript.Echo
   Next
Next


Function WMIDateStringToDate(dtmDate)
WScript.Echo dtm: 
	WMIDateStringToDate = CDate(Mid(dtmDate, 5, 2) & "/" & _
	Mid(dtmDate, 7, 2) & "/" & Left(dtmDate, 4) _
	& " " & Mid (dtmDate, 9, 2) & ":" & Mid(dtmDate, 11, 2) & ":" & Mid(dtmDate,13, 2))
End Function

