' Description: Displays IIS filter setting information.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsFilterSetting")
 
For Each objItem in colItems
    Wscript.Echo "Filter Description: " & objItem.FilterDescription
    Wscript.Echo "Filter Enable Cache: " & objItem.FilterEnableCache
    Wscript.Echo "Filter Enabled: " & objItem.FilterEnabled
    Wscript.Echo "Filter Flags: " & objItem.FilterFlags
    Wscript.Echo "Filter Path: " & objItem.FilterPath
    Wscript.Echo "Filter State: " & objItem.FilterState
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Notify Access Denied: " &  _
        objItem.NotifyAccessDenied
    Wscript.Echo "Notify Authentication Complete: " &  _
        objItem.NotifyAuthComplete
    Wscript.Echo "Notify Authentication: " &  _
        objItem.NotifyAuthentication
    Wscript.Echo "Notify End Of Net Session: " &  _
        objItem.NotifyEndOfNetSession
    Wscript.Echo "Notify End Of Request: " &  _
        objItem.NotifyEndOfRequest
    Wscript.Echo "Notify Log: " & objItem.NotifyLog
    Wscript.Echo "Notify Non-Secure Port: " &  _
        objItem.NotifyNonSecurePort
    Wscript.Echo "Notify Order High: " & objItem.NotifyOrderHigh
    Wscript.Echo "Notify Order Low: " & objItem.NotifyOrderLow
    Wscript.Echo "Notify Order Medium: " & objItem.NotifyOrderMedium
    Wscript.Echo "Notify Pre-Proc Headers: " &  _
        objItem.NotifyPreProcHeaders
    Wscript.Echo "Notify Read Raw Data: " &  _
        objItem.NotifyReadRawData
    Wscript.Echo "Notify Secure Port: " & objItem.NotifySecurePort
    Wscript.Echo "Notify Send Raw Data: " &  _
        objItem.NotifySendRawData
    Wscript.Echo "Notify Send Response: " &  _
        objItem.NotifySendResponse
    Wscript.Echo "Notify Url Map: " & objItem.NotifyUrlMap
    Wscript.Echo "Setting ID: " & objItem.SettingID
    Wscript.Echo "Win32 Error: " & objItem.Win32Error
Next

