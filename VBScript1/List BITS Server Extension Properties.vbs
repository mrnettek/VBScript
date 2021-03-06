' Description: Displays BITS server extension properties.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsWebVirtualDirSetting")

For Each objItem in colItems
    Wscript.Echo "BITS Host ID: " & objItem.BITSHostId
    Wscript.Echo "BITS Host ID Fallback Timeout: " & _
        objItem.BITSHostIdFallbackTimeout
    Wscript.Echo "BITS Maximum Upload Size: " & _
        objItem.BITSMaximumUploadSize
    Wscript.Echo "BITS Server Notification Type: " & _
        objItem.BITSServerNotificationType
    Wscript.Echo "BITS Server Notification URL: " & _
        objItem.BITSServerNotificationURL
    Wscript.Echo "BITS Session Timeout: " & objItem.BITSSessionTimeout
    Wscript.Echo "BITS Upload Enabled: " & objItem.BITSUploadEnabled
Next

