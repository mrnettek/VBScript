' Description: Temporary event consumer that monitors the registry for any changes to HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\CSDVersion.


Set wmiServices = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set wmiSink = WScript.CreateObject("WbemScripting.SWbemSink", "SINK_") 
 
wmiServices.ExecNotificationQueryAsync wmiSink, _ 
    "SELECT * FROM RegistryValueChangeEvent WHERE " & _
        "Hive='HKEY_LOCAL_MACHINE' AND " & _ 
            "KeyPath='SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion'" _
                & " AND ValueName='CSDVersion'" 
 
WScript.Echo "Listening for Registry Change Events..." & vbCrLf 
 
While(1) 
    WScript.Sleep 1000 
Wend 
 
Sub SINK_OnObjectReady(wmiObject, wmiAsyncContext) 
    WScript.Echo "Received Registry Change Event" & vbCrLf & _ 
        wmiObject.GetObjectText_() 
End Sub

