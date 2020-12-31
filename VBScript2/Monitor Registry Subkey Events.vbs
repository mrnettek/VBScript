' Description: Temporary event consumer that monitors the registry for any changes to HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion.


Set wmiServices = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\default")

Set wmiSink = WScript.CreateObject("WbemScripting.SWbemSink", "SINK_") 
 
wmiServices.ExecNotificationQueryAsync wmiSink, _ 
    "SELECT * FROM RegistryKeyChangeEvent WHERE Hive='HKEY_LOCAL_MACHINE' " & _
        "AND KeyPath='SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion'" 
 
WScript.Echo "Listening for Registry Change Events..." & vbCrLf 
 
While(1) 
    WScript.Sleep 1000 
Wend 
 
Sub SINK_OnObjectReady(wmiObject, wmiAsyncContext) 
    WScript.Echo "Received Registry Change Event" & vbCrLf & _ 
        wmiObject.GetObjectText_() 
End Sub

