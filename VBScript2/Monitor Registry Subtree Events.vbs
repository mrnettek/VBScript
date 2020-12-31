' Description: Temporary event consumer that monitors the registry for any changes to the HKLM portion of the registry.


Set wmiServices = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\default")
 
Set wmiSink = WScript.CreateObject("WbemScripting.SWbemSink", "SINK_") 
 
wmiServices.ExecNotificationQueryAsync wmiSink, _ 
    "SELECT * FROM RegistryTreeChangeEvent WHERE Hive= " _
        & "'HKEY_LOCAL_MACHINE' AND RootPath=''" 
 
 
WScript.Echo "Listening for Registry Change Events..." & vbCrLf 
 
While(1) 
    WScript.Sleep 1000 
Wend 
 
Sub SINK_OnObjectReady(wmiObject, wmiAsyncContext) 
    WScript.Echo "Received Registry Change Event" & vbCrLf & _ 
        wmiObject.GetObjectText_() 
End Sub

