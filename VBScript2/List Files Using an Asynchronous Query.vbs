' Description: Uses an asynchronous query to enumerate all the files on a computer. This is primarily a demonstration script; if actually run, it could take an hour or more to complete, depending on the number of files on the computer.


Const POPUP_DURATION = 120
Const OK_BUTTON = 0

Set objWSHShell = Wscript.CreateObject("Wscript.Shell")

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set objSink = WScript.CreateObject("WbemScripting.SWbemSink","SINK_")
objWMIService.ExecQueryAsync objSink, "Select * from CIM_DataFile"
objPopup = objWshShell.Popup("Starting file retrieval", _
    POPUP_DURATION, "File Retrieval", OK_BUTTON)

Sub SINK_OnObjectReady(objEvent, objAsyncContext)
    Wscript.Echo objEvent.Name
End Sub

