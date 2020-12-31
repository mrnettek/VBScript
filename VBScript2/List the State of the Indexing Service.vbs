' Description: Returns information about the current state of the Indexing Service on the local computer.


On Error Resume Next

Set objAdminIS = CreateObject("Microsoft.ISAdm")
Wscript.Echo "Is running: " & objAdminIS.IsRunning
Wscript.Echo "Is paused: " & objAdminIS.IsPaused
Wscript.Echo "Computer name: " & objAdminIS.MachineName

