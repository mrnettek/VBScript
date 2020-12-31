' Description: Lists all current Virtual Server tasks.


On Error Resume Next

Set objVS = CreateObject("VirtualServer.Application")
set colTasks = objVS.Tasks

For Each objTask in colTasks
     Wscript.Echo "Description: " & objTask.Description
     Wscript.Echo "ID: " & objTask.ID
     Wscript.Echo "Is cancellable: " & objTask.IsCancelable
     Wscript.Echo "Is complete: " & objTask.IsComplete
     Wscript.Echo "Percent completed: " & objTask.PercentCompleted
     Wscript.Echo "Result: " & objTask.Result
     Wscript.Echo 
Next

