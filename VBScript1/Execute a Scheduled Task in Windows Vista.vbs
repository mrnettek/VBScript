Set objTaskService = CreateObject("Schedule.Service")
objTaskService.Connect

Set objRootFolder = objTaskService.GetFolder("\")
Set objTask = objRootFolder.GetTask("Test Task")

objTask.Run vbNull
  


