strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\default")

Set objItem = objWMIService.Get("SystemRestore")
errResults = objItem.Enable("")
  


