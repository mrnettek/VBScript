strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root")

Set objItem = objWMIService.Get("__Namespace")
Set objNamespace = objItem.SpawnInstance_

objNamespace.Name = "ScriptCenter"
objNamespace.Put_
  


