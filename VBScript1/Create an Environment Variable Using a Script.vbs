strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set objVariable = objWMIService.Get("Win32_Environment").SpawnInstance_

objVariable.Name = "TestValue"
objVariable.UserName = "System"
objVariable.VariableValue = "This is a test"
objVariable.Put_
  


