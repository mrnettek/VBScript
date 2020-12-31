strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * From Win32_Environment Where Name = 'Path'")

For Each objItem in colItems
    strPath = objItem.VariableValue & ";C:\Scripts\"
    objItem.VariableValue = strPath
    objItem.Put_
Next
  


