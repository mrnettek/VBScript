strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colShares = objWMIService.ExecQuery _
    ("Select * from Win32_Share Where Name = 'C$'")

For Each objShare in colShares
    objShare.Delete
Next
  


