strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colFiles = objWMIService.ExecQuery _
    ("Select * from CIM_DataFile Where Drive = 'C:' and Extension = 'csg'")

For Each objFile in colFiles
    objFile.Delete
Next
  


