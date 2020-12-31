strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colFiles = objWMIService.ExecQuery _
    ("Select * from CIM_DataFile Where Extension = 'mdb' OR Extension = 'ldb'")

For Each objFile in colFiles
    Wscript.Echo objFile.Name
Next
  


