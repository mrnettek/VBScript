strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colFiles = objWMIService.ExecQuery _
    ("Select * from CIM_DataFile where Extension = 'mp3'")
For Each objFile in colFiles
    Wscript.Echo objFile.Name 
Next
  


