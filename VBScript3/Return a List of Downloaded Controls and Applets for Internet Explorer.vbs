strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & _
    "\root\cimv2\Applications\MicrosoftIE")

Set colIESettings = objWMIService.ExecQuery _
    ("Select * from MicrosoftIE_Object")

For Each strIESetting in colIESettings
    Wscript.Echo "Code base: " & strIESetting.CodeBase
    Wscript.Echo "Program file: " & strIESetting.ProgramFile
    Wscript.Echo "Status: " & strIESetting.Status
    Wscript.Echo
Next
  


