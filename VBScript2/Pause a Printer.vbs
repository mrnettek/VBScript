' Description: Pauses a printer named ArtDepartmentPrinter.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colInstalledPrinters =  objWMIService.ExecQuery _
    ("Select * from Win32_Printer Where Name = 'ArtDepartmentPrinter'")

For Each objPrinter in colInstalledPrinters 
    ObjPrinter.Pause()
Next

