' Description: Configures a printer so that documents can only be printed between 8:00 AM and 6:00 PM.


dtmStartTime= "********080000.000000+000"
dtmEndTime= "********180000.000000+000"

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colPrinters = objWMIService.ExecQuery _
    ("Select * From Win32_Printer Where DeviceID = 'ArtDepartmentPrinter' ")

For Each objPrinter in colPrinters
    objPrinter.StartTime = dtmStartTime
    objPrinter.UntilTime = dtmEndTime
    objPrinter.Put_
Next

