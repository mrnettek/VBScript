' Description: Deletes all the print jobs for a printer named HP QuietJet.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colInstalledPrinters =  objWMIService.ExecQuery _
    ("Select * from Win32_Printer Where Name = 'HP QuietJet'")

For Each objPrinter in colInstalledPrinters
    objPrinter.CancelAllJobs()
Next

