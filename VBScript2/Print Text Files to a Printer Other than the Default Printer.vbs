strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colPrinters =  objWMIService.ExecQuery _
    ("Select * from Win32_Printer Where Default = TRUE")

For Each objPrinter in colPrinters
    strOldDefault = objPrinter.Name
    strOldDefault = Replace(strOldDefault, "\", "\\")
Next

Set colPrinters =  objWMIService.ExecQuery _
    ("Select * from Win32_Printer Where Name = '\\\\atl-ps-01\\printer2'")

For Each objPrinter in colPrinters
    objPrinter.SetDefaultPrinter()
Next

Wscript.Sleep 2000

TargetFolder = "C:\Logs" 
Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace(TargetFolder) 
Set colItems = objFolder.Items
For Each objItem in colItems
    objItem.InvokeVerbEx("Print")
Next

Set colPrinters =  objWMIService.ExecQuery _
    ("Select * from Win32_Printer Where Name = '" & strOldDefault & "'")

For Each objPrinter in colPrinters
    objPrinter.SetDefaultPrinter()
Next
  


