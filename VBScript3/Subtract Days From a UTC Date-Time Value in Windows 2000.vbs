strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set objOS = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
 
For Each strOS in objOS
    dtmInstallDate = strOS.InstallDate
    dtmDate = CDate(Mid(dtmInstallDate, 5, 2) & "/" & _
        Mid(dtmInstallDate, 7, 2) & "/" & Left(dtmInstallDate, 4))
    Wscript.Echo dtmDate - 30
Next
  


