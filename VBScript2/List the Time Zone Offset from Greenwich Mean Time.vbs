' Description: Determines the difference (in minutes) between the time zone in use on the specified computer and Greenwich Mean Time. The time zone offset can be extremely useful in WMI scripts that need to work with date-time values.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colTimeZone = objWMIService.ExecQuery("Select * from Win32_TimeZone")
 
For Each objTimeZone in colTimeZone
    Wscript.Echo "Offset: "& objTimeZone.Bias 
Next

