strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * From Win32_LocalTime")

For Each objItem in colItems
    Wscript.Echo "Month: " & objItem.Month
    Wscript.Echo "Day: " & objItem.Day
    Wscript.Echo "Year: " & objItem.Year
    Wscript.Echo "Hour: " & objItem.Hour
    Wscript.Echo "Minute: " & objItem.Minute
    Wscript.Echo "Second: " & objItem.Second
Next
  


