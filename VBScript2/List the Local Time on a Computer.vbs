' Description: Returns information about the local time configured on a computer.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_LocalTime")

For Each objItem in colItems
    Wscript.Echo "Day: " & objItem.Day
    Wscript.Echo "Day of Week: " & objItem.DayOfWeek
    Wscript.Echo "Hour: " & objItem.Hour
    Wscript.Echo "Minute: " & objItem.Minute
    Wscript.Echo "Month: " & objItem.Month
    Wscript.Echo "Quarter: " & objItem.Quarter
    Wscript.Echo "Second: " & objItem.Second
    Wscript.Echo "Week in Month: " & objItem.WeekInMonth
    Wscript.Echo "Year: " & objItem.Year
    Wscript.Echo
Next

