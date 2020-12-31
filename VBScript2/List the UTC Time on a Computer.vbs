' Description: Reports the Universal Time Coordinate (UTC) time on a computer.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_UTCTime")

For Each objItem in colItems
    Wscript.Echo "Day: " & objItem.Day
    Wscript.Echo "Day of the Week: " & objItem.DayOfWeek
    Wscript.Echo "Hour: " & objItem.Hour
    Wscript.Echo "Milliseconds: " & objItem.Milliseconds
    Wscript.Echo "Minute: " & objItem.Minute
    Wscript.Echo "Month: " & objItem.Month
    Wscript.Echo "Quarter: " & objItem.Quarter
    Wscript.Echo "Second: " & objItem.Second
    Wscript.Echo "Week in the Month: " & objItem.WeekInMonth
    Wscript.Echo "Year: " & objItem.Year
Next

