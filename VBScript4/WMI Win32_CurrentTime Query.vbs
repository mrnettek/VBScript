On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_CurrentTime",,48)
For Each objItem in colItems
    Wscript.Echo "Day: " & objItem.Day
    Wscript.Echo "DayOfWeek: " & objItem.DayOfWeek
    Wscript.Echo "Hour: " & objItem.Hour
    Wscript.Echo "Milliseconds: " & objItem.Milliseconds
    Wscript.Echo "Minute: " & objItem.Minute
    Wscript.Echo "Month: " & objItem.Month
    Wscript.Echo "Quarter: " & objItem.Quarter
    Wscript.Echo "Second: " & objItem.Second
    Wscript.Echo "WeekInMonth: " & objItem.WeekInMonth
    Wscript.Echo "Year: " & objItem.Year
Next

