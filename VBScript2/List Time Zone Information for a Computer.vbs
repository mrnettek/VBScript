' Description: Retrieve information about the time zone configured on a computer.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_TimeZone")

For Each objItem in colItems
    Wscript.Echo "Bias: " & objItem.Bias
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Daylight Bias: " & objItem.DaylightBias
    Wscript.Echo "Daylight Day: " & objItem.DaylightDay
    Wscript.Echo "Daylight Day of Week: " & objItem.DaylightDayOfWeek
    Wscript.Echo "Daylight Hour: " & objItem.DaylightHour
    Wscript.Echo "Daylight Millisecond: " & objItem.DaylightMillisecond
    Wscript.Echo "Daylight Minute: " & objItem.DaylightMinute
    Wscript.Echo "Daylight Month: " & objItem.DaylightMonth
    Wscript.Echo "Daylight Name: " & objItem.DaylightName
    Wscript.Echo "Daylight Second: " & objItem.DaylightSecond
    Wscript.Echo "Daylight Year: " & objItem.DaylightYear
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Setting ID: " & objItem.SettingID
    Wscript.Echo "Standard Bias: " & objItem.StandardBias
    Wscript.Echo "Standard Day: " & objItem.StandardDay
    Wscript.Echo "Standard Day of Week: " & objItem.StandardDayOfWeek
    Wscript.Echo "Standard Hour: " & objItem.StandardHour
    Wscript.Echo "Standard Millisecond: " & objItem.StandardMillisecond
    Wscript.Echo "Standard Minute: " & objItem.StandardMinute
    Wscript.Echo "Standard Month: " & objItem.StandardMonth
    Wscript.Echo "Standard Name: " & objItem.StandardName
    Wscript.Echo "Standard Second: " & objItem.StandardSecond
    Wscript.Echo "Standard Year: " & objItem.StandardYear
    Wscript.Echo
Next

