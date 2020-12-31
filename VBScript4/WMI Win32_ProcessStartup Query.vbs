On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_ProcessStartup",,48)
For Each objItem in colItems
    Wscript.Echo "CreateFlags: " & objItem.CreateFlags
    Wscript.Echo "EnvironmentVariables: " & objItem.EnvironmentVariables
    Wscript.Echo "ErrorMode: " & objItem.ErrorMode
    Wscript.Echo "FillAttribute: " & objItem.FillAttribute
    Wscript.Echo "PriorityClass: " & objItem.PriorityClass
    Wscript.Echo "ShowWindow: " & objItem.ShowWindow
    Wscript.Echo "Title: " & objItem.Title
    Wscript.Echo "WinstationDesktop: " & objItem.WinstationDesktop
    Wscript.Echo "X: " & objItem.X
    Wscript.Echo "XCountChars: " & objItem.XCountChars
    Wscript.Echo "XSize: " & objItem.XSize
    Wscript.Echo "Y: " & objItem.Y
    Wscript.Echo "YCountChars: " & objItem.YCountChars
    Wscript.Echo "YSize: " & objItem.YSize
Next

