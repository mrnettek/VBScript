' Description: Returns configuration information for all the Terminal Service session directories found on a computer.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_TSSessionDirectory")

For Each objItem in colItems
    Wscript.Echo "Session Directory active: " & objItem.SessionDirectoryActive
    Wscript.Echo "Session Directory cluster name: " & _
        objItem.SessionDirectoryClusterName
    Wscript.Echo "Session Directory expose server IP address: " & _
        objItem.SessionDirectoryExposeServerIP
    Wscript.Echo "Session Directory IP address: " & _
        objItem.SessionDirectoryIPAddress
    Wscript.Echo "Session Directory location: " & _
        objItem.SessionDirectoryLocation
    Wscript.Echo "Setting ID: " & objItem.SettingID
    Wscript.Echo
Next

