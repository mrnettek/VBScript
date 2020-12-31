On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_TSSessionDirectory",,48)
For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "SessionDirectoryActive: " & objItem.SessionDirectoryActive
    Wscript.Echo "SessionDirectoryClusterName: " & objItem.SessionDirectoryClusterName
    Wscript.Echo "SessionDirectoryExposeServerIP: " & objItem.SessionDirectoryExposeServerIP
    Wscript.Echo "SessionDirectoryLocation: " & objItem.SessionDirectoryLocation
    Wscript.Echo "SettingID: " & objItem.SettingID
Next

