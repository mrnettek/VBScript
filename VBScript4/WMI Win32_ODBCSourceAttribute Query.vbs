On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_ODBCSourceAttribute",,48)
For Each objItem in colItems
    Wscript.Echo "Attribute: " & objItem.Attribute
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "DataSource: " & objItem.DataSource
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "SettingID: " & objItem.SettingID
    Wscript.Echo "Value: " & objItem.Value
Next

