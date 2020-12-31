On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_Binary",,48)
For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Data: " & objItem.Data
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "ProductCode: " & objItem.ProductCode
    Wscript.Echo "SettingID: " & objItem.SettingID
Next

