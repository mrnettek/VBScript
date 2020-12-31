On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_Property",,48)
For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "ProductCode: " & objItem.ProductCode
    Wscript.Echo "Property: " & objItem.Property
    Wscript.Echo "SettingID: " & objItem.SettingID
    Wscript.Echo "Value: " & objItem.Value
Next

