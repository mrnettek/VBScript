On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PatchPackage",,48)
For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "PatchID: " & objItem.PatchID
    Wscript.Echo "ProductCode: " & objItem.ProductCode
    Wscript.Echo "SettingID: " & objItem.SettingID
Next

