On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_Patch",,48)
For Each objItem in colItems
    Wscript.Echo "Attributes: " & objItem.Attributes
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "File: " & objItem.File
    Wscript.Echo "PatchSize: " & objItem.PatchSize
    Wscript.Echo "ProductCode: " & objItem.ProductCode
    Wscript.Echo "Sequence: " & objItem.Sequence
    Wscript.Echo "SettingID: " & objItem.SettingID
Next

