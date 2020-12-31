On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_DisplayConfiguration",,48)
For Each objItem in colItems
    Wscript.Echo "BitsPerPel: " & objItem.BitsPerPel
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "DeviceName: " & objItem.DeviceName
    Wscript.Echo "DisplayFlags: " & objItem.DisplayFlags
    Wscript.Echo "DisplayFrequency: " & objItem.DisplayFrequency
    Wscript.Echo "DitherType: " & objItem.DitherType
    Wscript.Echo "DriverVersion: " & objItem.DriverVersion
    Wscript.Echo "ICMIntent: " & objItem.ICMIntent
    Wscript.Echo "ICMMethod: " & objItem.ICMMethod
    Wscript.Echo "LogPixels: " & objItem.LogPixels
    Wscript.Echo "PelsHeight: " & objItem.PelsHeight
    Wscript.Echo "PelsWidth: " & objItem.PelsWidth
    Wscript.Echo "SettingID: " & objItem.SettingID
    Wscript.Echo "SpecificationVersion: " & objItem.SpecificationVersion
Next

