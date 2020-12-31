On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PrinterConfiguration",,48)
For Each objItem in colItems
    Wscript.Echo "BitsPerPel: " & objItem.BitsPerPel
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Collate: " & objItem.Collate
    Wscript.Echo "Color: " & objItem.Color
    Wscript.Echo "Copies: " & objItem.Copies
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "DeviceName: " & objItem.DeviceName
    Wscript.Echo "DisplayFlags: " & objItem.DisplayFlags
    Wscript.Echo "DisplayFrequency: " & objItem.DisplayFrequency
    Wscript.Echo "DitherType: " & objItem.DitherType
    Wscript.Echo "DriverVersion: " & objItem.DriverVersion
    Wscript.Echo "Duplex: " & objItem.Duplex
    Wscript.Echo "FormName: " & objItem.FormName
    Wscript.Echo "HorizontalResolution: " & objItem.HorizontalResolution
    Wscript.Echo "ICMIntent: " & objItem.ICMIntent
    Wscript.Echo "ICMMethod: " & objItem.ICMMethod
    Wscript.Echo "LogPixels: " & objItem.LogPixels
    Wscript.Echo "MediaType: " & objItem.MediaType
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Orientation: " & objItem.Orientation
    Wscript.Echo "PaperLength: " & objItem.PaperLength
    Wscript.Echo "PaperSize: " & objItem.PaperSize
    Wscript.Echo "PaperWidth: " & objItem.PaperWidth
    Wscript.Echo "PelsHeight: " & objItem.PelsHeight
    Wscript.Echo "PelsWidth: " & objItem.PelsWidth
    Wscript.Echo "PrintQuality: " & objItem.PrintQuality
    Wscript.Echo "Scale: " & objItem.Scale
    Wscript.Echo "SettingID: " & objItem.SettingID
    Wscript.Echo "SpecificationVersion: " & objItem.SpecificationVersion
    Wscript.Echo "TTOption: " & objItem.TTOption
    Wscript.Echo "VerticalResolution: " & objItem.VerticalResolution
    Wscript.Echo "XResolution: " & objItem.XResolution
    Wscript.Echo "YResolution: " & objItem.YResolution
Next

