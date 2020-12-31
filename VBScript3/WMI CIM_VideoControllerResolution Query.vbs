On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from CIM_VideoControllerResolution",,48)
For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "HorizontalResolution: " & objItem.HorizontalResolution
    Wscript.Echo "MaxRefreshRate: " & objItem.MaxRefreshRate
    Wscript.Echo "MinRefreshRate: " & objItem.MinRefreshRate
    Wscript.Echo "NumberOfColors: " & objItem.NumberOfColors
    Wscript.Echo "RefreshRate: " & objItem.RefreshRate
    Wscript.Echo "ScanMode: " & objItem.ScanMode
    Wscript.Echo "SettingID: " & objItem.SettingID
    Wscript.Echo "VerticalResolution: " & objItem.VerticalResolution
Next

