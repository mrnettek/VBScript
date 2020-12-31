On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_ClassicCOMClassSetting",,48)
For Each objItem in colItems
    Wscript.Echo "AppID: " & objItem.AppID
    Wscript.Echo "AutoConvertToClsid: " & objItem.AutoConvertToClsid
    Wscript.Echo "AutoTreatAsClsid: " & objItem.AutoTreatAsClsid
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "ComponentId: " & objItem.ComponentId
    Wscript.Echo "Control: " & objItem.Control
    Wscript.Echo "DefaultIcon: " & objItem.DefaultIcon
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "InprocHandler: " & objItem.InprocHandler
    Wscript.Echo "InprocHandler32: " & objItem.InprocHandler32
    Wscript.Echo "InprocServer: " & objItem.InprocServer
    Wscript.Echo "InprocServer32: " & objItem.InprocServer32
    Wscript.Echo "Insertable: " & objItem.Insertable
    Wscript.Echo "JavaClass: " & objItem.JavaClass
    Wscript.Echo "LocalServer: " & objItem.LocalServer
    Wscript.Echo "LocalServer32: " & objItem.LocalServer32
    Wscript.Echo "LongDisplayName: " & objItem.LongDisplayName
    Wscript.Echo "ProgId: " & objItem.ProgId
    Wscript.Echo "SettingID: " & objItem.SettingID
    Wscript.Echo "ShortDisplayName: " & objItem.ShortDisplayName
    Wscript.Echo "ThreadingModel: " & objItem.ThreadingModel
    Wscript.Echo "ToolBoxBitmap32: " & objItem.ToolBoxBitmap32
    Wscript.Echo "TreatAsClsid: " & objItem.TreatAsClsid
    Wscript.Echo "TypeLibraryId: " & objItem.TypeLibraryId
    Wscript.Echo "Version: " & objItem.Version
    Wscript.Echo "VersionIndependentProgId: " & objItem.VersionIndependentProgId
Next

