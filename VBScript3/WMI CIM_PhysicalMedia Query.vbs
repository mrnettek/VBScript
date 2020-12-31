On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from CIM_PhysicalMedia",,48)
For Each objItem in colItems
    Wscript.Echo "Capacity: " & objItem.Capacity
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "CleanerMedia: " & objItem.CleanerMedia
    Wscript.Echo "CreationClassName: " & objItem.CreationClassName
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "HotSwappable: " & objItem.HotSwappable
    Wscript.Echo "InstallDate: " & objItem.InstallDate
    Wscript.Echo "Manufacturer: " & objItem.Manufacturer
    Wscript.Echo "MediaDescription: " & objItem.MediaDescription
    Wscript.Echo "MediaType: " & objItem.MediaType
    Wscript.Echo "Model: " & objItem.Model
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "OtherIdentifyingInfo: " & objItem.OtherIdentifyingInfo
    Wscript.Echo "PartNumber: " & objItem.PartNumber
    Wscript.Echo "PoweredOn: " & objItem.PoweredOn
    Wscript.Echo "Removable: " & objItem.Removable
    Wscript.Echo "Replaceable: " & objItem.Replaceable
    Wscript.Echo "SerialNumber: " & objItem.SerialNumber
    Wscript.Echo "SKU: " & objItem.SKU
    Wscript.Echo "Status: " & objItem.Status
    Wscript.Echo "Tag: " & objItem.Tag
    Wscript.Echo "Version: " & objItem.Version
    Wscript.Echo "WriteProtectOn: " & objItem.WriteProtectOn
Next

