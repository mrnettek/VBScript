On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PhysicalMemory",,48)
For Each objItem in colItems
    Wscript.Echo "BankLabel: " & objItem.BankLabel
    Wscript.Echo "Capacity: " & objItem.Capacity
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "CreationClassName: " & objItem.CreationClassName
    Wscript.Echo "DataWidth: " & objItem.DataWidth
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "DeviceLocator: " & objItem.DeviceLocator
    Wscript.Echo "FormFactor: " & objItem.FormFactor
    Wscript.Echo "HotSwappable: " & objItem.HotSwappable
    Wscript.Echo "InstallDate: " & objItem.InstallDate
    Wscript.Echo "InterleaveDataDepth: " & objItem.InterleaveDataDepth
    Wscript.Echo "InterleavePosition: " & objItem.InterleavePosition
    Wscript.Echo "Manufacturer: " & objItem.Manufacturer
    Wscript.Echo "MemoryType: " & objItem.MemoryType
    Wscript.Echo "Model: " & objItem.Model
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "OtherIdentifyingInfo: " & objItem.OtherIdentifyingInfo
    Wscript.Echo "PartNumber: " & objItem.PartNumber
    Wscript.Echo "PositionInRow: " & objItem.PositionInRow
    Wscript.Echo "PoweredOn: " & objItem.PoweredOn
    Wscript.Echo "Removable: " & objItem.Removable
    Wscript.Echo "Replaceable: " & objItem.Replaceable
    Wscript.Echo "SerialNumber: " & objItem.SerialNumber
    Wscript.Echo "SKU: " & objItem.SKU
    Wscript.Echo "Speed: " & objItem.Speed
    Wscript.Echo "Status: " & objItem.Status
    Wscript.Echo "Tag: " & objItem.Tag
    Wscript.Echo "TotalWidth: " & objItem.TotalWidth
    Wscript.Echo "TypeDetail: " & objItem.TypeDetail
    Wscript.Echo "Version: " & objItem.Version
Next

