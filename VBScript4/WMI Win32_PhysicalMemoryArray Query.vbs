On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PhysicalMemoryArray",,48)
For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "CreationClassName: " & objItem.CreationClassName
    Wscript.Echo "Depth: " & objItem.Depth
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Height: " & objItem.Height
    Wscript.Echo "HotSwappable: " & objItem.HotSwappable
    Wscript.Echo "InstallDate: " & objItem.InstallDate
    Wscript.Echo "Location: " & objItem.Location
    Wscript.Echo "Manufacturer: " & objItem.Manufacturer
    Wscript.Echo "MaxCapacity: " & objItem.MaxCapacity
    Wscript.Echo "MemoryDevices: " & objItem.MemoryDevices
    Wscript.Echo "MemoryErrorCorrection: " & objItem.MemoryErrorCorrection
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
    Wscript.Echo "Use: " & objItem.Use
    Wscript.Echo "Version: " & objItem.Version
    Wscript.Echo "Weight: " & objItem.Weight
    Wscript.Echo "Width: " & objItem.Width
Next

