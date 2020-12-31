On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from CIM_PhysicalFrame",,48)
For Each objItem in colItems
    Wscript.Echo "AudibleAlarm: " & objItem.AudibleAlarm
    Wscript.Echo "BreachDescription: " & objItem.BreachDescription
    Wscript.Echo "CableManagementStrategy: " & objItem.CableManagementStrategy
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "CreationClassName: " & objItem.CreationClassName
    Wscript.Echo "Depth: " & objItem.Depth
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Height: " & objItem.Height
    Wscript.Echo "HotSwappable: " & objItem.HotSwappable
    Wscript.Echo "InstallDate: " & objItem.InstallDate
    Wscript.Echo "LockPresent: " & objItem.LockPresent
    Wscript.Echo "Manufacturer: " & objItem.Manufacturer
    Wscript.Echo "Model: " & objItem.Model
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "OtherIdentifyingInfo: " & objItem.OtherIdentifyingInfo
    Wscript.Echo "PartNumber: " & objItem.PartNumber
    Wscript.Echo "PoweredOn: " & objItem.PoweredOn
    Wscript.Echo "Removable: " & objItem.Removable
    Wscript.Echo "Replaceable: " & objItem.Replaceable
    Wscript.Echo "SecurityBreach: " & objItem.SecurityBreach
    Wscript.Echo "SerialNumber: " & objItem.SerialNumber
    Wscript.Echo "ServiceDescriptions: " & objItem.ServiceDescriptions
    Wscript.Echo "ServicePhilosophy: " & objItem.ServicePhilosophy
    Wscript.Echo "SKU: " & objItem.SKU
    Wscript.Echo "Status: " & objItem.Status
    Wscript.Echo "Tag: " & objItem.Tag
    Wscript.Echo "Version: " & objItem.Version
    Wscript.Echo "VisibleAlarm: " & objItem.VisibleAlarm
    Wscript.Echo "Weight: " & objItem.Weight
    Wscript.Echo "Width: " & objItem.Width
Next

