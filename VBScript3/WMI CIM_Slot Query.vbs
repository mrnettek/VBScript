On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from CIM_Slot",,48)
For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "ConnectorPinout: " & objItem.ConnectorPinout
    Wscript.Echo "ConnectorType: " & objItem.ConnectorType
    Wscript.Echo "CreationClassName: " & objItem.CreationClassName
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "HeightAllowed: " & objItem.HeightAllowed
    Wscript.Echo "InstallDate: " & objItem.InstallDate
    Wscript.Echo "LengthAllowed: " & objItem.LengthAllowed
    Wscript.Echo "Manufacturer: " & objItem.Manufacturer
    Wscript.Echo "MaxDataWidth: " & objItem.MaxDataWidth
    Wscript.Echo "Model: " & objItem.Model
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Number: " & objItem.Number
    Wscript.Echo "OtherIdentifyingInfo: " & objItem.OtherIdentifyingInfo
    Wscript.Echo "PartNumber: " & objItem.PartNumber
    Wscript.Echo "PoweredOn: " & objItem.PoweredOn
    Wscript.Echo "PurposeDescription: " & objItem.PurposeDescription
    Wscript.Echo "SerialNumber: " & objItem.SerialNumber
    Wscript.Echo "SKU: " & objItem.SKU
    Wscript.Echo "SpecialPurpose: " & objItem.SpecialPurpose
    Wscript.Echo "Status: " & objItem.Status
    Wscript.Echo "SupportsHotPlug: " & objItem.SupportsHotPlug
    Wscript.Echo "Tag: " & objItem.Tag
    Wscript.Echo "ThermalRating: " & objItem.ThermalRating
    Wscript.Echo "VccMixedVoltageSupport: " & objItem.VccMixedVoltageSupport
    Wscript.Echo "Version: " & objItem.Version
    Wscript.Echo "VppMixedVoltageSupport: " & objItem.VppMixedVoltageSupport
Next

