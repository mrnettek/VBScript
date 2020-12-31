On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from CIM_PhysicalConnector",,48)
For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "ConnectorPinout: " & objItem.ConnectorPinout
    Wscript.Echo "ConnectorType: " & objItem.ConnectorType
    Wscript.Echo "CreationClassName: " & objItem.CreationClassName
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "InstallDate: " & objItem.InstallDate
    Wscript.Echo "Manufacturer: " & objItem.Manufacturer
    Wscript.Echo "Model: " & objItem.Model
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "OtherIdentifyingInfo: " & objItem.OtherIdentifyingInfo
    Wscript.Echo "PartNumber: " & objItem.PartNumber
    Wscript.Echo "PoweredOn: " & objItem.PoweredOn
    Wscript.Echo "SerialNumber: " & objItem.SerialNumber
    Wscript.Echo "SKU: " & objItem.SKU
    Wscript.Echo "Status: " & objItem.Status
    Wscript.Echo "Tag: " & objItem.Tag
    Wscript.Echo "Version: " & objItem.Version
Next

