On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from CIM_UnitaryComputerSystem",,48)
For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "CreationClassName: " & objItem.CreationClassName
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "InitialLoadInfo: " & objItem.InitialLoadInfo
    Wscript.Echo "InstallDate: " & objItem.InstallDate
    Wscript.Echo "LastLoadInfo: " & objItem.LastLoadInfo
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "NameFormat: " & objItem.NameFormat
    Wscript.Echo "PowerManagementCapabilities: " & objItem.PowerManagementCapabilities
    Wscript.Echo "PowerManagementSupported: " & objItem.PowerManagementSupported
    Wscript.Echo "PowerState: " & objItem.PowerState
    Wscript.Echo "PrimaryOwnerContact: " & objItem.PrimaryOwnerContact
    Wscript.Echo "PrimaryOwnerName: " & objItem.PrimaryOwnerName
    Wscript.Echo "ResetCapability: " & objItem.ResetCapability
    Wscript.Echo "Roles: " & objItem.Roles
    Wscript.Echo "Status: " & objItem.Status
Next

