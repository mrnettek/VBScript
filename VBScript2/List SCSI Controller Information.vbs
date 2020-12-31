' Description: Returns information about all the SCSI controllers found on a computer.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_SCSIController")

For Each objItem in colItems
    Wscript.Echo "Availability: " & objItem.Availability
    Wscript.Echo "Configuration Manager Error Code: " & _
         objItem.ConfigManagerErrorCode
    Wscript.Echo "Configuration Manager User Configuration: " & _
        objItem.ConfigManagerUserConfig
    Wscript.Echo "Device ID: " & objItem.DeviceID
    Wscript.Echo "Driver Name: " & objItem.DriverName
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "PNP Device ID: " & objItem.PNPDeviceID
    Wscript.Echo "Protocol Supported: " & objItem.ProtocolSupported
    Wscript.Echo "Status Information: " & objItem.StatusInfo
    Wscript.Echo
Next

