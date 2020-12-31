' Description: Lists configuration settings for Virtual Server.


On Error Resume Next

Set objVS = CreateObject("VirtualServer.Application")

Wscript.Echo "Available system capacity: " & objVS.AvailableSystemCapacity
Wscript.Echo "Default additions path: " & objVS.DefaultAdditionsPath
Wscript.Echo "Default VM configuration path: " & _
    objVS.DefaultVMConfigurationPath
Wscript.Echo "Default VN configuration path: " & _
    objVS.DefaultVNConfigurationPath
Wscript.Echo "Disk image search paths: " & objVS.DiskImageSearchPaths
Wscript.Echo "Existing configuration paths: " & _
    objVS.ExistingConfigurationPaths
Wscript.Echo "Maximum floppy drives per VM: " & objVS.MaximumFloppyDrivesPerVM
Wscript.Echo "Maximum memory per VM: " & objVS.MaximumMemoryPerVM
Wscript.Echo "Maximum network adapters per VM: " & _
    objVS.MaximumNetworkAdaptersPerVM
Wscript.Echo "Maximum number of IDE buses: " & objVS.MaximumNumberOfIDEBuses
Wscript.Echo "Maximum nmber of SCSI controllers: " & _
    objVS.MaximumNumberOfSCSIControllers
Wscript.Echo "Maximum parallel ports per VM: " & _
    objVS.MaximumParallelPortsPerVM
Wscript.Echo "Maximum serial ports per VM: " & objVS.MaximumSerialPortsPerVM
Wscript.Echo "Minimum memory per VM: " & objVS.MinimumMemoryPerVM
Wscript.Echo "Name: " & objVS.Name
Wscript.Echo "Product ID: " & objVS.ProductID
Wscript.Echo "Script search paths: " & objVS.ScriptSearchPaths
Wscript.Echo "Suggested maximum memory per VM: " & _
    objVS.SuggestedMaximumMemoryPerVM
Wscript.Echo "Uptime: " & objVS.Uptime
Wscript.Echo "Version: " & objVS.Version
Wscript.Echo "VMRC admin address: " & objVS.VMRCAdminAddress
Wscript.Echo "VMRC admin port number: " & objVS.VMRCAdminPortNumber
Wscript.Echo "VMRC authenticator: " & objVS.VMRCAuthenticator
Wscript.Echo "VMRC enabled: " & objVS.VMRCEnabled
Wscript.Echo "VMRC encryption certificate: " & objVS.VMRCEncryptionCertificate
Wscript.Echo "VMRC encryption certificate request: " & _
    objVS.VMRCEncryptionCertificateRequest
Wscript.Echo "VMRC encryption enabled: " & objVS.VMRCEncryptionEnabled
Wscript.Echo "VMRC idle connection timeout: " & objVS.VMRCIdleConnectionTimeout
Wscript.Echo "VMRC idle connection timeout enabled: " & _
    objVS.VMRCIdleConnectionTimeoutEnabled
Wscript.Echo "VMRC X-Resolution: " & objVS.VMRCXResolution
Wscript.Echo "VMRC Y-Resolution: " & objVS.VMRCYResolution

