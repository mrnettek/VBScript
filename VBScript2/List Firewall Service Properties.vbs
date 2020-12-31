' Description: Lists service properties for the Windows Firewall current profile.


Set objFirewall = CreateObject("HNetCfg.FwMgr")
Set objPolicy = objFirewall.LocalPolicy.CurrentProfile

Set colServices = objPolicy.Services

For Each objService in colServices
    Wscript.Echo "Service name: " & objService.Name
    Wscript.Echo "Service enabled: " & objService.Enabled
    Wscript.Echo "Service type: " & objService.Type
    Wscript.Echo "Service IP version: " & objService.IPVersion
    Wscript.Echo "Service scope: " & objService.Scope
    Wscript.Echo "Service remote addresses: " & objService.RemoteAddresses
    Wscript.Echo "Service customized: " & objService.Customized
    Set colPorts = objService.GloballyOpenPorts
    For Each objPort in colPorts
        Wscript.Echo "Port name: " & objPort.Name
        Wscript.Echo "Port number: " & objPort.Port
        Wscript.Echo "Port enabled: " & objPort.Enabled
        Wscript.Echo "Port built-in: " & objPort.BuiltIn
        Wscript.Echo "Port IP version: " & objPort.IPVersion
        Wscript.Echo "Port protocol: " & objPort.Protocol
        Wscript.Echo "Port remote addresses: " & objPort.RemoteAddresses
        Wscript.Echo "Port scope: " & objPort.Scope
    Next
    Wscript.Echo
Next

