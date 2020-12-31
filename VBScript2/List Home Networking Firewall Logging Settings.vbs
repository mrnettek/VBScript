' Description: Enumerates the configuration settings for Internet Connection Firewall logging.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & _
        "\root\microsoft\homenet")

Set colItems = objWMIService.ExecQuery _
    ("Select * from HNet_FirewallLoggingSettings")

For Each objItem in colItems
    Wscript.Echo "ID: " & objItem.ID
    Wscript.Echo "Log Connections: " & objItem.LogConnections
    Wscript.Echo "Log Dropped Packets: " & objItem.LogDroppedPackets
    Wscript.Echo "Max File Size: " & objItem.MaxFileSize
    Wscript.Echo "Path: " & objItem.Path
Next

