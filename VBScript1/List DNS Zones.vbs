' Description: Returns information about all the DNS zones on a DNS server.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & _
        "\root\MicrosoftDNS")

Set colItems = objWMIService.ExecQuery("Select * from MicrosoftDNS_Zone")

For Each objItem in colItems
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Zone Type: " & objItem.ZoneType
    Wscript.Echo "Aging: " & objItem.Aging
    Wscript.Echo "Allow Update: " & objItem.AllowUpdate
    Wscript.Echo "Autocreated: " & objItem.AutoCreated
    Wscript.Echo "Available For Scavenge Time: " & _
        objItem.AvailForScavengeTime
    Wscript.Echo "Container Name: " & objItem.ContainerName
    Wscript.Echo "Data File: " & objItem.DataFile
    Wscript.Echo "Disable WINS Record Replication: " & _
        objItem.DisableWINSRecordReplication
    Wscript.Echo "DNS Server Name: " & objItem.DnsServerName
    Wscript.Echo "Directory Service-Integrated: " & objItem.DsIntegrated
    Wscript.Echo "Forwarder Slave: " & objItem.ForwarderSlave
    Wscript.Echo "Forwarder Timeout: " & objItem.ForwarderTimeout
    Wscript.Echo "Last Successful SOA Check: " & _
        objItem.LastSuccessfulSoaCheck
    Wscript.Echo "LastSuccessful Xfr: " & objItem.LastSuccessfulXfr
    Wscript.Echo "Local Master Servers: " & objItem.LocalMasterServers
    Wscript.Echo "Master Servers: " & objItem.MasterServers
    Wscript.Echo "No-Refresh Interval: " & objItem.NoRefreshInterval
    Wscript.Echo "Notify: " & objItem.Notify
    Wscript.Echo "Notify Servers: " & objItem.NotifyServers
    Wscript.Echo "Paused: " & objItem.Paused
    Wscript.Echo "Refresh Interval: " & objItem.RefreshInterval
    Wscript.Echo "Reverse: " & objItem.Reverse
    Wscript.Echo "Scavenge Servers: " & objItem.ScavengeServers
    Wscript.Echo "Secondary Servers: " & objItem.SecondaryServers
    Wscript.Echo "Secure Secondaries: " & objItem.SecureSecondaries
    Wscript.Echo "Shutdown: " & objItem.Shutdown
    Wscript.Echo "Use Wins: " & objItem.UseWins
    Wscript.Echo
Next

