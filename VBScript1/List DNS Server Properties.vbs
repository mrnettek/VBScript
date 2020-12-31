' Description: Returns detailed information about a DNS server.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & _
        "\root\MicrosoftDNS")

Set colItems = objWMIService.ExecQuery("Select * from MicrosoftDNS_Server")

For Each objItem in colItems
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Address Answer Limit: " & objItem.AddressAnswerLimit
    Wscript.Echo "Allow Update: " & objItem.AllowUpdate
    Wscript.Echo "Autocache Update: " & objItem.AutoCacheUpdate
    Wscript.Echo "Autoconfig File Zones: " & objItem.AutoConfigFileZones
    Wscript.Echo "Bind Secondaries: " & objItem.BindSecondaries
    Wscript.Echo "Boot Method: " & objItem.BootMethod
    Wscript.Echo "Default Aging State: " & objItem.DefaultAgingState
    Wscript.Echo "Default No-Refresh Interval: " & _
        objItem.DefaultNoRefreshInterval
    Wscript.Echo "Default Refresh Interval: " & objItem.DefaultRefreshInterval
    Wscript.Echo "Disable AutoReverse Zones: " & _
        objItem.DisableAutoReverseZones
    Wscript.Echo "Disjoint Nets: " & objItem.DisjointNets
    Wscript.Echo "Directory Service Available: " & objItem.DsAvailable
    Wscript.Echo "Directory Service Polling Interval: " & _
        objItem.DsPollingInterval
    Wscript.Echo "Directory Service Tombstone Interval: " & _
        objItem.DsTombstoneInterval
    Wscript.Echo "EDNS Cache Timeout: " & objItem.EDnsCacheTimeout
    Wscript.Echo "Enable Directory Partitions: " & _
        objItem.EnableDirectoryPartitions
    Wscript.Echo "Enable DNSSec: " & objItem.EnableDnsSec
    Wscript.Echo "Enable EDNS Probes: " & objItem.EnableEDnsProbes
    Wscript.Echo "Event Log Level: " & objItem.EventLogLevel
    Wscript.Echo "Forward Delegations: " & objItem.ForwardDelegations
    Wscript.Echo "Forwarders: " & objItem.Forwarders
    Wscript.Echo "Forwarding Timeout: " & objItem.ForwardingTimeout
    Wscript.Echo "Is Slave: " & objItem.IsSlave
    Wscript.Echo "Listen Addresses: " & objItem.ListenAddresses
    Wscript.Echo "Local Net Priority: " & objItem.LocalNetPriority
    Wscript.Echo "Logfile Maximum Size: " & objItem.LogFileMaxSize
    Wscript.Echo "Logfile Path: " & objItem.LogFilePath
    Wscript.Echo "Log IPFilter List: " & objItem.LogIPFilterList
    Wscript.Echo "Log Level: " & objItem.LogLevel
    Wscript.Echo "Loose Wildcarding: " & objItem.LooseWildcarding
    Wscript.Echo "Maximum Cache Time-to-Live: " & objItem.MaxCacheTTL
    Wscript.Echo "Maximum Negative Cache Time-to-Live: " & _
        objItem.MaxNegativeCacheTTL
    Wscript.Echo "Name Check Flag: " & objItem.NameCheckFlag
    Wscript.Echo "No Recursion: " & objItem.NoRecursion
    Wscript.Echo "Recursion Retry: " & objItem.RecursionRetry
    Wscript.Echo "Recursion Timeout: " & objItem.RecursionTimeout
    Wscript.Echo "RoundRobin: " & objItem.RoundRobin
    Wscript.Echo "Rpc Protocol: " & objItem.RpcProtocol
    Wscript.Echo "Scavenging Interval: " & objItem.ScavengingInterval
    Wscript.Echo "Secure Responses: " & objItem.SecureResponses
    Wscript.Echo "Send Port: " & objItem.SendPort
    For Each strAddress in objItem.ServerAddresses
        Wscript.Echo "Server Address: " & strAddress
    Next
    Wscript.Echo "Started: " & objItem.Started
    Wscript.Echo "Start Mode: " & objItem.StartMode
    Wscript.Echo "Strict File Parsing: " & objItem.StrictFileParsing
    Wscript.Echo "Update Options: " & objItem.UpdateOptions
    Wscript.Echo "Version: " & objItem.Version
    Wscript.Echo "Write Authority NS: " & objItem.WriteAuthorityNS
    Wscript.Echo "Xfr Connect Timeout: " & objItem.XfrConnectTimeout
Next

