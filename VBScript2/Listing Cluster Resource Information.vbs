' Description: Returns cluster resource information.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\mscluster")

Set colItems = objWMIService.ExecQuery("Select * from MSCluster_Resource")

For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Characteristics: " & objItem.Characteristics
    Wscript.Echo "Core resource: " & objItem.CoreResource
    Wscript.Echo "Cryptographic checkpoints: " & objItem.CryptoCheckpoints
    Wscript.Echo "Debug prefix: " & objItem.DebugPrefix
    Wscript.Echo "Delete requires all nodes: " & _
        objItem.DeleteRequiresAllNodes
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Flags: " & objItem.Flags
    Wscript.Echo "Installation date: " & objItem.InstallDate
    Wscript.Echo "Is alive poll interval: " & objItem.IsAlivePollInterval
    Wscript.Echo "Load balance analysis interval: " & _
        objItem.LoadBalAnalysisInterval
    Wscript.Echo "Load balance minimum memory units: " & _
        objItem.LoadBalMinMemoryUnits
    Wscript.Echo "Load balance minimum processor units: " & _
        objItem.LoadBalMinProcessorUnits
    Wscript.Echo "Load balance sample interval: " & _
        objItem.LoadBalSampleInterval
    Wscript.Echo "Load balance startup interval: " & _
        objItem.LoadBalStartupInterval
    Wscript.Echo "Local quorum capable: " & objItem.LocalQuorumCapable
    Wscript.Echo "Looks alive poll interval: " & _
        objItem.LooksAlivePollInterval
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Pending timeout: " & objItem.PendingTimeout
    Wscript.Echo "Persistent state: " & objItem.PersistentState
    Wscript.Echo "Private properties: " & objItem.PrivateProperties
    Wscript.Echo "Quorum capable: " & objItem.QuorumCapable
    Wscript.Echo "Registry checkpoints: " & objItem.RegistryCheckpoints
    Wscript.Echo "Resource class: " & objItem.ResourceClass
    Wscript.Echo "Restart action: " & objItem.RestartAction
    Wscript.Echo "Restart period: " & objItem.RestartPeriod
    Wscript.Echo "Restart threshold: " & objItem.RestartThreshold
    Wscript.Echo "Retry period on failure: " & _
        objItem.RetryPeriodOnFailure
    Wscript.Echo "Separate monitor: " & objItem.SeparateMonitor
    Wscript.Echo "State: " & objItem.State
    Wscript.Echo "Status: " & objItem.Status
    Wscript.Echo "Subclass: " & objItem.Subclass
    Wscript.Echo "Type: " & objItem.Type
    Wscript.Echo
Next

