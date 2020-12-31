' Description: Retrieves and displays SNMP MIB-II (RFC 1213) System Group information from an SNMP managed node using the WMI SNMP Provider.


strTargetSnmpDevice = "192.168.0.1"
 
Set objWmiLocator = CreateObject("WbemScripting.SWbemLocator")
Set objWmiServices = objWmiLocator.ConnectServer("", "root\snmp\localhost")
 
Set objWmiNamedValueSet = CreateObject("WbemScripting.SWbemNamedValueSet")
objWmiNamedValueSet.Add "AgentAddress", strTargetSnmpDevice
objWmiNamedValueSet.Add "AgentReadCommunityName", "public"
 
Set colSystem = objWmiServices.InstancesOf("SNMP_RFC1213_MIB_system", , _
    objWmiNamedValueSet)
 
For Each objSystem In colSystem
    WScript.Echo "sysContact:  " & objSystem.sysContact  & vbCrLf & _
        "sysDescr:    " & objSystem.sysDescr    & vbCrLf & _
            "sysLocation: " & objSystem.sysLocation & vbCrLf & _
                "sysName:     " & objSystem.sysName     & vbCrLf & _
                    "sysObjectID: " & objSystem.sysObjectID & vbCrLf & _
                        "sysServices: " & objSystem.sysServices & vbCrLf & _
                            "sysUpTime:   " & objSystem.sysUpTime
Next

