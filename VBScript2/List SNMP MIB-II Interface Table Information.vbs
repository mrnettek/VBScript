' Description: Retrieves and displays SNMP MIB-II (RFC 1213) Interface Table information from an SNMP managed node using the WMI SNMP Provider.


strTargetSnmpDevice = "192.168.0.1"
 
Set objWmiLocator = CreateObject("WbemScripting.SWbemLocator")
Set objWmiServices = objWmiLocator.ConnectServer("", "root\snmp\localhost")
 
Set objWmiNamedValueSet = CreateObject("WbemScripting.SWbemNamedValueSet")
objWmiNamedValueSet.Add "AgentAddress", strTargetSnmpDevice
objWmiNamedValueSet.Add "AgentReadCommunityName", "public"
 
Set colIfTable = objWmiServices.InstancesOf("SNMP_RFC1213_MIB_ifTable", , _
    objWmiNamedValueSet)
 
For Each objInterface In colIfTable
    WScript.Echo "ifIndex [Key]:        " & objInterface.ifIndex  & vbCrLf & _
        "   ifAdminStatus:     " & objInterface.ifAdminStatus     & vbCrLf & _
        "   ifDescr:           " & objInterface.ifDescr           & vbCrLf & _
        "   ifInDiscards:      " & objInterface.ifInDiscards      & vbCrLf & _
        "   ifInErrors:        " & objInterface.ifInErrors        & vbCrLf & _
        "   ifInNUcastPkts:    " & objInterface.ifInNUcastPkts    & vbCrLf & _
        "   ifInOctets:        " & objInterface.ifInOctets        & vbCrLf & _
        "   ifInUcastPkts:     " & objInterface.ifInUcastPkts     & vbCrLf & _
        "   ifInUnknownProtos: " & objInterface.ifInUnknownProtos & vbCrLf & _
        "   ifLastChange:      " & objInterface.ifLastChange      & vbCrLf & _
        "   ifMtu:             " & objInterface.ifMtu             & vbCrLf & _
        "   ifOperStatus:      " & objInterface.ifOperStatus      & vbCrLf & _
        "   ifOutDiscards:     " & objInterface.ifOutDiscards     & vbCrLf & _
        "   ifOutErrors:       " & objInterface.ifOutErrors       & vbCrLf & _
        "   ifOutNUcastPkts:   " & objInterface.ifOutNUcastPkts   & vbCrLf & _
        "   ifOutOctets:       " & objInterface.ifOutOctets       & vbCrLf & _
        "   ifOutQLen:         " & objInterface.ifOutQLen         & vbCrLf & _
        "   ifOutUcastPkts:    " & objInterface.ifOutUcastPkts    & vbCrLf & _
        "   ifPhysAddress:     " & objInterface.ifPhysAddress     & vbCrLf & _
        "   ifSpecific:        " & objInterface.ifSpecific        & vbCrLf & _
        "   ifSpeed:           " & objInterface.ifSpeed           & vbCrLf & _
        "   ifType:            " & objInterface.ifType            & vbCrLf
Next

