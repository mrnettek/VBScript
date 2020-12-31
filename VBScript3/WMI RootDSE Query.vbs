On Error Resume Next

Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

arrComputers = Array(".")
For Each strComputer In arrComputers
   WScript.Echo
   WScript.Echo "=========================================="
   WScript.Echo "Computer: " & strComputer
   WScript.Echo "=========================================="

   Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\directory\LDAP")
   Set colItems = objWMIService.ExecQuery("SELECT * FROM RootDSE", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "configurationNamingContext: " & objItem.configurationNamingContext
      WScript.Echo "currentTime: " & objItem.currentTime
      WScript.Echo "defaultNamingContext: " & objItem.defaultNamingContext
      WScript.Echo "dnsHostName: " & objItem.dnsHostName
      WScript.Echo "dsServiceName: " & objItem.dsServiceName
      WScript.Echo "highestCommittedUSN: " & objItem.highestCommittedUSN
      WScript.Echo "LDAPServiceName: " & objItem.LDAPServiceName
      strnamingContexts = Join(objItem.namingContexts, ",")
         WScript.Echo "namingContexts: " & strnamingContexts
      WScript.Echo "rootDomainNamingContext: " & objItem.rootDomainNamingContext
      WScript.Echo "schemaNamingContext: " & objItem.schemaNamingContext
      WScript.Echo "serverName: " & objItem.serverName
      WScript.Echo "subschemaSubentry: " & objItem.subschemaSubentry
      WScript.Echo "supportedCapabilities: " & objItem.supportedCapabilities
      strsupportedControl = Join(objItem.supportedControl, ",")
         WScript.Echo "supportedControl: " & strsupportedControl
      strsupportedLDAPPolicies = Join(objItem.supportedLDAPPolicies, ",")
         WScript.Echo "supportedLDAPPolicies: " & strsupportedLDAPPolicies
      strsupportedLDAPVersion = Join(objItem.supportedLDAPVersion, ",")
         WScript.Echo "supportedLDAPVersion: " & strsupportedLDAPVersion
      strsupportedSASLMechanisms = Join(objItem.supportedSASLMechanisms, ",")
         WScript.Echo "supportedSASLMechanisms: " & strsupportedSASLMechanisms
      WScript.Echo
   Next
Next

