' Description: Returns the name of the domain controller used to authenticate the logged-on user of a computer.


Set objDomain = GetObject("LDAP://rootDse")

objDC = objDomain.Get("dnsHostName")
Wscript.Echo "Authenticating domain controller:" & objDC

