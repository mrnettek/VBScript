On Error Resume Next
Set objADSysInfo = CreateObject("ADSystemInfo")
strUser = objADSysInfo.UserName
Set objUser = GetObject("LDAP://" & strUser)
For Each strGroup in objUser.memberOf
    Set objGroup = GetObject("LDAP://" & strGroup)
    Wscript.Echo objGroup.CN
Next
  


