On Error Resume Next

Set objSysInfo = CreateObject("ADSystemInfo")

strUser = objSysInfo.UserName
Set objUser = GetObject("LDAP://" & strUser)

Wscript.Echo objUser.AdsPath
  


