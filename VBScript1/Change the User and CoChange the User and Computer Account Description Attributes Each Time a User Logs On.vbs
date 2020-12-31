Set objSysInfo = CreateObject("ADSystemInfo")

Set objUser = GetObject("LDAP://" & objSysInfo.UserName)
Set objComputer = GetObject("LDAP://" & objSysInfo.ComputerName)

strMessage = objUser.CN & " logged on to " & objComputer.CN & " " & Now & "."

objUser.Description = strMessage
objUser.SetInfo

objComputer.Description = strMessage
objComputer.SetInfo
  


