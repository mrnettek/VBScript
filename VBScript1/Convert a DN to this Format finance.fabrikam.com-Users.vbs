On Error Resume Next

Set objSysInfo = CreateObject("ADSystemInfo")

strUser = objSysInfo.UserName
Set objUser = GetObject("LDAP://" & strUser)

strDN = objUser.distinguishedName

arrDN = Split(strDN, ",")

For i = 0 to UBound(arrDN)
    intLength = Len(arrDN(i))
    intCounter = intLength - 3
    arrDN(i) = Right(arrDn(i), intCounter)
Next

For i = 1 to UBound(arrDN) - 1
    strNewName = strNewName & arrDN(i) & "."
Next

intLastItem = UBound(arrDN)
strNewName = strNewName & arrDN(intLastItem) & "/" & arrDN(0)

Wscript.Echo strNewName
  


