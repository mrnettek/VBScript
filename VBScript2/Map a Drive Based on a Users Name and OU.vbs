Set objSysInfo = CreateObject("ADSystemInfo")
strUser = objSysInfo.UserName
Set objUser = GetObject("LDAP://" & strUser)

strUserName = objUser.samAccountName
strOUPath = objUser.Parent

arrContainers = Split(strOUPath, ",")
arrOU = Split(arrContainers(0), "=")
strOU = arrOU(1)

strDrive = "\\Mission\Apps\Timesheets\" & strOU & "\" & strUserName
Set objNetwork = CreateObject("Wscript.Network")
objNetwork.MapNetworkDrive "K:", strDrive


