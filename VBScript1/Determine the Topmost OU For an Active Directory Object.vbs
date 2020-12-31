Set objSysInfo = CreateObject("ADSystemInfo")
strDN = objSysInfo.UserName

arrDN = Split(strDN, ",")

For i = UBound(arrDN) to 0 Step -1
    If Left(arrDN(i), 3) = "OU=" Then
        arrOU = Split(arrDN(i), "=")
        Wscript.Echo arrOU(1)
        Exit For
    End If
Next
  


