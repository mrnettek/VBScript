Dim arrNames()
intSize = 0

Set objGroup = GetObject("LDAP://CN=Accountants,OU=Finance,DC=fabrikam,DC=com")

For Each strUser in objGroup.Member
    Set objUser =  GetObject("LDAP://" & strUser)
    ReDim Preserve arrNames(intSize)
    arrNames(intSize) = objUser.CN
    intSize = intSize + 1
Next

For i = (UBound(arrNames) - 1) to 0 Step -1
    For j= 0 to i
        If UCase(arrNames(j)) > UCase(arrNames(j+1)) Then
            strHolder = arrNames(j+1)
            arrNames(j+1) = arrNames(j)
            arrNames(j) = strHolder
        End If
    Next
Next 

For Each strName in arrNames
    Wscript.Echo strName
Next
  


