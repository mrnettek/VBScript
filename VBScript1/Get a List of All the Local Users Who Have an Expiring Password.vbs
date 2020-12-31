Const ADS_UF_DONT_EXPIRE_PASSWD = &H10000
 
strComputer = "atl-ws-01"

Set colAccounts = GetObject("WinNT://" & strComputer & "")

colAccounts.Filter = Array("user")

For Each objUser In colAccounts
    If Not objUser.UserFlags AND ADS_UF_DONT_EXPIRE_PASSWD Then
        Wscript.Echo objUser.Name
    End If
Next
  


