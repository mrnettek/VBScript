Const ADS_UF_PASSWD_CANT_CHANGE = &H0040

Set objUser = GetObject("WinNT://atl-ws-01/kenmyer")

If Not objUser.UserFlags AND ADS_UF_PASSWD_CANT_CHANGE Then
    objPasswordNoChangeFlag = objUser.UserFlags XOR ADS_UF_PASSWD_CANT_CHANGE
    objUser.Put "userFlags", objPasswordNoChangeFlag 
    objUser.SetInfo
End If
  


