Set objUser = GetObject("LDAP://cn=Ken Myer,ou=Finance,dc=fabrikam,dc=com")
strHomeDirectory = objUser.homeDirectory

Set objExcel = CreateObject("Excel.Application")
objExcel.DefaultFilePath = strHomeDirectory
objExcel.Quit
  


