' Description: Deletes an individual computer account in Active Directory.


strComputer = "atl-pro-040"

set objComputer = GetObject("LDAP://CN=" & strComputer & _
    ",CN=Computers,DC=fabrikam,DC=com")
objComputer.DeleteObject (0)

