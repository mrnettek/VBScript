' Description: Resets a computer account password in Active Directory.


Set objComputer = GetObject _
    ("LDAP://CN=atl-dc-01,CN=Computers,DC=Reskit,DC=COM")

objComputer.SetPassword "atl-dc-01$"

