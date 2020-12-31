' Description: Demonstration script that changes the location attribute for a computer account in Active Directory.


Set objComputer = GetObject _ 
    ("LDAP://CN=atl-dc-01,CN=Computers,DC=fabrikam,DC=com")

objComputer.Put "Location" , "Building 37, Floor 2, Room 2133"
objComputer.SetInfo

