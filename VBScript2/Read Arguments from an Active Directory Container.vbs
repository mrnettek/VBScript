' Description: Demonstration script that retrieves the names of all the computers in an Active Directory container, and then returns service information from each of those computers.


Set objDictionary = CreateObject("Scripting.Dictionary")

i = 0
Set objOU = GetObject("LDAP://CN=Computers, DC=fabrikam, DC=com")
objOU.Filter = Array("Computer")

For Each objComputer in objOU 
    objDictionary.Add i, objComputer.CN
    i = i + 1
Next

For Each objItem in objDictionary
    Set colServices = GetObject("winmgmts://" & _
        objDictionary.Item(objItem) _
            & "").ExecQuery("Select * from Win32_Service")
    Wscript.Echo colServices.Count
Next

