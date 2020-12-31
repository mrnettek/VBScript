' Description: Returns the properties of each Internet Connection Firewall connection.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & _
        "\root\microsoft\homenet")

Set colItems = objWMIService.ExecQuery("Select * from HNet_Connection")

For Each objItem in colItems
    Wscript.Echo "GUID: " & objItem.GUID
    Wscript.Echo "Is LAN Connection: " & objItem.IsLANConnection
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Phone Book Path: " & objItem.PhoneBookPath
Next

