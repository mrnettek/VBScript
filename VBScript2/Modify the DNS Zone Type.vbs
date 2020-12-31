' Description: Changes the DNS zone type of the zone accounting.fabrikam.com to a primary DNS zone.


intPrimaryZone = 1
 
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & _
        "\root\MicrosoftDNS")

Set colItems = objWMIService.ExecQuery _
    ("Select * from MicrosoftDNS_Zone Where Name = 'accounting.fabrikam.com'")

For Each objItem in colItems
    errResult = objItem.ChangeZoneType(intPrimaryZone, true)
Next

