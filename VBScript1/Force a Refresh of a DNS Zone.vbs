' Description: Forces the refresh of the secondary DNS zone accounting.fabrikam.com from its master zone.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & _
        "\root\MicrosoftDNS")

Set colItems = objWMIService.ExecQuery _
    ("Select * from MicrosoftDNS_Zone Where Name = 'accounting.fabrikam.com'")

For Each objItem in colItems
    objItem.ForceRefresh()
Next

