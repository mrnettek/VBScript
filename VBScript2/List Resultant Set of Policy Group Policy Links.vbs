' Description: Returns information about Group Policy links used in applying Group Policy to a computer. To return information about Group Policy applied to the logged-on user, change the WMI class from root\rsop\computer to root\rsop\user.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:\\" & strComputer & "\root\rsop\computer")

Set colItems = objWMIService.ExecQuery("Select * from RSOP_GPLink")

For Each objItem in colItems
    Wscript.Echo "GPO: " & objItem.GPO
    Wscript.Echo "Applied Order: " & objItem.AppliedOrder
    Wscript.Echo "Enabled: " & objItem.Enabled
    Wscript.Echo "Link Order: " & objItem.LinkOrder
    Wscript.Echo "No Overrride: " & objItem.NoOverride
    Wscript.Echo "SOM Order: " & objItem.SOMOrder
    Wscript.Echo
Next

