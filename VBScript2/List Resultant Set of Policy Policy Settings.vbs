' Description: Returns information about policy object settings used in applying Group Policy to a computer. To return information about Group Policy applied to the logged-on user, change the WMI class from root\rsop\computer to root\rsop\user.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:\\" & strComputer & "\root\rsop\computer")

Set colItems = objWMIService.ExecQuery("Select * from RSOP_PolicySetting")

For Each objItem in colItems
    Wscript.Echo "GPO ID: " & objItem.GPOID
    Wscript.Echo "ID: " & objItem.ID
    Wscript.Echo "Precedence: " & objItem.Precedence
    Wscript.Echo "SOM ID: " & objItem.SOMID
    Wscript.Echo
Next

