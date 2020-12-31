' Description: Disables audio mapping on a computer running Terminal Services. To enable audio mapping, pass the value 1 (rather than 0) to the SetClientProperty method. To configure a different property value, replace the parameter AudioMapping with the appropriate property name.


Const DISABLE = 0
 
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_TSClientSetting")

For Each objItem in colItems
    errResult = objItem.SetClientProperty("AudioMapping", DISABLE)
Next

