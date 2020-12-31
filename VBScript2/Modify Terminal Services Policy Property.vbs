' Description: Disables the UseTempFolders policy on a computer running Terminal Services. To enable this policy, pass the value 1 (rather than 0) to the SetPolicyPropertyName method. To configure a different policy, simply replace the method parameter “UseTempFolders” with the appropriate policy name.


Const DISABLE_POLICY = 0
 
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_TerminalServiceSetting")

For Each objItem in colItems
    errResult = objItem.SetPolicyPropertyName("UseTempFolders", DISABLE_POLICY)
Next

