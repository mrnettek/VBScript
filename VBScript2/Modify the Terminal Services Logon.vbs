' Description: Configures the credentials (name, password, and domain) used to log on to Terminal Services.


Const TS_USER = "tsuser"
CONST TS_DOMAIN = "fabrikam"
CONST TS_PASSWORD = "password"
 
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_TSLogonSetting")

For Each objItem in colItems
    errResult = objItem.ExplicitLogon(TS_USER, TS_DOMAIN, TS_PASSWORD)
Next

