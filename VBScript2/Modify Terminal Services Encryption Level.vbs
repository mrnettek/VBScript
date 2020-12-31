' Description: Configures Terminal Services to use high encryption (parameter value of 3). To use low-encryption, pass the SetEncryptionLevel method a value of 1. Client-compatible encryption uses a value of 2, and FIPS-compliant encryption uses a value of 4.


Const HIGH_ENCRYPTION = 3
 
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_TSGeneralSetting")

For Each objItem in colItems
    errResult = objItem.SetEncryptionLevel(HIGH_ENCRYPTION)
Next

