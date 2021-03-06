' Description: Creates a shadow copy storage area -- on drive E -- for storing shadow copies of drive C. This script reserves 130,023,424 bytes of storage space for the shadow copies.


Const VOLUME = "C:\"
Const DIFFERENTIAL_VOLUME = "E:\"
Const MAXIMUM_SPACE = 130023424
 
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set objShadowStorage = objWMIService.Get("Win32_ShadowStorage")
errResult = objShadowStorage.Create(VOLUME, DIFFERENTIAL_VOLUME, MAXIMUM_SPACE)

