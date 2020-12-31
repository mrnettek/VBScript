' Description: Retrieves service data from a computer, and then outputs that data in tabular format in a command window.


Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colServices = objWMIService.ExecQuery("Select * from Win32_Service")

For Each objService in colServices
    intPadding = 50 - Len(objService.DisplayName)
    intPadding2 = 17 - Len(objService.StartMode)
    strDisplayName = objService.DisplayName & Space(intPadding)
    strStartMode = objService.StartMode & Space(intPadding2)
    Wscript.Echo strDisplayName & strStartMode & objService.State 
Next

