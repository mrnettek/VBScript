' Description: Performs a system restore on a computer using system restore point No. 20. To perform a system restore using a different system restore point, simply change the value of the constant RESTORE_POINT.


Const RESTORE_POINT = 20
 
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\default")

Set objItem = objWMIService.Get("SystemRestore")
errResults = objItem.Restore(RESTORE_POINT)

