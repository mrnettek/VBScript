' Description: Disables system restore on a computer. This is equivalent to selecting the checkbox Turn off System Restore (found by right-clicking My Computer, clicking Properties, and then clicking on the System Restore tab in the resulting dialog box).


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\default")

Set objItem = objWMIService.Get("SystemRestore")
errResults = objItem.Disable("")

