On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_ThreadStartTrace",,48)
For Each objItem in colItems
    Wscript.Echo "ProcessID: " & objItem.ProcessID
    Wscript.Echo "SECURITY_DESCRIPTOR: " & objItem.SECURITY_DESCRIPTOR
    Wscript.Echo "StackBase: " & objItem.StackBase
    Wscript.Echo "StackLimit: " & objItem.StackLimit
    Wscript.Echo "StartAddr: " & objItem.StartAddr
    Wscript.Echo "ThreadID: " & objItem.ThreadID
    Wscript.Echo "TIME_CREATED: " & objItem.TIME_CREATED
    Wscript.Echo "UserStackBase: " & objItem.UserStackBase
    Wscript.Echo "UserStackLimit: " & objItem.UserStackLimit
    Wscript.Echo "WaitMode: " & objItem.WaitMode
    Wscript.Echo "Win32StartAddr: " & objItem.Win32StartAddr
Next

