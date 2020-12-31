On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_ProcessStopTrace",,48)
For Each objItem in colItems
    Wscript.Echo "PageDirectoryBase: " & objItem.PageDirectoryBase
    Wscript.Echo "ParentProcessID: " & objItem.ParentProcessID
    Wscript.Echo "ProcessID: " & objItem.ProcessID
    Wscript.Echo "ProcessName: " & objItem.ProcessName
    Wscript.Echo "SECURITY_DESCRIPTOR: " & objItem.SECURITY_DESCRIPTOR
    Wscript.Echo "SessionID: " & objItem.SessionID
    Wscript.Echo "Sid: " & objItem.Sid
    Wscript.Echo "TIME_CREATED: " & objItem.TIME_CREATED
Next

