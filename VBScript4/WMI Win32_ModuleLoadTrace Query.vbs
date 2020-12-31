On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_ModuleLoadTrace",,48)
For Each objItem in colItems
    Wscript.Echo "FileName: " & objItem.FileName
    Wscript.Echo "ImageBase: " & objItem.ImageBase
    Wscript.Echo "ImageSize: " & objItem.ImageSize
    Wscript.Echo "ProcessID: " & objItem.ProcessID
    Wscript.Echo "SECURITY_DESCRIPTOR: " & objItem.SECURITY_DESCRIPTOR
    Wscript.Echo "TIME_CREATED: " & objItem.TIME_CREATED
Next

