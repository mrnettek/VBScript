' Description: Returns information about Group Policy Objects used in applying Group Policy to a computer. To return information about Group Policy applied to the logged-on user, change the WMI class from root\rsop\computer to root\rsop\user.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:\\" & strComputer & "\root\rsop\computer")

Set colItems = objWMIService.ExecQuery("Select * from RSOP_GPO")

For Each objItem in colItems  
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "GUID Name: " & objItem.GUIDName
    Wscript.Echo "ID: " & objItem.ID
    Wscript.Echo "Access Denied: " & objItem.AccessDenied
    Wscript.Echo "Enabled: " & objItem.Enabled
    Wscript.Echo "File System path: " & objItem.FileSystemPath
    Wscript.Echo "Filter Allowed: " & objItem.FilterAllowed
    Wscript.Echo "Filter ID: " & objItem.FilterId
    Wscript.Echo "Version: " & objItem.Version
    Wscript.Echo
Next

