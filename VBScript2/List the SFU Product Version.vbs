' Description: Displays the Services for UNIX product version.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\sfuadmin")

Set colItems = objWMIService.ExecQuery("Select * from ProductVer")

For Each objItem in colItems
    Wscript.Echo "KeyName: " & objItem.KeyName
    Wscript.Echo "Operating System Version: " & objItem.OsVersion
    Wscript.Echo "Software Type: " & objItem.SoftwareType
    Wscript.Echo
Next

