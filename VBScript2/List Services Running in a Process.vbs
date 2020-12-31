' Description: Returns a list of services running in the Services.exe process.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colListOfServices = objWMIService.ExecQuery _
    ("Select * from Win32_Service Where " & _
        "PathName = 'C:\WINDOWS\system32\services.exe'")

For Each objService in colListOfServices
    Wscript.Echo objService.DisplayName
Next

