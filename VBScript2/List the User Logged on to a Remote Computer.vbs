' Description: Returns the user name of the user currently logged on to a remote computer. To use this script, replace atl-ws-01 with the name of the remote computer you want to check. Although this script will run on Windows NT 4.0, Windows 98, and Windows 2000, it will not always return information.


strComputer = "atl-ws-o1"
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 

Set colComputer = objWMIService.ExecQuery _
    ("Select * from Win32_ComputerSystem")
 
For Each objComputer in colComputer
    Wscript.Echo "Logged-on user: " & objComputer.UserName
Next

