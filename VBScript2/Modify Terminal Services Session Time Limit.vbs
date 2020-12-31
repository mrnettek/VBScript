' Description: Allows unlimited-length active sessions in Terminal Services. To enforce a time limit on active sessions, pass the TimeLimit method a value other than 0. Time limits must be expressed in milliseconds; for example, a one-minute session limit would be expressed as 60000: 60 seconds times 1000 milliseconds. A one-hour time limit would be expressed as 3600000.


Const NO_SESSION_LIMIT = 0
 
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_TSSessionSetting")

For Each objItem in colItems
    errResult = objItem.TimeLimit("ActiveSessionLimit", NO_SESSION_LIMIT)
Next

