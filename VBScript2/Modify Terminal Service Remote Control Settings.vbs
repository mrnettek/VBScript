' Description: Enables remote user to view, but not actively control, a session. Permission of the logged-on session user is not required.


Const ENABLE_NO_INPUT_NO_NOTIFY = 4
 
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
  ("Select * from Win32_TSRemoteControlSetting Where " & _
        "TerminalName = 'Accounting'")

For Each objItem in colItems
  errResult = objItem.RemoteControl(ENABLE_NO_INPUT_NO_NOTIFY)
Next

