strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_TerminalServiceSetting")

For Each objItem in colItems
    If objItem.AllowTSConnections = 0 Then
        strStatus = "Disabled"
    Else
        strStatus = "Enabled"
    End If
    Wscript.Echo "Terminal Services status: " & strStatus
Next
  


