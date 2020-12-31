strComputer = "."
i = 0

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colControllers = objWMIService.ExecQuery _
    ("Select * From Win32_USBController")

For Each objController in colControllers
    If Instr(objController.Name, "Enhanced") Then
        i = i + 1
    End If
Next

Wscript.Echo "No. of USB 2.0 Ports: " & i
  


