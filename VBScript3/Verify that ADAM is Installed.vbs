strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colServices = objWMIService.ExecQuery _
    ("Select * from Win32_Service where Name Like '%ADAM_%'")

If colServices.Count = 0 Then
    Wscript.Echo "ADAM is not installed."
Else
    For Each objService in colServices
        Wscript.Echo objService.Name & " -- " & objService.State
    Next
End If
  


