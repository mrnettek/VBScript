' Description: Identifies all the TCP/IP printer ports on a computer, and indicates which ports are being used and which ports are available.


Set objDictionary = CreateObject("Scripting.Dictionary")

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colPrinters =  objWMIService.ExecQuery _
    ("Select * from Win32_Printer")

For Each objPrinter in colPrinters 
    objDictionary.Add objPrinter.PortName, objPrinter.PortName
Next

Set colPorts = objWMIService.ExecQuery _
    ("Select * from Win32_TCPIPPrinterPort")
For Each objPort in colPorts
    If objDictionary.Exists(objPort.Name) Then
        strUsedPorts = strUsedPorts & _
            objDictionary.Item(objPort.Name) & VbCrLf
    Else
        strFreePorts = strFreePorts & objPort.Name & vbCrLf
    End If
Next

Wscript.Echo "The following ports are in use: " & VbCrLf & strUsedPorts
Wscript.Echo "The following ports are available: " & VbCrLf & strFreePorts

