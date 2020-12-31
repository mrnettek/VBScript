' Description: Sets the IP connection metric for a network adapter to 100. Connection metrics can range from 1 to 9,999, with a default value of 1.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colNetCards = objWMIService.ExecQuery _
    ("Select * From Win32_NetworkAdapterConfiguration Where IPEnabled = True")

For Each objNetCard in colNetCards
    objNetCard.SetIPConnectionMetric(100)
Next

