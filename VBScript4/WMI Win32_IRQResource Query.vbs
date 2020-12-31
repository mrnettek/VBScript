On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_IRQResource",,48)
For Each objItem in colItems
    Wscript.Echo "Availability: " & objItem.Availability
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "CreationClassName: " & objItem.CreationClassName
    Wscript.Echo "CSCreationClassName: " & objItem.CSCreationClassName
    Wscript.Echo "CSName: " & objItem.CSName
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Hardware: " & objItem.Hardware
    Wscript.Echo "InstallDate: " & objItem.InstallDate
    Wscript.Echo "IRQNumber: " & objItem.IRQNumber
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Shareable: " & objItem.Shareable
    Wscript.Echo "Status: " & objItem.Status
    Wscript.Echo "TriggerLevel: " & objItem.TriggerLevel
    Wscript.Echo "TriggerType: " & objItem.TriggerType
    Wscript.Echo "Vector: " & objItem.Vector
Next

