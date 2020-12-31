strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_PnPSignedDriver")

For Each objItem in colItems
    Wscript.Echo "Device ID: " & objItem.DeviceID
    Wscript.Echo "Device Name: " & objItem.DeviceName
    dtmWMIDate = objItem.DriverDate
    strReturn = WMIDateStringToDate(dtmWMIDate)
    Wscript.Echo "Driver Date: " & strReturn
    Wscript.Echo "Driver Version: " & objItem.DriverVersion
    Wscript.Echo "Is Signed: " & objItem.IsSigned
    Wscript.Echo
Next
 
Function WMIDateStringToDate(dtmWMIDate)
    If Not IsNull(dtmWMIDate) Then
        WMIDateStringToDate = CDate(Mid(dtmWMIDate, 5, 2) & "/" & _
            Mid(dtmWMIDate, 7, 2) & "/" & Left(dtmWMIDate, 4) _
                & " " & Mid (dtmWMIDate, 9, 2) & ":" & _
                    Mid(dtmWMIDate, 11, 2) & ":" & Mid(dtmWMIDate,13, 2))
    End If
End Function
  


