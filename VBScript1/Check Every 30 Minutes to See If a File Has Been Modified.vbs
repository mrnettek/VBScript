strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colFiles = objWMIService.ExecQuery _
    ("Select * from CIM_Datafile Where Name = 'C:\\Scripts\\Application.log'")

For Each objFile in colFiles
    strOriginalTimestamp = objFile.LastModified
Next

Wscript.Echo "Monitoring application log file: " & Now
Wscript.Echo

Do While True
    Wscript.Sleep 1800000
    Set colFiles = objWMIService.ExecQuery _
        ("Select * from CIM_Datafile Where Name = 'C:\\Scripts\\Application.log'")

    For Each objFile in colFiles
        strLatestTimestamp = objFile.LastModified
    Next 

    If strLatestTimestamp <> strOriginalTimestamp Then
        strOriginalTimestamp = strLatestTimeStamp
    Else
        Wscript.Echo "ALERT: " & Now
        Wscript.Echo "The application log file has not been modified in the last 30 minutes."
        Wscript.Echo
        strOriginalTimestamp = strLatestTimeStamp
    End If
Loop
  


