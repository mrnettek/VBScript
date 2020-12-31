' Description: Issues an alert if free disk space for any hard drive on a computer falls below 100 megabytes.


Const LOCAL_HARD_DISK = 3

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colMonitoredDisks = objWMIService.ExecNotificationQuery _
    ("Select * from __instancemodificationevent within 30 where " _
        & "TargetInstance isa 'Win32_LogicalDisk'")
i = 0

Do While i = 0
    Set objDiskChange = colMonitoredDisks.NextEvent
    If objDiskChange.TargetInstance.DriveType = LOCAL_HARD_DISK Then
        If objDiskChange.TargetInstance.Size < 100000000 Then
            Wscript.Echo "Hard disk space is below 100000000 bytes."
        End If
    End If
Loop

