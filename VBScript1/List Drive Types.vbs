' Description: Identifies the drive type (floppy drive, hard drive, CD-ROM, etc.) for each physical drive installed on a computer.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colDisks = objWMIService.ExecQuery _
    ("Select * from Win32_LogicalDisk")

For Each objDisk in colDisks
    Wscript.Echo "DeviceID: "& objDisk.DeviceID       
    Select Case objDisk.DriveType
        Case 1
            Wscript.Echo "No root directory. Drive type could not be " _
                & "determined."
        Case 2
            Wscript.Echo "DriveType: "& "Removable drive."
        Case 3
            Wscript.Echo "DriveType: "& "Local hard disk."
        Case 4
            Wscript.Echo "DriveType: "& "Network disk."      
        Case 5
            Wscript.Echo "DriveType: "& "Compact disk."      
        Case 6
            Wscript.Echo "DriveType: "& "RAM disk."   
        Case Else
            Wscript.Echo "Drive type could not be determined."
    End Select
Next

