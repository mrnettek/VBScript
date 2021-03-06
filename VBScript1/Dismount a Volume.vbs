' Description: Dismounts volume E from the file system. If you modify this script to dismount a different volume (such as X), note that your WQL query must specify the drive letter followed by a colon and then followed by two slashes. Thus volume X would be listed as X:\\.

The two parameters: 1) force the volume to be dismounted, even if users are currently connected to it; and, 2) place the volume in a no-automount, offline state. This script can be modified by setting either (or both) of these parameters to False.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * From Win32_Volume Where Name = 'E:\\'")

For Each objItem in colItems
    objItem.Dismount(True, True)
Next

