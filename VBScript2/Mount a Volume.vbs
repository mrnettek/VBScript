' Description: Mounts volume Z to the file system. If you modify this script to mount a different volume (such as X), note that the volume name in your WQL query must include the drive letter followed by a colon and then followed by two  slashes. Thus drive X would be listed as X:\\.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * From Win32_Volume Where Name = 'Z:\\'")

For Each objItem in colItems
     objItem.Mount()
Next

