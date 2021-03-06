' Description: Formats drive D using the NTFS file system. Drives could also be formatted using the FAT or FAT32 file systems. Note that this method cannot be used to format floppy drives. If you modify this script to format a different volume (such as drive X), also note that your WQL query must specify the drive letter followed by a colon and then followed by two  slashes. Thus drive X would be listed as X:\\.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colVolumes = objWMIService.ExecQuery _
    ("Select * from Win32_Volume Where Name = 'D:\\'")

For Each objVolume in colVolumes
    errResult = objVolume.Format("NTFS")
Next

