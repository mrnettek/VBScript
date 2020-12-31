' Description: Uses WMI to move the folder C:\Scripts to C:\Admins\Documents\Archive\VBScript.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colFolders = objWMIService.ExecQuery _
    ("Select * from Win32_Directory where name = 'c:\\Scripts'")

For Each objFolder in colFolders
    errResults = objFolder.Rename("C:\Admins\Documents\Archive\VBScript")
Next

