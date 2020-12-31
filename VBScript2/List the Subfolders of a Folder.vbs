' Description: Uses a WMI Associators of query to list all the subfolders of the folder C:\Scripts.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colSubfolders = objWMIService.ExecQuery _
    ("Associators of {Win32_Directory.Name='c:\scripts'} " _
        & "Where AssocClass = Win32_Subdirectory " _
            & "ResultRole = PartComponent")

For Each objFolder in colSubfolders
    Wscript.Echo objFolder.Name
Next

