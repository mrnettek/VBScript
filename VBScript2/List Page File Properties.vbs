' Description: Lists the properties of all the page files on a computer.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colPageFiles = objWMIService.ExecQuery("Select * from Win32_PageFile")

For Each objPageFile in colPageFiles
    Wscript.Echo "Creation Date: " & objPageFile.CreationDate
    Wscript.Echo "Description: " & objPageFile.Description
    Wscript.Echo "Drive: " & objPageFile.Drive        
    Wscript.Echo "File Name: " & objPageFile.FileName  
    Wscript.Echo "File Size: " & objPageFile.FileSize  
    Wscript.Echo "Initial Size: " & objPageFile.InitialSize
    Wscript.Echo "Install Date: " & objPageFile.InstallDate
    Wscript.Echo "Maximum Size: " & objPageFile.MaximumSize
    Wscript.Echo "Name: " & objPageFile.Name  
    Wscript.Echo "Path: " & objPageFile.Path  
Next

