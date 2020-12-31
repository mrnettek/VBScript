' Description: Returns the name and product code of binary information (such as bitmaps, icons, executable files, and so on) used by a Windows Installer application.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_Binary")

For Each objItem in colItems
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Product Code: " & objItem.ProductCode
    Wscript.Echo
Next

