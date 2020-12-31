strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * From Win32_OperatingSystem")

For Each objItem in colItems
    If Instr(objItem.Caption, "Vista") Then
        Select Case objItem.OperatingSystemSKU
             Case 0 strVersion = "Undefined."
             Case 1 strVersion = "Ultimate Edition." 
             Case 2 strVersion = "Home Basic Edition." 
             Case 3 strVersion = "Home Basic Premium Edition." 
             Case 4 strVersion = "Enterprise Edition." 
             Case 5 strVersion = "Home Basic N Edition." 
             Case 6 strVersion = "Business Edition." 
             Case 7 strVersion = "Standard Server Edition." 
             Case 8 strVersion = "Datacenter Server Edition." 
             Case 9 strVersion = "Small Business Server Edition." 
             Case 10 strVersion = "Enterprise Server Edition." 
             Case 11 strVersion = "Starter Edition." 
             Case 12 strVersion = "Datacenter Server Core Edition." 
             Case 13 strVersion = "Standard Server Core Edition." 
             Case 14 strVersion = "Enterprise Server Core Edition." 
             Case 15 strVersion = "Enterprise Server IA64 Edition." 
             Case 16 strVersion = "Business N Edition." 
             Case 17 strVersion = "Web Server Edition." 
             Case 18 strVersion = "Cluster Server Edition." 
             Case 19 strVersion = "Home Server Edition." 
             Case 20 strVersion = "Storage Express Server Edition." 
             Case 21 strVersion = "Storage Standard Server Edition." 
             Case 22 strVersion = "Storage Workgroup Server Edition." 
             Case 23 strVersion = "Storage Enterprise Server Edition." 
             Case 24 strVersion = "Server For Small Business Edition." 
             Case 25 strVersion = "Small Business Server Premium Edition." 
        End Select
        Wscript.Echo "This computer is running Windows Vista " & strVersion
    Else
        Wscript.Echo "This computer is not running Windows Vista."
    End If
Next
  


