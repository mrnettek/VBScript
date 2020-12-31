' MrNetTek
' eddiejackson.net/blog
' 10/14/2019
' free for public use 
' free to claim as your own

on error resume next
 
Const HARD_DISK = 3
 
strComputer = "."
 
Set objShell = CreateObject("WScript.Shell")
 
RegPath = "hklm\software\HD_Specs"
 
 
'RETURN ALL LOCAL HARD DRIVES
 
Set objWMIService = GetObject("winmgmts:"  & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
 
Set colDisks = objWMIService.ExecQuery("SELECT * FROM Win32_LogicalDisk")
 
For Each objDisk in colDisks 
  
Select Case objDisk.DriveType
 
 Case 1
    'Wscript.Echo "No root directory."
 Case 2
    'Wscript.Echo objDisk.DeviceID  & " Removable drive."
     
 Case 3
         objShell.Run "reg add " & chr(34) & RegPath & chr(34) " /v Drive_" & objDisk.DeviceID & " /t REG_SZ /d " & " ""Local hard disk"" "  & " /reg:64 /f",0,true
 Case 4
    'Wscript.Echo objDisk.DeviceID  & " Network disk." 
 Case 5
    'Wscript.Echo objDisk.DeviceID  & " Compact disk." 
 Case 6
    'Wscript.Echo objDisk.DeviceID  & " RAM disk." 
Case Else
    'Wscript.Echo objDisk.DeviceID  & " Drive type could not be determined." 
End Select
 
Next
 
 
Set objWMIService = GetObject("winmgmts:"  & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
 
Set colDisks = objWMIService.ExecQuery("Select * from Win32_LogicalDisk Where DriveType = " & HARD_DISK & " AND DeviceID = 'C:'")
 
'FREE SPACE
For Each objDisk in colDisks   
    
   numSize = FormatNumber(objDisk.FreeSpace /(1024^3), 3 )
    
   numSize = left(numSize, len(numSize) -1 )
    
   objShell.Run "reg add " & chr(34) & RegPath & chr(34) " /v FreeSpace /t REG_SZ /d " & numSize  & " /reg:64 /f",0,true
 
Next
 
numSize = ""
 
 
'SIZE
For Each objDisk in colDisks   
 
   numSize = Int(objDisk.Size /1073741824)   
 
   objShell.Run "reg add " & chr(34) & RegPath & chr(34) " /v TotalSize /t REG_SZ /d " & numSize  & " /reg:64 /f",0,true
 
Next
 
 
'DIRTY BIT
For Each objDisk in colDisks   
    
   objShell.Run "reg add " & chr(34) & RegPath & chr(34) " /v DirtyBit /t REG_SZ /d " & objDisk.VolumeDirty  & " /reg:64 /f",0,true
    
Next
 
 
'MODEL
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
 
Set colItems = objWMIService.ExecQuery("Select * from Win32_DiskDrive")
 
For Each objItem in colItems    
 
    objShell.Run "reg add " & chr(34) & RegPath & chr(34) " /v Model /t REG_SZ /d " & chr(34) & objItem.Model & chr(34) & " /reg:64 /f",0,true
 
Next
 
objShell.Run "cmd /c " & chr(34) & "C:\Program Files (x86)\LANDesk\LDClient\LDISCN32.EXE" & chr(34) & " /F /SYNC",0,true
 
WScript.Quit