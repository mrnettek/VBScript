' Description: Enumerates all the volume mount points on a computer.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery _
    ("SELECT * FROM Win32_MountPoint")

For Each objItem In colItems
  WScript.Echo "Directory: " & objItem.Directory
  WScript.Echo "Volume: " & objItem.Volume
  WScript.Echo
Next

