On Error Resume Next

Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

arrComputers = Array(".")
For Each strComputer In arrComputers
   WScript.Echo
   WScript.Echo "=========================================="
   WScript.Echo "Computer: " & strComputer
   WScript.Echo "=========================================="

   Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\WMI")
   Set colItems = objWMIService.ExecQuery("SELECT * FROM MSDiskDriver_Geometry", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Active: " & objItem.Active
      WScript.Echo "BytesPerSector: " & objItem.BytesPerSector
      WScript.Echo "Cylinders: " & objItem.Cylinders
      WScript.Echo "InstanceName: " & objItem.InstanceName
      WScript.Echo "MediaType: " & objItem.MediaType
      WScript.Echo "SectorsPerTrack: " & objItem.SectorsPerTrack
      WScript.Echo "TracksPerCylinder: " & objItem.TracksPerCylinder
      WScript.Echo
   Next
Next

