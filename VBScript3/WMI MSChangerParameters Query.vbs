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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM MSChangerParameters", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Active: " & objItem.Active
      WScript.Echo "InstanceName: " & objItem.InstanceName
      WScript.Echo "MagazineSize: " & objItem.MagazineSize
      WScript.Echo "NumberOfCleanerSlots: " & objItem.NumberOfCleanerSlots
      WScript.Echo "NumberOfDoors: " & objItem.NumberOfDoors
      WScript.Echo "NumberOfDrives: " & objItem.NumberOfDrives
      WScript.Echo "NumberOfIEPorts: " & objItem.NumberOfIEPorts
      WScript.Echo "NumberOfSlots: " & objItem.NumberOfSlots
      WScript.Echo "NumberOfTransports: " & objItem.NumberOfTransports
      WScript.Echo
   Next
Next

