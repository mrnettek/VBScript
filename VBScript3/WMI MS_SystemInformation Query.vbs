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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM MS_SystemInformation", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Active: " & objItem.Active
      WScript.Echo "BaseBoardManufacturer: " & objItem.BaseBoardManufacturer
      WScript.Echo "BaseBoardProduct: " & objItem.BaseBoardProduct
      WScript.Echo "BaseBoardVersion: " & objItem.BaseBoardVersion
      WScript.Echo "BIOSReleaseDate: " & objItem.BIOSReleaseDate
      WScript.Echo "BIOSVendor: " & objItem.BIOSVendor
      WScript.Echo "BIOSVersion: " & objItem.BIOSVersion
      WScript.Echo "InstanceName: " & objItem.InstanceName
      WScript.Echo "SystemManufacturer: " & objItem.SystemManufacturer
      WScript.Echo "SystemProductName: " & objItem.SystemProductName
      WScript.Echo "SystemVersion: " & objItem.SystemVersion
      WScript.Echo
   Next
Next

