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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM MSStorageDriver_ATAPISmartData", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Active: " & objItem.Active
      WScript.Echo "Checksum: " & objItem.Checksum
      WScript.Echo "ErrorLogCapability: " & objItem.ErrorLogCapability
      WScript.Echo "ExtendedPollTimeInMinutes: " & objItem.ExtendedPollTimeInMinutes
      WScript.Echo "InstanceName: " & objItem.InstanceName
      WScript.Echo "Length: " & objItem.Length
      WScript.Echo "OfflineCollectCapability: " & objItem.OfflineCollectCapability
      WScript.Echo "OfflineCollectionStatus: " & objItem.OfflineCollectionStatus
      strReserved = Join(objItem.Reserved, ",")
         WScript.Echo "Reserved: " & strReserved
      WScript.Echo "SelfTestStatus: " & objItem.SelfTestStatus
      WScript.Echo "ShortPollTimeInMinutes: " & objItem.ShortPollTimeInMinutes
      WScript.Echo "SmartCapability: " & objItem.SmartCapability
      WScript.Echo "TotalTime: " & objItem.TotalTime
      strVendorSpecific = Join(objItem.VendorSpecific, ",")
         WScript.Echo "VendorSpecific: " & strVendorSpecific
      WScript.Echo "VendorSpecific2: " & objItem.VendorSpecific2
      WScript.Echo "VendorSpecific3: " & objItem.VendorSpecific3
      strVendorSpecific4 = Join(objItem.VendorSpecific4, ",")
         WScript.Echo "VendorSpecific4: " & strVendorSpecific4
      WScript.Echo
   Next
Next

