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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM MSSmBios_RawSMBiosTables", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Active: " & objItem.Active
      WScript.Echo "DmiRevision: " & objItem.DmiRevision
      WScript.Echo "InstanceName: " & objItem.InstanceName
      WScript.Echo "Size: " & objItem.Size
      strSMBiosData = Join(objItem.SMBiosData, ",")
         WScript.Echo "SMBiosData: " & strSMBiosData
      WScript.Echo "SmbiosMajorVersion: " & objItem.SmbiosMajorVersion
      WScript.Echo "SmbiosMinorVersion: " & objItem.SmbiosMinorVersion
      WScript.Echo "Used20CallingMethod: " & objItem.Used20CallingMethod
      WScript.Echo
   Next
Next

