On Error Resume Next

Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

arrComputers = Array(".")
For Each strComputer In arrComputers
   WScript.Echo
   WScript.Echo "=========================================="
   WScript.Echo "Computer: " & strComputer
   WScript.Echo "=========================================="

   Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
   Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_LogicalShareAccess", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "AccessMask: " & objItem.AccessMask
      WScript.Echo "GuidInheritedObjectType: " & objItem.GuidInheritedObjectType
      WScript.Echo "GuidObjectType: " & objItem.GuidObjectType
      WScript.Echo "Inheritance: " & objItem.Inheritance
      WScript.Echo "SecuritySetting: " & objItem.SecuritySetting
      WScript.Echo "Trustee: " & objItem.Trustee
      WScript.Echo "Type: " & objItem.Type
      WScript.Echo
   Next
Next

