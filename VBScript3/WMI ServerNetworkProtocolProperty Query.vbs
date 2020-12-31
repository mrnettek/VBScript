On Error Resume Next

Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

arrComputers = Array(".")
For Each strComputer In arrComputers
   WScript.Echo
   WScript.Echo "=========================================="
   WScript.Echo "Computer: " & strComputer
   WScript.Echo "=========================================="

   Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\Microsoft\SqlServer\ComputerManagement10")
   Set colItems = objWMIService.ExecQuery("SELECT * FROM ServerNetworkProtocolProperty", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "InstanceName: " & objItem.InstanceName
      WScript.Echo "IPAddressName: " & objItem.IPAddressName
      WScript.Echo "PropertyName: " & objItem.PropertyName
      WScript.Echo "PropertyNumVal: " & objItem.PropertyNumVal
      WScript.Echo "PropertyStrVal: " & objItem.PropertyStrVal
      WScript.Echo "PropertyType: " & objItem.PropertyType
      WScript.Echo "PropertyValType: " & objItem.PropertyValType
      WScript.Echo "ProtocolName: " & objItem.ProtocolName
      WScript.Echo
   Next
Next

