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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM SqlServiceAdvancedProperty", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "IsReadOnly: " & objItem.IsReadOnly
      WScript.Echo "PropertyIndex: " & objItem.PropertyIndex
      WScript.Echo "PropertyName: " & objItem.PropertyName
      WScript.Echo "PropertyNumValue: " & objItem.PropertyNumValue
      WScript.Echo "PropertyStrValue: " & objItem.PropertyStrValue
      WScript.Echo "PropertyValueType: " & objItem.PropertyValueType
      WScript.Echo "ServiceName: " & objItem.ServiceName
      WScript.Echo "SqlServiceType: " & objItem.SqlServiceType
      WScript.Echo
   Next
Next

