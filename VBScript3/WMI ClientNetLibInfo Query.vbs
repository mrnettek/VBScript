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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM ClientNetLibInfo", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Date: " & objItem.Date
      WScript.Echo "FileName: " & objItem.FileName
      WScript.Echo "ProtocolDisplayName: " & objItem.ProtocolDisplayName
      WScript.Echo "ProtocolName: " & objItem.ProtocolName
      WScript.Echo "Size: " & objItem.Size
      WScript.Echo "Version: " & objItem.Version
      WScript.Echo
   Next
Next

