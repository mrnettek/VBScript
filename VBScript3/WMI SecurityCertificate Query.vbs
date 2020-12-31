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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM SecurityCertificate", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      strContext = Join(objItem.Context, ",")
         WScript.Echo "Context: " & strContext
      WScript.Echo "ExpirationDate: " & objItem.ExpirationDate
      WScript.Echo "FriendlyName: " & objItem.FriendlyName
      WScript.Echo "IssuedBy: " & objItem.IssuedBy
      WScript.Echo "IssuedTo: " & objItem.IssuedTo
      WScript.Echo "Name: " & objItem.Name
      WScript.Echo "SHA: " & objItem.SHA
      WScript.Echo "StartDate: " & objItem.StartDate
      WScript.Echo "SystemStore: " & objItem.SystemStore
      WScript.Echo
   Next
Next

