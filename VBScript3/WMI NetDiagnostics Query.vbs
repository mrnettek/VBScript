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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM NetDiagnostics", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "bIEProxy: " & objItem.bIEProxy
      WScript.Echo "id: " & objItem.id
      WScript.Echo "IEProxy: " & objItem.IEProxy
      WScript.Echo "IEProxyPort: " & objItem.IEProxyPort
      WScript.Echo "NewsNNTPPort: " & objItem.NewsNNTPPort
      WScript.Echo "NewsServer: " & objItem.NewsServer
      WScript.Echo
   Next
Next

