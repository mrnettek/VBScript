On Error Resume Next

Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

arrComputers = Array(".")
For Each strComputer In arrComputers
   WScript.Echo
   WScript.Echo "=========================================="
   WScript.Echo "Computer: " & strComputer
   WScript.Echo "=========================================="

   Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\ServiceModel")
   Set colItems = objWMIService.ExecQuery("SELECT * FROM XmlDictionaryReaderQuotas", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "MaxArrayLength: " & objItem.MaxArrayLength
      WScript.Echo "MaxBytesPerRead: " & objItem.MaxBytesPerRead
      WScript.Echo "MaxDepth: " & objItem.MaxDepth
      WScript.Echo "MaxNameTableCharCount: " & objItem.MaxNameTableCharCount
      WScript.Echo "MaxStringContentLength: " & objItem.MaxStringContentLength
      WScript.Echo
   Next
Next

