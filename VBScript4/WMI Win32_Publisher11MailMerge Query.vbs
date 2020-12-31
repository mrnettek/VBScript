On Error Resume Next

Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

arrComputers = Array(".")
For Each strComputer In arrComputers
   WScript.Echo
   WScript.Echo "=========================================="
   WScript.Echo "Computer: " & strComputer
   WScript.Echo "=========================================="

   Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\MSAPPS11")
   Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_Publisher11MailMerge", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "ActiveRecord: " & objItem.ActiveRecord
      WScript.Echo "DataSourceName: " & objItem.DataSourceName
      WScript.Echo "DataSourceType: " & objItem.DataSourceType
      WScript.Echo "DocumentName: " & objItem.DocumentName
      WScript.Echo "DocumentType: " & objItem.DocumentType
      WScript.Echo "FieldCodes: " & objItem.FieldCodes
      WScript.Echo "FieldCount: " & objItem.FieldCount
      WScript.Echo "FieldNames: " & objItem.FieldNames
      WScript.Echo "MergeTo: " & objItem.MergeTo
      WScript.Echo "Notables: " & objItem.Notables
      WScript.Echo "State: " & objItem.State
      WScript.Echo "SuppressBlankLines: " & objItem.SuppressBlankLines
      WScript.Echo "ViewFieldCodes: " & objItem.ViewFieldCodes
      WScript.Echo
   Next
Next

