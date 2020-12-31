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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_AccessProject", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "CreateDate: " & WMIDateStringToDate(objItem.CreateDate)
      WScript.Echo "DatabaseName: " & objItem.DatabaseName
      WScript.Echo "DBMSName: " & objItem.DBMSName
      WScript.Echo "DBMSVersion: " & objItem.DBMSVersion
      WScript.Echo "DSNName: " & objItem.DSNName
      WScript.Echo "LoginName: " & objItem.LoginName
      WScript.Echo "Name: " & objItem.Name
      WScript.Echo "Path: " & objItem.Path
      WScript.Echo "ServerName: " & objItem.ServerName
      WScript.Echo "Size: " & objItem.Size
      WScript.Echo
   Next
Next


Function WMIDateStringToDate(dtmDate)
WScript.Echo dtm: 
	WMIDateStringToDate = CDate(Mid(dtmDate, 5, 2) & "/" & _
	Mid(dtmDate, 7, 2) & "/" & Left(dtmDate, 4) _
	& " " & Mid (dtmDate, 9, 2) & ":" & Mid(dtmDate, 11, 2) & ":" & Mid(dtmDate,13, 2))
End Function

