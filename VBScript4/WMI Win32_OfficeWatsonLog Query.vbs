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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_OfficeWatsonLog", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "AppName: " & objItem.AppName
      WScript.Echo "AppStamp: " & objItem.AppStamp
      WScript.Echo "AppVersion: " & objItem.AppVersion
      WScript.Echo "BucketID: " & objItem.BucketID
      WScript.Echo "BucketTable: " & objItem.BucketTable
      WScript.Echo "Category: " & objItem.Category
      WScript.Echo "Date: " & WMIDateStringToDate(objItem.Date)
      WScript.Echo "Debug: " & objItem.Debug
      WScript.Echo "Event: " & objItem.Event
      WScript.Echo "ModuleName: " & objItem.ModuleName
      WScript.Echo "ModuleStamp: " & objItem.ModuleStamp
      WScript.Echo "ModuleVersion: " & objItem.ModuleVersion
      WScript.Echo "Offset: " & objItem.Offset
      WScript.Echo
   Next
Next


Function WMIDateStringToDate(dtmDate)
WScript.Echo dtm: 
	WMIDateStringToDate = CDate(Mid(dtmDate, 5, 2) & "/" & _
	Mid(dtmDate, 7, 2) & "/" & Left(dtmDate, 4) _
	& " " & Mid (dtmDate, 9, 2) & ":" & Mid(dtmDate, 11, 2) & ":" & Mid(dtmDate,13, 2))
End Function

