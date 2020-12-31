On Error Resume Next

Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

arrComputers = Array(".")
For Each strComputer In arrComputers
   WScript.Echo
   WScript.Echo "=========================================="
   WScript.Echo "Computer: " & strComputer
   WScript.Echo "=========================================="

   Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2\Applications\MicrosoftIE")
   Set colItems = objWMIService.ExecQuery("SELECT * FROM MicrosoftIE_FileVersion", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Caption: " & objItem.Caption
      WScript.Echo "Company: " & objItem.Company
      WScript.Echo "Date: " & WMIDateStringToDate(objItem.Date)
      WScript.Echo "Description: " & objItem.Description
      WScript.Echo "File: " & objItem.File
      WScript.Echo "Path: " & objItem.Path
      WScript.Echo "SettingID: " & objItem.SettingID
      WScript.Echo "Size: " & objItem.Size
      WScript.Echo "Version: " & objItem.Version
      WScript.Echo
   Next
Next


Function WMIDateStringToDate(dtmDate)
WScript.Echo dtm: 
	WMIDateStringToDate = CDate(Mid(dtmDate, 5, 2) & "/" & _
	Mid(dtmDate, 7, 2) & "/" & Left(dtmDate, 4) _
	& " " & Mid (dtmDate, 9, 2) & ":" & Mid(dtmDate, 11, 2) & ":" & Mid(dtmDate,13, 2))
End Function

