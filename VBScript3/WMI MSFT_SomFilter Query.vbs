On Error Resume Next

Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

arrComputers = Array(".")
For Each strComputer In arrComputers
   WScript.Echo
   WScript.Echo "=========================================="
   WScript.Echo "Computer: " & strComputer
   WScript.Echo "=========================================="

   Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\Policy")
   Set colItems = objWMIService.ExecQuery("SELECT * FROM MSFT_SomFilter", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Author: " & objItem.Author
      WScript.Echo "ChangeDate: " & WMIDateStringToDate(objItem.ChangeDate)
      WScript.Echo "CreationDate: " & WMIDateStringToDate(objItem.CreationDate)
      WScript.Echo "Description: " & objItem.Description
      WScript.Echo "Domain: " & objItem.Domain
      WScript.Echo "ID: " & objItem.ID
      WScript.Echo "Name: " & objItem.Name
      strRules = Join(objItem.Rules, ",")
         WScript.Echo "Rules: " & strRules
      WScript.Echo "SourceOrganization: " & objItem.SourceOrganization
      WScript.Echo
   Next
Next


Function WMIDateStringToDate(dtmDate)
WScript.Echo dtm: 
	WMIDateStringToDate = CDate(Mid(dtmDate, 5, 2) & "/" & _
	Mid(dtmDate, 7, 2) & "/" & Left(dtmDate, 4) _
	& " " & Mid (dtmDate, 9, 2) & ":" & Mid(dtmDate, 11, 2) & ":" & Mid(dtmDate,13, 2))
End Function

