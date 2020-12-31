strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * From Win32_Product Where Name = 'QuickTime'")

If colItems.Count = 0 Then
    Wscript.Echo "QuickTime is not installed on this computer."
Else
    For Each objItem in colItems
        Wscript.Echo "QuickTime version: " & objItem.Version
    Next
End If
  


