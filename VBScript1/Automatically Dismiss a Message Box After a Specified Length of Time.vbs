Const wshYes = 6
Const wshNo = 7
Const wshYesNoDialog = 4
Const wshQuestionMark = 32

Set objShell = CreateObject("Wscript.Shell")

intReturn = objShell.Popup("Do you want to delete this file?", _
    10, "Delete File", wshYesNoDialog + wshQuestionMark)

If intReturn = wshYes Then
    Wscript.Echo "You clicked the Yes button."
ElseIf intReturn = wshNo Then
    Wscript.Echo "You clicked the No button."
Else
    Wscript.Echo "The popup timed out."
End If
  


