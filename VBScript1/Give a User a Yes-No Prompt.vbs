intAnswer = _
    Msgbox("Do you want to delete these files?", _
        vbYesNo, "Delete Files")

If intAnswer = vbYes Then
    Msgbox "You answered yes."
Else
    Msgbox "You answered no."
End If
  


