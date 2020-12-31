' Description: Sample HTML function that lists all the items that were selected in a multi-select listbox.


Sub RunScript
    For i = 0 to (Dropdown1.Options.Length - 1)
        If (Dropdown1.Options(i).Selected) Then
            strComputer = strComputer & Dropdown1.Options(i).Value & vbcrlf
        End If
    Next
    Msgbox strComputer
End Sub

