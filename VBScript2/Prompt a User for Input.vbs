' Description: Sample HTML function that prompts the user for input of some type, and then displays the information entered.


Sub RunScript
    strAnswer = window.prompt("Please enter the domain name.", "fabrikam.com")
    If IsNull(strAnswer) Then
        Msgbox "You clicked the Cancel button"
    Else
        Msgbox "You entered: " & strAnswer
    End If
End Sub

