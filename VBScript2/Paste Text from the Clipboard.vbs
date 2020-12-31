' Description: Sample HTML function that pastes data from the clipboard into a SPAN named DataArea.


Sub RunScript
    DataArea.InnerHTML = document.parentwindow.clipboardData.GetData("text")
End Sub

