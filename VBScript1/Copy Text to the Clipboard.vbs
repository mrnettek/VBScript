' Description: Sample HTML script for copying text found in a text area named BasicTextArea to the clipboard.


Sub RunScript
    strCopy = BasicTextArea.Value
    document.parentwindow.clipboardData.SetData "text", strCopy
End Sub

