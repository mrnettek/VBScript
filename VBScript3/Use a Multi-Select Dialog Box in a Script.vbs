Set objDialog = CreateObject("UserAccounts.CommonDialog")
objDialog.Filter = "VBScript Scripts|*.vbs|All Files|*.*"
objDialog.Flags = &H0200
objDialog.FilterIndex = 1
objDialog.InitialDir = "C:\Scripts"
intResult = objDialog.ShowOpen

If intResult = 0 Then
    Wscript.Quit
Else
    Wscript.Echo objDialog.FileName
End If
  


