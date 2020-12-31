' Description: Demonstration script that displays a File Open dialog box (open to the folder C:\Scripts), and then echoes back the name of the selected file.


Set objDialog = CreateObject("UserAccounts.CommonDialog")

objDialog.Filter = "VBScript Scripts|*.vbs|All Files|*.*"
objDialog.FilterIndex = 1
objDialog.InitialDir = "C:\Scripts"
intResult = objDialog.ShowOpen
 
If intResult = 0 Then
    Wscript.Quit
Else
    Wscript.Echo objDialog.FileName
End If

