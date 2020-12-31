' Description: Uses the Shell object to move the folder C:\Scripts to D:\Archives. Displays the Copying Files progress dialog as the folder is being moved.


Const FOF_CREATEPROGRESSDLG = &H0&

TargetFolder = "D:\Archive" 
Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.NameSpace(TargetFolder) 
objFolder.MoveHere "C:\Scripts", FOF_CREATEPROGRESSDLG

