Const FOF_CREATEPROGRESSDLG = &H0&

strTargetFolder = "D:\Scripts" 

Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.NameSpace(strTargetFolder) 

objFolder.CopyHere "C:\Scripts\*.*", FOF_CREATEPROGRESSDLG

