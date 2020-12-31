Const olFolderInbox = 6

strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * From Win32_Process Where Name = 'outlook.exe'")

If colItems.Count = 0 Then
    Set objOutlook = CreateObject("Outlook.Application")
    Set objNamespace = objOutlook.GetNamespace("MAPI")
    objNamespace.Logon "Default Outlook Profile",, False, True    
    Set objFolder = objNamespace.GetDefaultFolder(olFolderInbox)
    objFolder.Display
End If

