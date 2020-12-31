Const ALL_USERS_DESKTOP = &H19&

Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace(ALL_USERS_DESKTOP)
Set objFolderItem = objFolder.ParseName("Accounts Payable Database.lnk")
Set objShellLink = objFolderItem.GetLink

objShellLink.Path = "\\atl-fs-01\accounting\payable.exe"
objShellLink.Save()
  


