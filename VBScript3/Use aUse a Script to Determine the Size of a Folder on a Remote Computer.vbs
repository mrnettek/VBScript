strComputer = "atl-ws-01"
Set objWMIService = GetObject _
    ("winmgmts:" & "!\\" & strComputer & "\root\cimv2")
Set colFolders = objWMIService.ExecQuery _
    ("Select * from Win32_Directory where Name = " _
        & "'c:\\scripts'")
For Each objFolder in colFolders
    Wscript.Echo objFolder.FileSize
Next
  


