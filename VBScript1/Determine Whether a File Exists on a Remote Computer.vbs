strComputer = "atl-fs-01"

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colFiles = objWMIService. _
    ExecQuery("Select * From CIM_DataFile Where Name = 'C:\\Scripts\\Test.vbs'")

If colFiles.Count = 0 Then
    Wscript.Echo "The file does not exist on the remote computer."
Else
    Wscript.Echo "The file exists on the remote computer."
End If
  


