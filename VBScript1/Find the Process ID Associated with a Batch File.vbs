strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * From Win32_Process")

For Each objItem in colItems
    If InStr(objItem.CommandLine, ".bat") Or InStr(objItem.CommandLine, ".cmd") Then
        Wscript.Echo "Batch file: " & objItem.CommandLine
        Wscript.Echo "Process ID: " & objItem.ProcessID
    End If
Next
  


