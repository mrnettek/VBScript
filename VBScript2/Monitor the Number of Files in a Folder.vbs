strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Do While True
    Set colFileList = objWMIService.ExecQuery _
        ("ASSOCIATORS OF {Win32_Directory.Name='C:\Logs'} Where " _
            & "ResultClass = CIM_DataFile")

    If colFileList.Count >= 100 Then
        Exit Do
    End If

    Wscript.Sleep 60000
Loop

Wscript.Echo "There are at least 100 log files in the target folder."



