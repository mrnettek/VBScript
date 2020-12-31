Set objShell = CreateObject("Wscript.Shell")

Set objWord = CreateObject("Word.Application")
Set colTasks = objWord.Tasks
i = 0

For Each objTask in colTasks
    strName = LCase(objTask.Name)
    If Instr(strName, "inventory.xls") Then
        i = 1
    End If
Next

strCmdLine = "excel.exe " & chr(34) & "C:\Scripts\Inventory.xls" & chr(34)

If i = 0 Then
    objShell.Run strCmdLine, 3
End If

objWord.Quit
  


