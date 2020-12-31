Set objShell = CreateObject("WScript.Shell")

Set objWshScriptExec = objShell.Exec("ping 192.168.1.1")

Set objStdOut = objWshScriptExec.StdOut

Do Until objStdOut.AtEndOfStream
    strLine = objStdOut.ReadLine
    If Len(strLine) > 2 Then
        WScript.Echo Now & " -- " & strLine
    Else
        Wscript.Echo strLine
    End If
Loop
  


