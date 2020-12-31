Set objShell = CreateObject("Shell.Application")
Set objShellWindows = objShell.Windows

If objShellWindows.Count = 0 Then
    Wscript.Echo "No browser windows are open to the Script Center."
    Wscript.Quit
End If

blnFound = False

For i = 0 to objShellWindows.Count - 1
    Set objIE = objShellWindows.Item(i)
    strURL = objIE.LocationURL
    If InStr(strURL, "http://www.microsoft.com/technet/scriptcenter")Then
        blnFound = True
    End If
Next

If blnFound Then
    Wscript.Echo "At least one browser window is open to the Script Center."
Else
    Wscript.Echo "No browser windows are open to the Script Center."
End If
  


