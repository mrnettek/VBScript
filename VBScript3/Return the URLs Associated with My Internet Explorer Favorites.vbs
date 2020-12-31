Const FAVORITES = &H6&

Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace(FAVORITES)

For Each objItem in objFolder.Items
    If objItem.IsLink Then
        Set objLink = objItem.GetLink
        Wscript.Echo objItem.Name
        Wscript.Echo objLink.Target
        Wscript.Echo
    End If
Next
  


