Const NETWORK_CONNECTIONS = &H31&

Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace(NETWORK_CONNECTIONS)

Set colItems = objFolder.Items
For Each objItem in colItems
    If objItem.Name = "Local Area Connection 2" Then
        objItem.Name = "Home Office Connection"
    End If
Next
  


