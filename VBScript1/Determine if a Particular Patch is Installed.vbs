Set objSession = CreateObject("Microsoft.Update.Session")
Set objSearcher = objSession.CreateUpdateSearcher
Set objResults = objSearcher.Search("Type='Software'")
Set colUpdates = objResults.Updates

For i = 0 to colUpdates.Count - 1
    If colUpdates.Item(i).Title = _
        "Security Update for Windows XP (KB899587)" Then
        If colUpdates.Item(i).IsInstalled <> 0 Then
            Wscript.Echo "This update is installed."
            Wscript.Quit
        Else
            Wscript.Echo "This update is not installed."
            Wscript.Quit
        End If
    End If
Next

Wscript.Echo "This update is not installed."
  


