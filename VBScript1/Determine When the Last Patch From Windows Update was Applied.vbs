Set objSession = CreateObject("Microsoft.Update.Session")
Set objSearcher = objSession.CreateUpdateSearcher

Set colHistory = objSearcher.QueryHistory(1, 1)

For Each objEntry in colHistory
    Wscript.Echo "Title: " & objEntry.Title
    Wscript.Echo "Update application date: " & objEntry.Date
Next
  


