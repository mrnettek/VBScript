On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_NTLogEvent",,48)
For Each objItem in colItems
    Wscript.Echo "Category: " & objItem.Category
    Wscript.Echo "CategoryString: " & objItem.CategoryString
    Wscript.Echo "ComputerName: " & objItem.ComputerName
    Wscript.Echo "Data: " & objItem.Data
    Wscript.Echo "EventCode: " & objItem.EventCode
    Wscript.Echo "EventIdentifier: " & objItem.EventIdentifier
    Wscript.Echo "EventType: " & objItem.EventType
    Wscript.Echo "InsertionStrings: " & objItem.InsertionStrings
    Wscript.Echo "Logfile: " & objItem.Logfile
    Wscript.Echo "Message: " & objItem.Message
    Wscript.Echo "RecordNumber: " & objItem.RecordNumber
    Wscript.Echo "SourceName: " & objItem.SourceName
    Wscript.Echo "TimeGenerated: " & objItem.TimeGenerated
    Wscript.Echo "TimeWritten: " & objItem.TimeWritten
    Wscript.Echo "Type: " & objItem.Type
    Wscript.Echo "User: " & objItem.User
Next

