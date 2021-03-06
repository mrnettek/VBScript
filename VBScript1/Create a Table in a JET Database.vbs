' Description: Demonstration script that creates a new table in a database named New_db.mdb.


Set objConnection = CreateObject("ADODB.Connection")

objConnection.Open _
    "Provider= Microsoft.Jet.OLEDB.4.0; " & _
        "Data Source=new_db.mdb" 

objConnection.Execute "CREATE TABLE EventTable(" & _
    "EventKey COUNTER ," & _
    "Category TEXT(50) ," & _
    "ComputerName TEXT(50) ," & _
    "EventCode INTEGER ," & _
    "RecordNumber INTEGER ," & _
    "SourceName TEXT(50) ," & _
    "TimeWritten DATETIME ," & _
    "UserName TEXT(50) ," & _
    "EventType TEXT(50) ," & _
    "Logfile TEXT(50) ," & _
    "Message MEMO)" 

objConnection.Close

