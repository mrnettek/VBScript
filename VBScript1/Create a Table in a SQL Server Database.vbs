strComputer = "atl-sql-01" 

Set objConnection = CreateObject("ADODB.Connection")

objConnection.Open _
    "Provider=SQLOLEDB;Data Source=" & strComputer & ";" & _
        "Trusted_Connection=Yes;Initial Catalog=Master" 

objConnection.Execute "CREATE TABLE TestTable (UserName TEXT,TotalAmount INTEGER)"
  


