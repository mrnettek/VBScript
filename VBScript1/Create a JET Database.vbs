' Description: Demonstration script that creates a new database named New_db.mdb.


Set objConnection = CreateObject("ADOX.Catalog")

objConnection.Create _
    "Provider = Microsoft.Jet.OLEDB.4.0; " & _
        "Data Source = new_db.mdb"

