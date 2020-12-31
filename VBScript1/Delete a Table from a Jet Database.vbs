Set objCatalog = CreateObject("ADOX.Catalog")
Set objConnection = CreateObject("ADODB.Connection")

objConnection.Open _
    "Provider= Microsoft.Jet.OLEDB.4.0; " & _
        "Data Source=inventory.mdb" 

Set objCatalog.ActiveConnection = objConnection
objCatalog.Tables.Delete "HardwareBackup"

objConnection.Close
  


