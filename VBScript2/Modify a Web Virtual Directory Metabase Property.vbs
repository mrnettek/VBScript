' Description: Demonstration script that modifies a Web virtual directory property (EnableReverseDNS) in the IIS metabase.


strComputer = "LocalHost"
Set objIIS = GetObject _
    ("IIS://" & strComputer & "/W3SVC/1/ROOT/Printers")

objIIS.EnableReverseDns = TRUE
objIIS.SetInfo

