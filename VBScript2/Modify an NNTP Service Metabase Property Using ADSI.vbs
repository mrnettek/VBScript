' Description: Demonstration script that modifies an NNTP service property (MaxConnections) in the IIS metabase.


strComputer = "LocalHost"
 
Set objIIS = GetObject("IIS://" & strComputer & "/NNTPSVC")
objIIS.MaxConnections = 5000
objIIS.SetInfo

