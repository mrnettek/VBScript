' Description: Demonstration script that modifies a Web service property (ConnectionTimeout) in the IIS metabase.


strComputer = "LocalHost"
Set objIIS = GetObject("IIS://" & strComputer & "/W3SVC")
objIIS.ConnectionTimeout = 60
objIIS.SetInfo

