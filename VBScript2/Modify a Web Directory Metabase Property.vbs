' Description: Demonstration script that modifies a Web directory property (DefaultDocFooter) in the IIS metabase.


strComputer = "LocalHost"
Set objIIS = GetObject _
    ("IIS://" & strComputer & "/W3SVC/1/ROOT/aspnet_client")

objIIS.DefaultDocFooter = "footer.htm"
objIIS.SetInfo

