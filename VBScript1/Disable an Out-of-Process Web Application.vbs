' Description: Disables an out-of-process Web application in the W3SVC/2142295254/root/aspnet_client_folder directory.


strComputer = "LocalHost"
Set objIIS = GetObject _
    ("IIS://" & strComputer & "/W3SVC/2142295254/root/aspnet_client_folder")

objIIS.AppDisableRecursive

