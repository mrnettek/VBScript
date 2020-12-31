' Description: Continues a paused Web server named W3SVC/2142295254 on an IIS server.


strComputer = "LocalHost"
Set objIIS = GetObject("IIS://" & strComputer & "/W3SVC/2142295254")

objIIS.Continue

