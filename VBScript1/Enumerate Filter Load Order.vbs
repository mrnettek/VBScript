' Description: Enumerates the filter load order on an IIS server.


strComputer = "LocalHost"
Set objIIS = GetObject("IIS://" & strComputer & "/W3SVC/Filters")

Wscript.Echo "Filter Load Order: " & objIIS.FilterLoadOrder

