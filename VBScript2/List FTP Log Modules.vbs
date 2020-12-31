' Description: Returns a list of all the FTP log modules found on a server.


strComputer = "LocalHost"
Set objIIS = GetObject("IIS://" & strComputer & "/MSFTPSVC/Info")

Wscript.Echo "Log Module List: " & objIIS.LogModuleList

