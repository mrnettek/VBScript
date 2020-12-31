' Description: Continues a paused FTP server named MSFTPSVC/1.


strComputer = "LocalHost"
Set objIIS = GetObject("IIS://" & strComputer & "/MSFTPSVC/1")

objIIS.Continue

