' Description: Demonstration script that modifies an FTP server property (DontLog) in the IIS metabase for an FTP server named MSFTPSVC/1.


On Error Resume Next
 
strComputer = "LocalHost"
Set objIIS = GetObject("IIS://" & strComputer & "/MSFTPSVC/1")

objIIS.DontLog = TRUE
objIIS.SetInfo

