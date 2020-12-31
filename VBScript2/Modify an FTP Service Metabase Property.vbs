' Description: Demonstration script that modifies an FTP service property (AllowAnonymous) in the IIS metabase.


strComputer = "LocalHost"
Set objIIS = GetObject("IIS://" & strComputer & "/MSFTPSVC")

objIIS.AllowAnonymous = FALSE
objIIS.SetInfo

