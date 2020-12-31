' Description: Demonstration script that modifies an SMTP service property (ServerComment) in the IIS metabase.


strComputer = "LocalHost"
 
Set objIIS = GetObject("IIS://" & strComputer & "/SMTPSVC")
objIIS.ServerComment = "This is an internal SMTP server."
objIIS.SetInfo

