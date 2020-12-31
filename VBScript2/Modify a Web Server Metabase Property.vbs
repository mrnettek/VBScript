' Description: Demonstration script that modifies a Web server property (PasswordExpirePrenotifyDays) in the IIS metabase. This script modifies the metabase for a Web server named W3SVC/1.


strComputer = "LocalHost"
Set objIIS = GetObject("IIS://" & strComputer & "/W3SVC/1")

objIIS.PasswordExpirePrenotifyDays = 10
objIIS.SetInfo

