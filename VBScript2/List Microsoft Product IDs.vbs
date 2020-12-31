' Description: Returns the Product IDs for Microsoft software products installed on a computer.


Set objMSInfo = CreateObject("MsPIDinfo.MsPID")
colMSApps = objMSInfo.GetPIDInfo()

For Each strApp in colMSApps
    Wscript.Echo strApp
Next

