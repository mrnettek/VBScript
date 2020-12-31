' Description: Returns information about the SMTP service on a computer.


strComputer = "LocalHost"
 
Set objIIS = GetObject("IIS://" & strComputer & "/SMTPSVC")
 
Wscript.Echo "Access Flags: " & objIIS.AccessFlags
Wscript.Echo "Disable Socket Pooling: " & _
    objIIS.DisableSocketPooling
Wscript.Echo "Allow Anonymous: " & objIIS.AllowAnonymous
Wscript.Echo "Don't Log: " & objIIS.DontLog
Wscript.Echo "Anonymous Only: " & objIIS.AnonymousOnly
Wscript.Echo "Anonymous Password Sync: " & _
    objIIS.AnonymousPasswordSync
Wscript.Echo "Anonymous User Name: " & objIIS.AnonymousUserName
Wscript.Echo "Anonymous User Password: " & _
    objIIS.AnonymousUserPass
Wscript.Echo "Connection Timeout: " & objIIS.ConnectionTimeout
Wscript.Echo "Log Ext File Flags: " & objIIS.LogExtFileFlags
Wscript.Echo "Log ODBC Data Source: " & objIIS.LogOdbcDataSource
Wscript.Echo "Log ODBC Password: " & objIIS.LogOdbcPassword
Wscript.Echo "Log File Directory: " & objIIS.LogFileDirectory
Wscript.Echo "Log ODBC Table Name: " & objIIS.LogOdbcTableName
Wscript.Echo "Log ODBC User Name: " & objIIS.LogOdbcUserName
Wscript.Echo "Log File Period: " & objIIS.LogFilePeriod
Wscript.Echo "Log Plugin Clsid: " & objIIS.LogPluginClsid
Wscript.Echo "Log File Truncate Size: " & _
    objIIS.LogFileTruncateSize
Wscript.Echo "Log Type: " & objIIS.LogType
Wscript.Echo "Maximum Connections: " & objIIS.MaxConnections
Wscript.Echo "Server Comment: " & objIIS.ServerComment
Wscript.Echo "Maximum Endpoint Connections: " & _
    objIIS.MaxEndpointConnections
Wscript.Echo "Server Listen Timeout: " & _
    objIIS.ServerListenTimeout
Wscript.Echo "Realm: " & objIIS.Realm
Wscript.Echo "Server AutoStart: " & objIIS.ServerAutoStart

