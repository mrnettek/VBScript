' Description: Lists the properties of the FTP service as configured on a server.


strComputer = "LocalHost"
Set objIIS = GetObject("IIS://" & strComputer & "/MSFTPSVC")
 
Wscript.Echo "Access Flags: " & objIIS.AccessFlags
Wscript.Echo "Directory levels to Scan: " & objIIS.DirectoryLevelsToScan
Wscript.Echo "Disable Socket Pooling: " & objIIS.DisableSocketPooling
Wscript.Echo "Allow Anonymous: " & objIIS.AllowAnonymous
Wscript.Echo "Don't Log: " & objIIS.DontLog
Wscript.Echo "Anonymous Only: " & objIIS.AnonymousOnly
Wscript.Echo "Exit Message: " & objIIS.ExitMessage
Wscript.Echo "Anonymous Password Sync: " & objIIS.AnonymousPasswordSync
Wscript.Echo "FTP Directory Browse Show Long Date: " & _
    objIIS.FtpDirBrowseShowLongDate
Wscript.Echo "Anonymous User Name: " & objIIS.AnonymousUserName
For Each strGreeting in objIIS.GreetingMessage
    Wscript.Echo "Greeting Message: " & strGreeting
Next
Wscript.Echo "Anonymous User Password: " & objIIS.AnonymousUserPass
Wscript.Echo "Connection Timeout: " & objIIS.ConnectionTimeout
Wscript.Echo "Log Ext File Flags: " & objIIS.LogExtFileFlags
Wscript.Echo "Log ODBC Data Source: " & objIIS.LogOdbcDataSource
Wscript.Echo "Log Anonymous: " & objIIS.LogAnonymous
Wscript.Echo "Log ODBC Password: " & objIIS.LogOdbcPassword
Wscript.Echo "Log File Directory: " & objIIS.LogFileDirectory
Wscript.Echo "Log ODBC Table Name: " & objIIS.LogOdbcTableName
Wscript.Echo "Log File Local Time Rollover: " & _
    objIIS.LogFileLocaltimeRollover
Wscript.Echo "Log ODBC User Name: " & objIIS.LogOdbcUserName
Wscript.Echo "Log File Period: " & objIIS.LogFilePeriod
Wscript.Echo "Log Plugin Clsid: " & objIIS.LogPluginClsid
Wscript.Echo "Log File Truncate Size: " & objIIS.LogFileTruncateSize
Wscript.Echo "Log Type: " & objIIS.LogType
Wscript.Echo "Log Non-Anonymous Message: " & objIIS.LogNonAnonymous
Wscript.Echo "Maximum Clients Message: " & objIIS.MaxClientsMessage
Wscript.Echo "Maximum Connections: " & objIIS.MaxConnections
Wscript.Echo "Server Comment: " & objIIS.ServerComment
Wscript.Echo "Maximum Endpoint Connections: " & objIIS.MaxEndpointConnections
Wscript.Echo "Server Listen Backlog: " & objIIS.ServerListenBacklog
Wscript.Echo "MS-DOS Directory Output: " & objIIS.MSDOSDirOutput
Wscript.Echo "Server Listen Timeout: " & objIIS.ServerListenTimeout
Wscript.Echo "Realm: " & objIIS.Realm
Wscript.Echo "Server Size: " & objIIS.ServerSize
Wscript.Echo "Server AutoStart: " & objIIS.ServerAutoStart

