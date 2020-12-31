' Description: Lists the properties of an FTP server named MSFTPSVC/1.


On Error Resume Next

strComputer = "LocalHost"
Set objServer = GetObject("IIS://" & strComputer & "/MSFTPSVC/1")
 
Wscript.Echo "Access Flags" & objServer.AccessFlags
script.Echo "Connection Timeout: " & objServer.ConnectionTimeout
Wscript.Echo "Default Logon Domain: " & objServer.DefaultLogonDomain
Wscript.Echo "Allow Anonymous: " & objServer.AllowAnonymous
Wscript.Echo "Disable SOcket Pooling: " & _
    objServer.DisableSocketPooling
Wscript.Echo "Anonymous Only: " & objServer.AnonymousOnly
Wscript.Echo "Don't Log: " & objServer.DontLog
Wscript.Echo "Anonymous Password Synch: " & _
    objServer.AnonymousPasswordSync
Wscript.Echo "Exit Message: " & objServer.ExitMessage
Wscript.Echo "Anonymous User Name: " & objServer.AnonymousUserName
Wscript.Echo "FTP Directory Browse Show Long Date: " & _
    objServer.FtpDirBrowseShowLongDate
Wscript.Echo "Anonymous User Password: " & _
    objServer.AnonymousUserPass
For Each strMessage in objServer.GreetingMessage
    Wscript.Echo "Greeting Message: " & strMessage
Next
Wscript.Echo "Log Anonymous: " & objServer.LogAnonymous
Wscript.Echo "Log Ext File Flags: " & objServer.LogExtFileFlags
Wscript.Echo "Log File Directory: " & objServer.LogFileDirectory
Wscript.Echo "Log File Local Time Rollover: " & _
    objServer.LogFileLocaltimeRollover
Wscript.Echo "Log File Period: " & objServer.LogFilePeriod
Wscript.Echo "Log File Truncate Size: " & _
    objServer.LogFileTruncateSize
Wscript.Echo "Log Non-Anonymous: " & objServer.LogNonAnonymous
Wscript.Echo "Log ODBC Data Source: " & _
    objServer.LogOdbcDataSource
Wscript.Echo "Log ODBC Password: " & objServer.LogOdbcPassword
Wscript.Echo "Log ODBC Table Name: " & objServer.LogOdbcTableName
Wscript.Echo "Log ODNC User Name: " & objServer.LogOdbcUserName
Wscript.Echo "Log Plugin Clsid: " & objServer.LogPluginClsid
Wscript.Echo "Log Type: " & objServer.LogType  
Wscript.Echo "Maximum Client Message: " & _
    objServer.MaxClientsMessage
Wscript.Echo "Maximum Connections: " & objServer.MaxConnections
Wscript.Echo "Server Comment: " & objServer.ServerComment
Wscript.Echo "Maximum Endpoint Connections: " & _
    objServer.MaxEndpointConnections
Wscript.Echo "Server Listen backlog: " & _
    objServer.ServerListenBacklog
Wscript.Echo "MS-DOS Directory Output: " & _
    objServer.MSDOSDirOutput
Wscript.Echo "Server Listen Timeout: " & _
    objServer.ServerListenTimeout
Wscript.Echo "Realm: " & objServer.Realm
Wscript.Echo "Server Size: " & objServer.ServerSize
Wscript.Echo "Server Autostart: " & objServer.ServerAutoStart
Wscript.Echo "Server State: " & objServer.ServerState

