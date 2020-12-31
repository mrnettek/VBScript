' Description: Displays information about all the FTP sites on an IIS server.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsFtpServerSetting")
 
For Each objItem in colItems
    Wscript.Echo "Access Execute: " & objItem.AccessExecute
    Wscript.Echo "Access Flags: " & objItem.AccessFlags
    Wscript.Echo "Access No Physical Directory: " & _
        objItem.AccessNoPhysicalDir
    Wscript.Echo "Access No Remote Execute: " & _
        objItem.AccessNoRemoteExecute
    Wscript.Echo "Access No Remote Read: " & _
        objItem.AccessNoRemoteRead
    Wscript.Echo "Access No Remote Script: " & _
        objItem.AccessNoRemoteScript
    Wscript.Echo "Access No Remote Write: " & _
        objItem.AccessNoRemoteWrite
    Wscript.Echo "Access Read: " & objItem.AccessRead
    Wscript.Echo "Access Script: " & objItem.AccessScript
    Wscript.Echo "Access Source: " & objItem.AccessSource
    Wscript.Echo "Access Write: " & objItem.AccessWrite
    Wscript.Echo "AD Connections Password: " & _
        objItem.ADConnectionsPassword
    Wscript.Echo "AD Connections User Name: " & _
        objItem.ADConnectionsUserName
    Wscript.Echo "Admin ACL Bin: " & objItem.AdminACLBin
    Wscript.Echo "Allow Anonymous: " & objItem.AllowAnonymous
    Wscript.Echo "Anonymous Only: " & objItem.AnonymousOnly
    Wscript.Echo "Anonymous Password Sync: " & _
        objItem.AnonymousPasswordSync
    Wscript.Echo "Anonymous User Name: " & _
        objItem.AnonymousUserName
    Wscript.Echo "Anonymous User Password: " & _
        objItem.AnonymousUserPass
    For Each objMessage in objItem.BannerMessage
        Wscript.Echo "Banner Message: " & objMessage
    Next
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Cluster Enabled: " & objItem.ClusterEnabled
    Wscript.Echo "Connection Timeout: " & _
        objItem.ConnectionTimeout
    Wscript.Echo "Default Logon Domain: " & _
        objItem.DefaultLogonDomain
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Disable Socket Pooling: " & _
        objItem.DisableSocketPooling
    Wscript.Echo "Don't Log: " & objItem.DontLog
    Wscript.Echo "Exit Message: " & objItem.ExitMessage
    Wscript.Echo "FTP Directory Browse Show Long Date: " & _
        objItem.FtpDirBrowseShowLongDate
    Wscript.Echo "FTP Log in Utf8: " & objItem.FtpLogInUtf8
    For Each objMessage in objItem.GreetingMessage
        Wscript.Echo "Greeting Message: " & objMessage
    Next
    Wscript.Echo "Log Anonymous: " & objItem.LogAnonymous
    Wscript.Echo "Log Ext File Bytes Received: " & _
        objItem.LogExtFileBytesRecv
    Wscript.Echo "Log Ext File Bytes Sent: " & _
        objItem.LogExtFileBytesSent
    Wscript.Echo "Log Ext File Client IP: " & _
        objItem.LogExtFileClientIp
    Wscript.Echo "Log Ext File Computer Name: " & _
        objItem.LogExtFileComputerName
    Wscript.Echo "Log Ext File Cookie: " & objItem.LogExtFileCookie
    Wscript.Echo "Log Ext File Date: " & objItem.LogExtFileDate
    Wscript.Echo "Log Ext File Flags: " & objItem.LogExtFileFlags
    Wscript.Echo "Log Ext File Host: " & objItem.LogExtFileHost
    Wscript.Echo "Log Ext File Http Status: " & _
        objItem.LogExtFileHttpStatus
    Wscript.Echo "Log Ext File Http SubStatus: " & _
        objItem.LogExtFileHttpSubStatus
    Wscript.Echo "Log Ext File Method: " & objItem.LogExtFileMethod
    Wscript.Echo "Log Ext File Protocol Version: " & _
        objItem.LogExtFileProtocolVersion
    Wscript.Echo "Log Ext File Referer: " & objItem.LogExtFileReferer
    Wscript.Echo "Log Ext File Server IP: " & _
        objItem.LogExtFileServerIp
    Wscript.Echo "Log Ext File Server Port: " & _
        objItem.LogExtFileServerPort
    Wscript.Echo "Log Ext File Site Name: " & _
        objItem.LogExtFileSiteName
    Wscript.Echo "Log Ext File Time: " & objItem.LogExtFileTime
    Wscript.Echo "Log Ext File Time Taken: " & _
        objItem.LogExtFileTimeTaken
    Wscript.Echo "Log Ext File URI Query: " & _
        objItem.LogExtFileUriQuery
    Wscript.Echo "Log Ext File Uri Stem: " & objItem.LogExtFileUriStem
    Wscript.Echo "Log Ext File User Agent: " & _
        objItem.LogExtFileUserAgent
    Wscript.Echo "Log Ext File User Name: " & objItem.LogExtFileUserName
    Wscript.Echo "Log Ext File Win32 Status: " & _
        objItem.LogExtFileWin32Status
    Wscript.Echo "Log File Directory: " & objItem.LogFileDirectory
    Wscript.Echo "Log File Local Time Rollover: " & _
        objItem.LogFileLocaltimeRollover
    Wscript.Echo "Log File Period: " & objItem.LogFilePeriod
    Wscript.Echo "Log File Truncate Size: " & _
        objItem.LogFileTruncateSize
    Wscript.Echo "Log Non-Anonymous: " & objItem.LogNonAnonymous
    Wscript.Echo "Log Odbc Data Source: " & objItem.LogOdbcDataSource
    Wscript.Echo "Log Odbc Password: " & objItem.LogOdbcPassword
    Wscript.Echo "Log Odbc Table Name: " & objItem.LogOdbcTableName
    Wscript.Echo "Log Odbc User Name: " & objItem.LogOdbcUserName
    Wscript.Echo "Log Plugin Clsid: " & objItem.LogPluginClsid
    Wscript.Echo "Log Type: " & objItem.LogType
    Wscript.Echo "Maximum Clients Message: " & _
        objItem.MaxClientsMessage
    Wscript.Echo "Maximum Connections: " & objItem.MaxConnections
    Wscript.Echo "Maximum Endpoint Connections: " & _
        objItem.MaxEndpointConnections
    Wscript.Echo "MS-DOS Directory Output: " & objItem.MSDOSDirOutput
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Realm: " & objItem.Realm
    Wscript.Echo "Server AutoStart: " & objItem.ServerAutoStart
    Wscript.Echo "Server Command: " & objItem.ServerCommand
    Wscript.Echo "Server Comment: " & objItem.ServerComment
    Wscript.Echo "Server ID: " & objItem.ServerID
    Wscript.Echo "Server Listen Backlog: " & _
        objItem.ServerListenBacklog
    Wscript.Echo "Server Listen Timeout: " & _
        objItem.ServerListenTimeout
    Wscript.Echo "Server Size: " & objItem.ServerSize
    Wscript.Echo "Setting ID: " & objItem.SettingID
    Wscript.Echo "User Isolation Mode: " & objItem.UserIsolationMode
    Wscript.Echo "Win32 Error: " & objItem.Win32Error
Next

