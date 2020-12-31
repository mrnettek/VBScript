' Description: Returns information about the global metabase NNTP server settings on an IIS server.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsNntpServiceSetting")
 
For Each objItem in colItems
    Wscript.Echo "Access Execute: " & objItem.AccessExecute
    Wscript.Echo "Access Flags: " & objItem.AccessFlags
    Wscript.Echo "Access No Physical Directory: " & _
        objItem.AccessNoPhysicalDir
    Wscript.Echo "Access No Remote Execute: " & _
        objItem.AccessNoRemoteExecute
    Wscript.Echo "Access No Remote Read: " & objItem.AccessNoRemoteRead
    Wscript.Echo "Access No Remote Script: " & _
        objItem.AccessNoRemoteScript
    Wscript.Echo "Access No Remote Write: " & _
        objItem.AccessNoRemoteWrite
    Wscript.Echo "Access Read: " & objItem.AccessRead
    Wscript.Echo "Access Script: " & objItem.AccessScript
    Wscript.Echo "Access Source: " & objItem.AccessSource
    Wscript.Echo "Access SSL: " & objItem.AccessSSL
    Wscript.Echo "Access SSL 128: " & objItem.AccessSSL128
    Wscript.Echo "Access SSL Flags: " & objItem.AccessSSLFlags
    Wscript.Echo "Access SSL Map Certificate: " & _
        objItem.AccessSSLMapCert
    Wscript.Echo "Access SSL Negotiate Certifcate: " & _
        objItem.AccessSSLNegotiateCert
    Wscript.Echo "Access SSL Require Certificate: " & _
        objItem.AccessSSLRequireCert
    Wscript.Echo "Access Write: " & objItem.AccessWrite
    Wscript.Echo "Admin ACL Bin: " & objItem.AdminACLBin
    Wscript.Echo "Admin Email: " & objItem.AdminEmail
    Wscript.Echo "Admin Name: " & objItem.AdminName
    Wscript.Echo "Allow Anonymous: " & objItem.AllowAnonymous
    Wscript.Echo "Allow Client Posts: " & objItem.AllowClientPosts
    Wscript.Echo "Allow Control Msgs: " & objItem.AllowControlMsgs
    Wscript.Echo "Allow Feed Posts: " & objItem.AllowFeedPosts
    Wscript.Echo "Anonymous Password Sync: " & _
        objItem.AnonymousPasswordSync
    Wscript.Echo "Anonymous User Name: " & objItem.AnonymousUserName
    Wscript.Echo "Anonymous User Password: " & objItem.AnonymousUserPass
    Wscript.Echo "Article Time Limit: " & objItem.ArticleTimeLimit
    Wscript.Echo "Authentication Anonymous: " & objItem.AuthAnonymous
    Wscript.Echo "Authentication Basic: " & objItem.AuthBasic
    Wscript.Echo "Authentication Flags: " & objItem.AuthFlags
    Wscript.Echo "Authentication MD5: " & objItem.AuthMD5
    Wscript.Echo "Authentication NTLM: " & objItem.AuthNTLM
    Wscript.Echo "Authentication Passport: " & objItem.AuthPassport
    Wscript.Echo "Az Enable: " & objItem.AzEnable
    Wscript.Echo "Az Scope Name: " & objItem.AzScopeName
    Wscript.Echo "Az Store Name: " & objItem.AzStoreName
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Client Post Hard Limit: " & _
        objItem.ClientPostHardLimit
    Wscript.Echo "Client Post Soft Limit: " & _
        objItem.ClientPostSoftLimit
    Wscript.Echo "Connection Timeout: " & objItem.ConnectionTimeout
    Wscript.Echo "Content Indexed: " & objItem.ContentIndexed
    Wscript.Echo "Default Moderator Domain: " & _
        objItem.DefaultModeratorDomain
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Directory Levels To Scan: " & _
        objItem.DirectoryLevelsToScan
    Wscript.Echo "Disable New News: " & objItem.DisableNewNews
    Wscript.Echo "Don't Log: " & objItem.DontLog
    Wscript.Echo "Feed Post Hard Limit: " & objItem.FeedPostHardLimit
    Wscript.Echo "Feed Post Soft Limit: " & objItem.FeedPostSoftLimit
    Wscript.Echo "Feed Report Period: " & objItem.FeedReportPeriod
    Wscript.Echo "Group Var List File: " & objItem.GroupVarListFile
    Wscript.Echo "History Expiration: " & objItem.HistoryExpiration
    Wscript.Echo "Honor Client Msg Ids: " & objItem.HonorClientMsgIds
    Wscript.Echo "Log Ext File Bytes Received: " & _
        objItem.LogExtFileBytesRecv
    Wscript.Echo "Log Ext File Bytes Sent: " & _
        objItem.LogExtFileBytesSent
    Wscript.Echo "Log Ext File Client IP: " & objItem.LogExtFileClientIp
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
    Wscript.Echo "Log Ext File Server IP: " & objItem.LogExtFileServerIp
    Wscript.Echo "Log Ext File Server Port: " & objItem.LogExtFileServerPort
    Wscript.Echo "Log Ext File Site Name: " & objItem.LogExtFileSiteName
    Wscript.Echo "Log Ext File Time: " & objItem.LogExtFileTime
    Wscript.Echo "Log Ext File Time Taken: " & objItem.LogExtFileTimeTaken
    Wscript.Echo "Log Ext File URI Query: " & objItem.LogExtFileUriQuery
    Wscript.Echo "Log Ext File URI Stem: " & objItem.LogExtFileUriStem
    Wscript.Echo "Log Ext File User Agent: " & objItem.LogExtFileUserAgent
    Wscript.Echo "Log Ext File User Name: " & objItem.LogExtFileUserName
    Wscript.Echo "Log Ext File Win32 Status: " & _
        objItem.LogExtFileWin32Status
    Wscript.Echo "Log File Directory: " & objItem.LogFileDirectory
    Wscript.Echo "Log File Period: " & objItem.LogFilePeriod
    Wscript.Echo "Log File Truncate Size: " & objItem.LogFileTruncateSize
    Wscript.Echo "Log Module List: " & objItem.LogModuleList
    Wscript.Echo "Log Odbc Data Source: " & objItem.LogOdbcDataSource
    Wscript.Echo "Log Odbc Password: " & objItem.LogOdbcPassword
    Wscript.Echo "Log Odbc Table Name: " & objItem.LogOdbcTableName
    Wscript.Echo "Log Odbc User Name: " & objItem.LogOdbcUserName
    Wscript.Echo "Log Plugin Clsid: " & objItem.LogPluginClsid
    Wscript.Echo "Log Type: " & objItem.LogType
    Wscript.Echo "Maximum Bandwidth: " & objItem.MaxBandwidth
    Wscript.Echo "Maximum Connections: " & objItem.MaxConnections
    Wscript.Echo "Maximum EndpointConnections: " & _
        objItem.MaxEndpointConnections
    Wscript.Echo "Maximum SearchResults: " & objItem.MaxSearchResults
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "News Crawler Time: " & objItem.NewsCrawlerTime
    Wscript.Echo "NNTP Command Log Mask: " & objItem.NntpCommandLogMask
    Wscript.Echo "NT Authentication Providers: " & _
        objItem.NTAuthenticationProviders
    Wscript.Echo "Server AutoStart: " & objItem.ServerAutoStart
    Wscript.Echo "Server Comment: " & objItem.ServerComment
    Wscript.Echo "Server Listen Backlog: " & _
        objItem.ServerListenBacklog
    Wscript.Echo "Server Listen Timeout: " & _
        objItem.ServerListenTimeout
    Wscript.Echo "Setting ID: " & objItem.SettingID
    Wscript.Echo "Shutdown Latency: " & objItem.ShutdownLatency
    Wscript.Echo "SMTP Server: " & objItem.SmtpServer
Next

