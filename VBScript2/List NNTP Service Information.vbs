' Description: Returns information about the NNTP service on a computer.


strComputer = "LocalHost"
 
Set objIIS = GetObject("IIS://" & strComputer & "/NNTPSVC")
 
Wscript.Echo "Access Execute: " & objIIS.AccessExecute
Wscript.Echo "Access Flags: " & objIIS.AccessFlags
Wscript.Echo "Access No Physical Directory: " & _
    objIIS.AccessNoPhysicalDir
Wscript.Echo "Access No Remote Execute: " & _
    objIIS.AccessNoRemoteExecute
Wscript.Echo "Access No Remote Read: " & objIIS.AccessNoRemoteRead
Wscript.Echo "Access No Remote Script: " & _
    objIIS.AccessNoRemoteScript
Wscript.Echo "Access No Remote Write: " & _
    objIIS.AccessNoRemoteWrite
Wscript.Echo "Access Read: " & objIIS.AccessRead
Wscript.Echo "Access Script: " & objIIS.AccessScript
Wscript.Echo "Access Source: " & objIIS.AccessSource
Wscript.Echo "Access SSL: " & objIIS.AccessSSL
Wscript.Echo "Access SSL 128: " & objIIS.AccessSSL128
Wscript.Echo "Access SSL Flags: " & objIIS.AccessSSLFlags
Wscript.Echo "Access SSL Map Certificate: " & _
    objIIS.AccessSSLMapCert
Wscript.Echo "Access SSL Negotiate Certificate: " & _
    objIIS.AccessSSLNegotiateCert
Wscript.Echo "Access SSL Require Certificate: " & _
    objIIS.AccessSSLRequireCert
Wscript.Echo "Access Write: " & objIIS.AccessWrite
For Each strACL in objIIS.AdminACLBin
    Wscript.Echo "Admin ACL Bin: " & strACL
Next
Wscript.Echo "Admin Email: " & objIIS.AdminEmail
Wscript.Echo "Admin Name: " & objIIS.AdminName
Wscript.Echo "Allow Clien tPosts: " & objIIS.AllowClientPosts
Wscript.Echo "Allow Control Msgs: " & objIIS.AllowControlMsgs
Wscript.Echo "Allow Feedback Posts: " & objIIS.AllowFeedPosts
Wscript.Echo "Anonymous Password Sync: " & _
    objIIS.AnonymousPasswordSync
Wscript.Echo "Anonymous User Name: " & objIIS.AnonymousUserName
Wscript.Echo "Anonymous User Password: " & objIIS.AnonymousUserPass
Wscript.Echo "Article Time Limit: " & objIIS.ArticleTimeLimit
Wscript.Echo "Authentication Anonymous: " & objIIS.AuthAnonymous
Wscript.Echo "Authentication Basic: " & objIIS.AuthBasic
Wscript.Echo "Authentication Flags: " & objIIS.AuthFlags
Wscript.Echo "Authentication MD5: " & objIIS.AuthMD5
Wscript.Echo "Authentication NTLM: " & objIIS.AuthNTLM
Wscript.Echo "Authentication Passport: " & objIIS.AuthPassport
Wscript.Echo "Az Enable: " & objIIS.AzEnable
Wscript.Echo "Az Scope Name: " & objIIS.AzScopeName
Wscript.Echo "Az Store Name: " & objIIS.AzStoreName
Wscript.Echo "Client Post Hard Limit: " & _
    objIIS.ClientPostHardLimit
Wscript.Echo "Client Post Soft Limit: " & _
    objIIS.ClientPostSoftLimit
Wscript.Echo "Connection Timeout: " & objIIS.ConnectionTimeout
Wscript.Echo "Content Indexed: " & objIIS.ContentIndexed
Wscript.Echo "Default Moderator Domain: " & _
    objIIS.DefaultModeratorDomain
Wscript.Echo "Disable New News: " & objIIS.DisableNewNews
Wscript.Echo "Don't Log: " & objIIS.DontLog
Wscript.Echo "Feed Post Hard Limit: " & objIIS.FeedPostHardLimit
Wscript.Echo "Feed Post Soft Limit: " & objIIS.FeedPostSoftLimit
Wscript.Echo "Feed Report Period: " & objIIS.FeedReportPeriod
Wscript.Echo "Group Var List File: " & objIIS.GroupVarListFile
Wscript.Echo "History Expiration: " & objIIS.HistoryExpiration
Wscript.Echo "Honor Client Msg Ids: " & objIIS.HonorClientMsgIds
Wscript.Echo "Log Ext File Bytes Received: " & _
    objIIS.LogExtFileBytesRecv
Wscript.Echo "Log Ext File Bytes Sent: " & _
    objIIS.LogExtFileBytesSent
Wscript.Echo "Log Ext File Client IP: " & objIIS.LogExtFileClientIp
Wscript.Echo "Log Ext File Computer Name: " & _
    objIIS.LogExtFileComputerName
Wscript.Echo "Log Ext File Cookie: " & objIIS.LogExtFileCookie
Wscript.Echo "Log Ext File Date: " & objIIS.LogExtFileDate
Wscript.Echo "Log Ext File Flags: " & objIIS.LogExtFileFlags
Wscript.Echo "Log Ext File Host: " & objIIS.LogExtFileHost
Wscript.Echo "Log Ext File Http Status: " & _
    objIIS.LogExtFileHttpStatus
Wscript.Echo "Log Ext File Http SubStatus: " & _
    objIIS.LogExtFileHttpSubStatus
Wscript.Echo "Log Ext File Method: " & objIIS.LogExtFileMethod
Wscript.Echo "Log Ext File Protocol Version: " & _
    objIIS.LogExtFileProtocolVersion
Wscript.Echo "Log Ext File Referer: " & objIIS.LogExtFileReferer
Wscript.Echo "Log Ext File Server IP: " & objIIS.LogExtFileServerIp
Wscript.Echo "Log Ext File Server Port: " & _
    objIIS.LogExtFileServerPort
Wscript.Echo "Log Ext File Site Name: " & objIIS.LogExtFileSiteName
Wscript.Echo "Log Ext File Time: " & objIIS.LogExtFileTime
Wscript.Echo "Log Ext File Time Taken: " & objIIS.LogExtFileTimeTaken
Wscript.Echo "Log Ext File URI Query: " & objIIS.LogExtFileUriQuery
Wscript.Echo "Log Ext File URI Stem: " & objIIS.LogExtFileUriStem
Wscript.Echo "Log Ext File User Agent: " & objIIS.LogExtFileUserAgent
Wscript.Echo "Log Ext File User Name: " & objIIS.LogExtFileUserName
Wscript.Echo "Log Ext FileWin32 Status: " & _
    objIIS.LogExtFileWin32Status
Wscript.Echo "Log File Directory: " & objIIS.LogFileDirectory
Wscript.Echo "Log File Period: " & objIIS.LogFilePeriod
Wscript.Echo "Log File Truncate Size: " & objIIS.LogFileTruncateSize
Wscript.Echo "Log Odbc Data Source: " & objIIS.LogOdbcDataSource
Wscript.Echo "Log Odbc Password: " & objIIS.LogOdbcPassword
Wscript.Echo "Log Odbc Table Name: " & objIIS.LogOdbcTableName
Wscript.Echo "Log Odbc User Name: " & objIIS.LogOdbcUserName
Wscript.Echo "Log Plugin Clsid: " & objIIS.LogPluginClsid
Wscript.Echo "Log Type: " & objIIS.LogType
Wscript.Echo "Maximum Bandwidth: " & objIIS.MaxBandwidth
Wscript.Echo "Maximum Connections: " & objIIS.MaxConnections
Wscript.Echo "Maximum Endpoint Connections: " & _
    objIIS.MaxEndpointConnections
Wscript.Echo "Maximum Search Results: " & objIIS.MaxSearchResults
Wscript.Echo "Name: " & objIIS.Name
Wscript.Echo "News Crawler Time: " & objIIS.NewsCrawlerTime
Wscript.Echo "NNTP Command Log Mask: " & objIIS.NntpCommandLogMask
Wscript.Echo "NT Authentication Providers: " & _
    objIIS.NTAuthenticationProviders
Wscript.Echo "Server AutoStart: " & objIIS.ServerAutoStart
Wscript.Echo "Server Comment: " & objIIS.ServerComment
Wscript.Echo "Server Listen Backlog: " & objIIS.ServerListenBacklog
Wscript.Echo "Server Listen Timeout: " & objIIS.ServerListenTimeout
Wscript.Echo "Shutdown Latency: " & objIIS.ShutdownLatency
Wscript.Echo "SMTP Server: " & objIIS.SmtpServer

