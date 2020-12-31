' Description: Displays information about all the Web sites on an IIS server.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery("Select * from IIsWebServerSetting")
 
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
    Wscript.Echo "Access No Remote Write: " & objItem.AccessNoRemoteWrite
    Wscript.Echo "Access Read: " & objItem.AccessRead
    Wscript.Echo "Access Script: " & objItem.AccessScript
    Wscript.Echo "Access Source: " & objItem.AccessSource
    Wscript.Echo "Access SSL: " & objItem.AccessSSL
    Wscript.Echo "Access SSL 128: " & objItem.AccessSSL128
    Wscript.Echo "Access SSL Flags: " & objItem.AccessSSLFlags
    Wscript.Echo "Access SSL Map Certificate: " & _
        objItem.AccessSSLMapCert
    Wscript.Echo "Access SSL Negotiate Certificate: " & _
        objItem.AccessSSLNegotiateCert
    Wscript.Echo "Access SSL Require Certificate: " & _
        objItem.AccessSSLRequireCert
    Wscript.Echo "Access Write: " & objItem.AccessWrite
    Wscript.Echo "Admin ACL Bin: " & objItem.AdminACLBin
    Wscript.Echo "Allow Keep Alive: " & objItem.AllowKeepAlive
    Wscript.Echo "Allow Path Information For Script Mappings: " & _
        objItem.AllowPathInfoForScriptMappings
    Wscript.Echo "Anonymous Password Sync: " & _
        objItem.AnonymousPasswordSync
    Wscript.Echo "Anonymous User Name: " & objItem.AnonymousUserName
    Wscript.Echo "Anonymous User Password: " & objItem.AnonymousUserPass
    Wscript.Echo "Applocation Allow Client Debug: " & _
        objItem.AppAllowClientDebug
    Wscript.Echo "Application Allow Debugging: " & _
        objItem.AppAllowDebugging
    Wscript.Echo "Application Friendly Name: " & objItem.AppFriendlyName
    Wscript.Echo "Application Oop Recover Limit: " & _
        objItem.AppOopRecoverLimit
    Wscript.Echo "Application Pool Id: " & objItem.AppPoolId
    Wscript.Echo "Application Wam Clsid: " & objItem.AppWamClsid
    Wscript.Echo "ASP Allow Out Of Process Components: " & _
        objItem.AspAllowOutOfProcComponents
    Wscript.Echo "ASP Allow Session State: " & _
        objItem.AspAllowSessionState
    Wscript.Echo "ASP Application Service Flags: " & _
        objItem.AspAppServiceFlags
    Wscript.Echo "ASP Buffering Limit: " & objItem.AspBufferingLimit
    Wscript.Echo "ASP Buffering On: " & objItem.AspBufferingOn
    Wscript.Echo "ASP Calc Line Number: " & objItem.AspCalcLineNumber
    Wscript.Echo "ASP Codepage: " & objItem.AspCodepage
    Wscript.Echo "ASP Disk Template Cache Directory: " & _
        objItem.AspDiskTemplateCacheDirectory
    Wscript.Echo "ASP Enable Application Restart: " & _
        objItem.AspEnableApplicationRestart
    Wscript.Echo "ASP Enable ASP Html Fallback: " & _
        objItem.AspEnableAspHtmlFallback
    Wscript.Echo "ASP Enable Chunked Encoding: " & o_
        bjItem.AspEnableChunkedEncoding
    Wscript.Echo "ASP Enable Parent Paths: " & _
        objItem.AspEnableParentPaths
    Wscript.Echo "ASP Enable Sxs: " & objItem.AspEnableSxs
    Wscript.Echo "ASP Enable Tracker: " & objItem.AspEnableTracker
    Wscript.Echo "ASP Enable Typelib Cache: " & _
        objItem.AspEnableTypelibCache
    Wscript.Echo "ASP Errors To NT Log: " & objItem.AspErrorsToNTLog
    Wscript.Echo "ASP Exception Catch Enable: " & _
        objItem.AspExceptionCatchEnable
    Wscript.Echo "ASP Execute In MTA: " & objItem.AspExecuteInMTA
    Wscript.Echo "ASP Keep Session ID Secure: " & _
        objItem.AspKeepSessionIDSecure
    Wscript.Echo "ASP LCID: " & objItem.AspLCID
    Wscript.Echo "ASP Log Error Requests: " & _
        objItem.AspLogErrorRequests
    Wscript.Echo "ASP Maximum Disk Template Cache Files: " & _
        objItem.AspMaxDiskTemplateCacheFiles
    Wscript.Echo "ASP Maximum Request Entity Allowed: " & _
        objItem.AspMaxRequestEntityAllowed
    Wscript.Echo "ASP Partition ID: " & objItem.AspPartitionID
    Wscript.Echo "ASP Processor Thread Maximum: " & _
        objItem.AspProcessorThreadMax
    Wscript.Echo "ASP Queue Connection Test Time: " & _
        objItem.AspQueueConnectionTestTime
    Wscript.Echo "ASP Queue Timeout: " & objItem.AspQueueTimeout
    Wscript.Echo "ASP Request Queue Maximum: " & _
        objItem.AspRequestQueueMax
    Wscript.Echo "ASP Run OnEnd Anonymously: " & _
        objItem.AspRunOnEndAnonymously
    Wscript.Echo "ASP Script Engine Cache Maximum: " & _
        objItem.AspScriptEngineCacheMax
    Wscript.Echo "ASP Script Error Message: " & _
        objItem.AspScriptErrorMessage
    Wscript.Echo "ASP Script Error Sent To Browser: " & _
        objItem.AspScriptErrorSentToBrowser
    Wscript.Echo "ASP Script File Cache Size: " & _
        objItem.AspScriptFileCacheSize
    Wscript.Echo "ASP Script Language: " & objItem.AspScriptLanguage
    Wscript.Echo "ASP Script Timeout: " & objItem.AspScriptTimeout
    Wscript.Echo "ASP Session Maximum: " & objItem.AspSessionMax
    Wscript.Echo "ASP Session Timeout: " & objItem.AspSessionTimeout
    Wscript.Echo "ASP Sxs Name: " & objItem.AspSxsName
    Wscript.Echo "ASP Track Threading Model: " & _
        objItem.AspTrackThreadingModel
    Wscript.Echo "ASP Use Partition: " & objItem.AspUsePartition
    Wscript.Echo "Authentication Advanced Notify Disable: " & _
        objItem.AuthAdvNotifyDisable
    Wscript.Echo "Authentication Anonymous: " & objItem.AuthAnonymous
    Wscript.Echo "Authentication Basic: " & objItem.AuthBasic
    Wscript.Echo "Authentication Change Disable: " & _
        objItem.AuthChangeDisable
    Wscript.Echo "Authentication Change Unsecure: " & _
        objItem.AuthChangeUnsecure
    Wscript.Echo "Authentication Flags: " & objItem.AuthFlags
    Wscript.Echo "Authentication MD5: " & objItem.AuthMD5
    Wscript.Echo "Authentication NTLM: " & objItem.AuthNTLM
    Wscript.Echo "Authentication Passport: " & objItem.AuthPassport
    Wscript.Echo "Authentication Persistence: " & objItem.AuthPersistence
    Wscript.Echo "Authentication PersistSingleRequest: " & _
        objItem.AuthPersistSingleRequest
    Wscript.Echo "Az Enable: " & objItem.AzEnable
    Wscript.Echo "Az Impersonation Level: " & objItem.AzImpersonationLevel
    Wscript.Echo "Az Scope Name: " & objItem.AzScopeName
    Wscript.Echo "Az Store Name: " & objItem.AzStoreName
    Wscript.Echo "Cache Control Custom: " & objItem.CacheControlCustom
    Wscript.Echo "Cache Control Maximum Age: " & objItem.CacheControlMaxAge
    Wscript.Echo "Cache Control No Cache: " & objItem.CacheControlNoCache
    Wscript.Echo "Cache ISAPI: " & objItem.CacheISAPI
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Certifcate Check Mode: " & objItem.CertCheckMode
    Wscript.Echo "CGI Timeout: " & objItem.CGITimeout
    Wscript.Echo "Cluster Enabled: " & objItem.ClusterEnabled
    Wscript.Echo "Connection Timeout: " & objItem.ConnectionTimeout
    Wscript.Echo "Content Indexed: " & objItem.ContentIndexed
    Wscript.Echo "Create CGI With New Console: " & _
        objItem.CreateCGIWithNewConsole
    Wscript.Echo "Create Process As User: " & objItem.CreateProcessAsUser
    Wscript.Echo "Default Doc: " & objItem.DefaultDoc
    Wscript.Echo "Default Doc Footer: " & objItem.DefaultDocFooter
    Wscript.Echo "Default Logon Domain: " & objItem.DefaultLogonDomain
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Directory Browse Flags: " & objItem.DirBrowseFlags
    Wscript.Echo "Directory Browse Show Date: " & objItem.DirBrowseShowDate
    Wscript.Echo "Directory Browse Show Extension: " & _
        objItem.DirBrowseShowExtension
    Wscript.Echo "Directory Browse Show Long Date: " & _
        objItem.DirBrowseShowLongDate
    Wscript.Echo "Directory Browse Show Size: " & objItem.DirBrowseShowSize
    Wscript.Echo "Directory Browse Show Time: " & objItem.DirBrowseShowTime
    Wscript.Echo "Disable Socket Pooling: " & objItem.DisableSocketPooling
    Wscript.Echo "Disable Static File Cache: " & _
        objItem.DisableStaticFileCache
    Wscript.Echo "Do Dynamic Compression: " & objItem.DoDynamicCompression
    Wscript.Echo "Don't Log: " & objItem.DontLog
    Wscript.Echo "Do Static Compression: " & objItem.DoStaticCompression
    Wscript.Echo "Enable Default Doc: " & objItem.EnableDefaultDoc
    Wscript.Echo "Enable Dir Browsing: " & objItem.EnableDirBrowsing
    Wscript.Echo "Enable Doc Footer: " & objItem.EnableDocFooter
    Wscript.Echo "Enable Reverse DNS: " & objItem.EnableReverseDns
    Wscript.Echo "FrontPage Web: " & objItem.FrontPageWeb
    Wscript.Echo "Http Expires: " & objItem.HttpExpires
    Wscript.Echo "Log Ext File Bytes Received: " & _
        objItem.LogExtFileBytesRecv
    Wscript.Echo "Log Ext File Bytes Sent: " & objItem.LogExtFileBytesSent
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
    Wscript.Echo "Log Ext File Win32 Status: " & objItem.LogExtFileWin32Status
    Wscript.Echo "Log File Directory: " & objItem.LogFileDirectory
    Wscript.Echo "Log File Local Time Rollover: " & _
        objItem.LogFileLocaltimeRollover
    Wscript.Echo "Log File Period: " & objItem.LogFilePeriod
    Wscript.Echo "Log File Truncate Size: " & objItem.LogFileTruncateSize
    Wscript.Echo "Log Odbc Data Source: " & objItem.LogOdbcDataSource
    Wscript.Echo "Log Odbc Password: " & objItem.LogOdbcPassword
    Wscript.Echo "Log Odbc Table Name: " & objItem.LogOdbcTableName
    Wscript.Echo "Log Odbc User Name: " & objItem.LogOdbcUserName
    Wscript.Echo "Logon Method: " & objItem.LogonMethod
    Wscript.Echo "Log Plugin Clsid: " & objItem.LogPluginClsid
    Wscript.Echo "Log Type: " & objItem.LogType
    Wscript.Echo "Maximum Bandwidth: " & objItem.MaxBandwidth
    Wscript.Echo "Maximum Bandwidth Blocked: " & _
        objItem.MaxBandwidthBlocked
    Wscript.Echo "Maximum Connections: " & objItem.MaxConnections
    Wscript.Echo "Maximum Endpoint Connections: " & _
        objItem.MaxEndpointConnections
    Wscript.Echo "Maximum Request Entity Allowed: " & _
        objItem.MaxRequestEntityAllowed
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "NT Authentication Providers: " & _
        objItem.NTAuthenticationProviders
    Wscript.Echo "Passport Require AD Mapping: " & _
        objItem.PassportRequireADMapping
    Wscript.Echo "Password Cache TTL: " & objItem.PasswordCacheTTL
    Wscript.Echo "Password Change Flags: " & objItem.PasswordChangeFlags
    Wscript.Echo "Password Expire Prenotify Days: " & _
        objItem.PasswordExpirePrenotifyDays
    Wscript.Echo "Pool Idc Timeout: " & objItem.PoolIdcTimeout
    Wscript.Echo "Process NT CRIf Logged On: " & _
        objItem.ProcessNTCRIfLoggedOn
    Wscript.Echo "Realm: " & objItem.Realm
    For Each strHeader in objItem.RedirectHeaders
        Wscript.Echo "Redirect Headers: " & strHeader
    Next
    Wscript.Echo "Revocation Freshness Time: " & _
        objItem.RevocationFreshnessTime
    Wscript.Echo "Revocation URL Retrieval Timeout: " & _
        objItem.RevocationURLRetrievalTimeout
    Wscript.Echo "Server AutoStart: " & objItem.ServerAutoStart
    Wscript.Echo "Server Command: " & objItem.ServerCommand
    Wscript.Echo "Server Comment: " & objItem.ServerComment
    Wscript.Echo "Server ID: " & objItem.ServerID
    Wscript.Echo "Server Listen Backlog: " & objItem.ServerListenBacklog
    Wscript.Echo "Server Listen Timeout: " & objItem.ServerListenTimeout
    Wscript.Echo "Server Size: " & objItem.ServerSize
    Wscript.Echo "Setting ID: " & objItem.SettingID
    Wscript.Echo "Shutdown Time Limit: " & objItem.ShutdownTimeLimit
    Wscript.Echo "SSI Exec Disable: " & objItem.SSIExecDisable
    Wscript.Echo "SSL Always Negotiate Client Certificate: " & _
        objItem.SSLAlwaysNegoClientCert
    Wscript.Echo "Ssl Ctl Identifier: " & objItem.SslCtlIdentifier
    Wscript.Echo "Ssl Ctl Store Name: " & objItem.SslCtlStoreName
    Wscript.Echo "SSL Store Name: " & objItem.SSLStoreName
    Wscript.Echo "Upload Read Ahead Size: " & _
        objItem.UploadReadAheadSize
    Wscript.Echo "Use Digest SSP: " & objItem.UseDigestSSP
    Wscript.Echo "Win32 Error: " & objItem.Win32Error
Next

