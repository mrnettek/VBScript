' Description: Returns property values for the Web service on an IIS server.


strComputer = "LocalHost"
 
Set objIIS = GetObject("IIS://" & strComputer & "/W3SVC")
 
Wscript.Echo "Access Flags: " & objIIS.AccessFlags
Wscript.Echo "ASP Enable Typelib Cache: " & _
    objIIS.AspEnableTypelibCache
Wscript.Echo "Access SSL Flags: " & objIIS.AccessSSLFlags
Wscript.Echo "ASP Errors to NT Log: " & objIIS.AspErrorsToNTLog
Wscript.Echo "ASP Exception Catch Enabled: " & _
    objIIS.AspExceptionCatchEnable
Wscript.Echo "Allow Path Information for Script Mappings: " & _
    objIIS.AllowPathInfoForScriptMappings
Wscript.Echo "ASP Log Error Requests: " & objIIS.AspLogErrorRequests
Wscript.Echo "Anonymous Password Synch: " & objIIS.AnonymousPasswordSync
Wscript.Echo "ASP Processor Thread Maximum: " & _
    objIIS.AspProcessorThreadMax
Wscript.Echo "Anonynmous User Name: " & objIIS.AnonymousUserName
Wscript.Echo "ASP Queue Connection Test Time: " & _
    objIIS.AspQueueConnectionTestTime
Wscript.Echo "Anonymous User Password: " & objIIS.AnonymousUserPass
Wscript.Echo "ASP Queue Timeout: " & objIIS.AspQueueTimeout
Wscript.Echo "Application Allow Client Debugging: " & _
    objIIS.AppAllowClientDebug
Wscript.Echo "ASP Request Queue Maximum: " & objIIS.AspRequestQueueMax
Wscript.Echo "Application Allow Debugging: " & objIIS.AppAllowDebugging
Wscript.Echo "ASP Script Engine Cache Maximum: " & _
    objIIS.AspScriptEngineCacheMax
Wscript.Echo "Application Friendly Name: " & objIIS.AppFriendlyName
Wscript.Echo "ASP Script Error Message: " & objIIS.AspScriptErrorMessage
Wscript.Echo "Application Isolated: " & objIIS.AppIsolated
Wscript.Echo "ASP Script Error Sent to Browser: " & _
    objIIS.AspScriptErrorSentToBrowser
Wscript.Echo "Application Package ID: " & objIIS.AppPackageID
Wscript.Echo "ASP Script File Cache Size: " & objIIS.AspScriptFileCacheSize
Wscript.Echo "Application Package Name: " & objIIS.AppPackageName
Wscript.Echo "ASP Script Language: " & objIIS.AspScriptLanguage
Wscript.Echo "Application Root: " & objIIS.AppRoot
Wscript.Echo "ASP Session Maximum: " & objIIS.AspSessionMax
Wscript.Echo "Application WAM Clsid: " & objIIS.AppWamClsid
Wscript.Echo "ASP Script Timeout: " & objIIS.AspScriptTimeout
Wscript.Echo "ASP Allow Out-of-Process Components: " & _
    objIIS.AspAllowOutOfProcComponents
Wscript.Echo "ASP Session Timeout: " & objIIS.AspSessionTimeout
Wscript.Echo "ASP Allow Session State: " & objIIS.AspAllowSessionState
Wscript.Echo "ASP Buffering On: " & objIIS.AspBufferingOn
Wscript.Echo "ASP Codepage: " & objIIS.AspCodepage
Wscript.Echo "ASP Enable Application Restart: " & _
    objIIS.AspEnableApplicationRestart
Wscript.Echo "ASP Enable ASP HTML Fallback: " & _
    objIIS.AspEnableAspHtmlFallback
Wscript.Echo "ASP Enabled Chunked Encoding: " & _
    objIIS.AspEnableChunkedEncoding
Wscript.Echo "ASP Enable parent Paths: " & objIIS.AspEnableParentPaths
Wscript.Echo "ASP Track Threading Model: " & objIIS.AspTrackThreadingModel    
Wscript.Echo "Authentication Flags: " & objIIS.AuthFlags
Wscript.Echo "Default Document: " & objIIS.DefaultDoc
Wscript.Echo "Authentication Persistence: " & objIIS.AuthPersistence
Wscript.Echo "Default Document Footer: " & objIIS.DefaultDocFooter
Wscript.Echo "Cache Control Custom: " & objIIS.CacheControlCustom
Wscript.Echo "Default Logon Domain: " & objIIS.DefaultLogonDomain
Wscript.Echo "Cache Control Maximum Age: " & objIIS.CacheControlMaxAge
Wscript.Echo "Directory Browse Flags: " & objIIS.DirBrowseFlags
Wscript.Echo "Cache Control No Cache: " & objIIS.CacheControlNoCache
Wscript.Echo "Directory Levels to Scan: " & objIIS.DirectoryLevelsToScan
Wscript.Echo "Cache ISAPI: " & objIIS.CacheISAPI
Wscript.Echo "Disable Socket Pooling: " & objIIS.DisableSocketPooling
Wscript.Echo "Content Indexed: " & objIIS.ContentIndexed
Wscript.Echo "Don't Log: " & objIIS.DontLog
Wscript.Echo "Connection Timeout: " & objIIS.ConnectionTimeout
Wscript.Echo "Enable Document Footer: " & objIIS.EnableDocFooter
Wscript.Echo "Enable Reverse DNS: " & objIIS.EnableReverseDns
For Each strError in objIIS.HttpErrors
    Wscript.Echo "HTTP Error: " & strError
Next
Wscript.Echo "HTTP Expires: " & objIIS.HttpExpires
For Each strPic in objIIS.HttpPics
    Wscript.Echo "HTTP Pic: " & strPic.Name
Next
Wscript.Echo "Create CGI with New Console: " & _
    objIIS.CreateCGIWithNewConsole
Wscript.Echo "Create Process as User: " & objIIS.CreateProcessAsUser
For Each strApp in objIIS.InProcessIsapiApps
    Wscript.Echo "In-Process ISAPI Application: " & strApp
Next
Wscript.Echo "Log Ext File Flags: " & objIIS.LogExtFileFlags
Wscript.Echo "Log ODBC Password: " & objIIS.LogOdbcPassword
Wscript.Echo "Log File Directory: " & objIIS.LogFileDirectory
Wscript.Echo "Log ODBC Table Name: " & objIIS.LogOdbcTableName
Wscript.Echo "Log File Local Time Rollover: " & _
    objIIS.LogFileLocaltimeRollover
Wscript.Echo "Log ODBC User Name: " & objIIS.LogOdbcUserName
Wscript.Echo "Log File Period: " & objIIS.LogFilePeriod
Wscript.Echo "Logon Method: " & objIIS.LogonMethod
Wscript.Echo "Log File Truncate Size: " & objIIS.LogFileTruncateSize
Wscript.Echo "Log Plugin Clsid: " & objIIS.LogPluginClsid
Wscript.Echo "Log ODBC Data Source: " & objIIS.LogOdbcDataSource
Wscript.Echo "Log Type: " & objIIS.LogType     
For Each strMap in objIIS.ScriptMaps
    Wscript.Echo "Script Map: " & strMap
Next
Wscript.Echo "Maximum Connections: " & objIIS.MaxConnections
Wscript.Echo "Server AutoStart: " & objIIS.ServerAutoStart
Wscript.Echo "Maximum Endpoint Connections: " & _
    objIIS.MaxEndpointConnections
For Each strBinding in objIIS.ServerBindings
    Wscript.Echo "Server Binding: " & strBinding.Name
Next
Wscript.Echo "Server Comments: " & objIIS.ServerComment
Wscript.Echo "Server Listen Backlog: " & objIIS.ServerListenBacklog
Wscript.Echo "NT Authentication Providers: " & _
    objIIS.NTAuthenticationProviders
Wscript.Echo "Server Listen Timeout: " & objIIS.ServerListenTimeout
Wscript.Echo "Password Cache Time-to-Live: " & _
    objIIS.PasswordCacheTTL
Wscript.Echo "Server Size: " & objIIS.ServerSize
Wscript.Echo "Password Change Flags: " & objIIS.PasswordChangeFlags
Wscript.Echo "SSI Exec Disable: " & objIIS.SSIExecDisable
Wscript.Echo "Password Expire Pre-Notify Days: " & _
    objIIS.PasswordExpirePrenotifyDays
Wscript.Echo "SSL Use DS Mapper: " & objIIS.SslUseDsMapper
Wscript.Echo "Pool IDC Timeout: " & objIIS.PoolIdcTimeout
Wscript.Echo "Process NTCR if Logged On: " & _
    objIIS.ProcessNTCRIfLoggedOn
Wscript.Echo "Upload Read-Ahead Size: " & _
    objIIS.UploadReadAheadSize
Wscript.Echo "Realm: " & objIIS.Realm
Wscript.Echo "WAM User Name: " & objIIS.WAMUserName
For Each strHeader in objIIS.RedirectHeaders
    Wscript.Echo "Redirect Header: " & strHeader.Name
Next
Wscript.Echo "WAM User Password: " & objIIS.WAMUserPass

