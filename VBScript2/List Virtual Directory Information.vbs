' Description: Returns global Web virtual directory metabase property values for a computer running IIS 6.0.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsWebVirtualDirSetting")
 
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
    Wscript.Echo "Anonymous Password Sync: " & _
        objItem.AnonymousPasswordSync
    Wscript.Echo "Anonymous User Name: " & objItem.AnonymousUserName
    Wscript.Echo "Anonymous User Password: " & objItem.AnonymousUserPass
    Wscript.Echo "Application Allow Client Debug: " & _
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
    Wscript.Echo "ASP Enable Chunked Encoding: " & _
        objItem.AspEnableChunkedEncoding
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
    Wscript.Echo "AspRunOnEndAnonymously: " & _
        objItem.AspRunOnEndAnonymously
    Wscript.Echo "ASP Script Engine Cache Max: " & _
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
    Wscript.Echo "Authentication Anonymous: " & objItem.AuthAnonymous
    Wscript.Echo "Authentication Basic: " & objItem.AuthBasic
    Wscript.Echo "Authentication Flags: " & objItem.AuthFlags
    Wscript.Echo "Authentication MD5: " & objItem.AuthMD5
    Wscript.Echo "Authentication NTLM: " & objItem.AuthNTLM
    Wscript.Echo "Authentication Passport: " & objItem.AuthPassport
    Wscript.Echo "Authentication Persistence: " & _
        objItem.AuthPersistence
    Wscript.Echo "Authentication Persist Single Request: " & _
        objItem.AuthPersistSingleRequest
    Wscript.Echo "Az Enable: " & objItem.AzEnable
    Wscript.Echo "Az Impersonation Level: " & objItem.AzImpersonationLevel
    Wscript.Echo "Az Scope Name: " & objItem.AzScopeName
    Wscript.Echo "Az Store Name: " & objItem.AzStoreName
    Wscript.Echo "BITS Allow Overwrites: " & objItem.BITSAllowOverwrites
    Wscript.Echo "BITS Cleanup Work Item Key: " & _
        objItem.BITSCleanupWorkItemKey
    Wscript.Echo "BITS Host ID: " & objItem.BITSHostId
    Wscript.Echo "BITS Host ID Fallback Timeout: " & _
        objItem.BITSHostIdFallbackTimeout
    Wscript.Echo "BITS Maximum Upload Size: " & _
        objItem.BITSMaximumUploadSize
    Wscript.Echo "BITS Server Notification Type: " & _
        objItem.BITSServerNotificationType
    Wscript.Echo "BITS Server Notification URL: " & _
        objItem.BITSServerNotificationURL
    Wscript.Echo "BITS Session Directory: " & _
        objItem.BITSSessionDirectory
    Wscript.Echo "BITS Session Timeout: " & objItem.BITSSessionTimeout
    Wscript.Echo "BITS Upload Enabled: " & objItem.BITSUploadEnabled
    Wscript.Echo "BITS Upload Metadata Version: " & _
        objItem.BITSUploadMetadataVersion
    Wscript.Echo "Cache Control Custom: " & objItem.CacheControlCustom
    Wscript.Echo "Cache Control Maximum Age: " & _
        objItem.CacheControlMaxAge
    Wscript.Echo "Cache Control No Cache: " & objItem.CacheControlNoCache
    Wscript.Echo "Cache ISAPI: " & objItem.CacheISAPI
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "CGI Timeout: " & objItem.CGITimeout
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
    Wscript.Echo "Disable Static File Cache: " & _
        objItem.DisableStaticFileCache
    Wscript.Echo "Do Dynamic Compression: " & objItem.DoDynamicCompression
    Wscript.Echo "Don't Log: " & objItem.DontLog
    Wscript.Echo "Do Static Compression: " & objItem.DoStaticCompression
    Wscript.Echo "Enable Default Doc: " & objItem.EnableDefaultDoc
    Wscript.Echo "Enable Directory Browsing: " & objItem.EnableDirBrowsing
    Wscript.Echo "Enable Doc Footer: " & objItem.EnableDocFooter
    Wscript.Echo "Enable Reverse Dns: " & objItem.EnableReverseDns
    Wscript.Echo "FrontPage Web: " & objItem.FrontPageWeb
    Wscript.Echo "Http Expires: " & objItem.HttpExpires
    Wscript.Echo "Http Redirect: " & objItem.HttpRedirect
    Wscript.Echo "Logon Method: " & objItem.LogonMethod
    Wscript.Echo "Maximum Request Entity Allowed: " & _
        objItem.MaxRequestEntityAllowed
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "NT Authentication Providers: " & _
        objItem.NTAuthenticationProviders
    Wscript.Echo "Passport Require AD Mapping: " & _
        objItem.PassportRequireADMapping
    Wscript.Echo "Path: " & objItem.Path
    Wscript.Echo "Pool Idc Timeout: " & objItem.PoolIdcTimeout
    Wscript.Echo "Realm: " & objItem.Realm
    For Each strHeader in objItem.RedirectHeaders
        Wscript.Echo "Redirect Headers: " &  strHeader
    Next
    Wscript.Echo "Setting ID: " & objItem.SettingID
    Wscript.Echo "Shutdown Time Limit: " & objItem.ShutdownTimeLimit
    Wscript.Echo "SSI Exec Disable: " & objItem.SSIExecDisable
    Wscript.Echo "UNC Password: " & objItem.UNCPassword
    Wscript.Echo "UNC User Name: " & objItem.UNCUserName
    Wscript.Echo "Upload Read Ahead Size: " & objItem.UploadReadAheadSize
    Wscript.Echo "Use Digest SSP: " & objItem.UseDigestSSP
    Wscript.Echo "Win32 Error: " & objItem.Win32Error
Next

