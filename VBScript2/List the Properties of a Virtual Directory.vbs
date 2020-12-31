' Description: Demonstration script that returns the properties of the W3SVC/1/ROOT/Printers virtual directory on an IIS server.


strComputer = "LocalHost"
Set objIIS = GetObject("IIS://" & strComputer & "/W3SVC/1/ROOT/Printers")
 
Wscript.Echo "Access Flags: " & objIIS.AccessFlags
Wscript.Echo "ASP Errors to NT Log: " & objIIS.AspErrorsToNTLog
Wscript.Echo "Access SSL Flags: " & objIIS.AccessSSLFlags
Wscript.Echo "ASP Exception Catch Enabled: " & _
    objIIS.AspExceptionCatchEnable
Wscript.Echo "Anonymous Password Sync: " & objIIS.AnonymousPasswordSync
Wscript.Echo "ASP Log Error Requests: " & objIIS.AspLogErrorRequests
Wscript.Echo "Anonymous User name: " & objIIS.AnonymousUserName
Wscript.Echo "ASP Processor Thread MAximum: " & _
    objIIS.AspProcessorThreadMax
Wscript.Echo "Anonymous User Password: " & objIIS.AnonymousUserPass
Wscript.Echo "ASP Queue Connection Test Time: " & _
    objIIS.AspQueueConnectionTestTime
Wscript.Echo "Application Allow Client Debugging: " & _
    objIIS.AppAllowClientDebug
Wscript.Echo "ASP Queue Timeout: " & objIIS.AspQueueTimeout
Wscript.Echo "Application Allow Debugging: " & objIIS.AppAllowDebugging
Wscript.Echo "ASP Request Queue Maximum: " & objIIS.AspRequestQueueMax
Wscript.Echo "Application Friendly Name: " & objIIS.AppFriendlyName
Wscript.Echo "ASP Script Engine Cache Maximum: " & _
    objIIS.AspScriptEngineCacheMax
Wscript.Echo "Application Isolated: " & objIIS.AppIsolated
Wscript.Echo "ASP Script Error Message: " & objIIS.AspScriptErrorMessage
Wscript.Echo "Application OOP Recover Limit: " & _
    objIIS.AppOopRecoverLimit
Wscript.Echo "ASP Script Error Sent to Browser: " & _
    objIIS.AspScriptErrorSentToBrowser
Wscript.Echo "Application Package ID: " & objIIS.AppPackageID
Wscript.Echo "ASP Script File Cache Size: " & objIIS.AspScriptFileCacheSize
Wscript.Echo "Application Package Name: " & objIIS.AppPackageName
Wscript.Echo "ASP Script Language: " & objIIS.AspScriptLanguage
Wscript.Echo "Application Root: " & objIIS.AppRoot
Wscript.Echo "ASP Script Timeout: " & objIIS.AspScriptTimeout
Wscript.Echo "Application WAM Clsid: " & objIIS.AppWamClsid
Wscript.Echo "ASP Session Maximum: " & objIIS.AspSessionMax
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
Wscript.Echo "ASP Enable Chunked Encoding: " & objIIS.AspEnableChunkedEncoding
Wscript.Echo "ASP Enabled Parent Paths: " & objIIS.AspEnableParentPaths
Wscript.Echo "ASP Track Threading Model: " & objIIS.AspTrackThreadingModel
Wscript.Echo "ASP Enable Typelib Cache: " & objIIS.AspEnableTypelibCache
Wscript.Echo "Authentication Flags: " & objIIS.AuthFlags
For Each strError in objIIS.HttpErrors
    Wscript.Echo "HTTP Error: " & strError
Next
Wscript.Echo "Authentication Persistence: " & objIIS.AuthPersistence
Wscript.Echo "HTTP Expires: " & objIIS.HttpExpires
Wscript.Echo "Cache Control Custom: " & objIIS.CacheControlCustom
For Each strPics in objIIS.HttpPics
    Wscript.Echo "HTTP Pic: " & strPics
Next
Wscript.Echo "Cache Control Maximum Age: " & objIIS.CacheControlMaxAge
Wscript.Echo "HTTP Redirect: " & objIIS.HttpRedirect
Wscript.Echo "Cache Control No Cache: " & objIIS.CacheControlNoCache
Wscript.Echo "Cache ISAPI: " & objIIS.CacheISAPI
Wscript.Echo "Logon Method: " & objIIS.LogonMethod
Wscript.Echo "Content Indexed: " & objIIS.ContentIndexed
For Each strMap in objIIS.MimeMap
    Wscript.Echo "MIME Map: " & strMao
Next
Wscript.Echo "Create CGI with New Console: " & _
    objIIS.CreateCGIWithNewConsole
Wscript.Echo "Path: " & objIIS.Path
Wscript.Echo "Create Process as User: " & objIIS.CreateProcessAsUser
Wscript.Echo "Pool IDC Timeout: " & objIIS.PoolIdcTimeout
Wscript.Echo "Default Document: " & objIIS.DefaultDoc
Wscript.Echo "Default Document Footer: " & objIIS.DefaultDocFooter
Wscript.Echo "Realm: " & objIIS.Realm
Wscript.Echo "Default Logon Domain: " & objIIS.DefaultLogonDomain
For Each strHeader in objIIS.RedirectHeaders
    Wscript.Echo "Redirect Header: " & strHeader
Next
Wscript.Echo "Directory Browse Flags: " & objIIS.DirBrowseFlags
For Each strScriptMap in objIIS.ScriptMaps
    Wscript.Echo "Script Map: " & strScriptMap
Next
Wscript.Echo "Don't Log: " & objIIS.DontLog
Wscript.Echo "SSI Exec Disable: " & objIIS.SSIExecDisable
Wscript.Echo "Enable Document Footer: " & objIIS.EnableDocFooter
Wscript.Echo "Enable Reverse DNS: " & objIIS.EnableReverseDns
Wscript.Echo "UNC Password: " & objIIS.UNCPassword
For Each strHeader in objIIS.HttpCustomHeaders
    Wscript.Echo "HTTP Custom Header: " & strHeader
Next
Wscript.Echo "UNC User Name: " & objIIS.UNCUserName
Wscript.Echo "Upload Read-Ahead Size: " & objIIS.UploadReadAheadSize

