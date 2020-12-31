' Description: Returns information about the POP3 service settings configured on a POP3 server.


strComputer = "LocalHost"
 
Set objIIS = GetObject("IIS://" & strComputer & "/POP3SVC")
 
Wscript.Echo "Access Execute: " & objIIS.AccessExecute
Wscript.Echo "Access Flags: " & objIIS.AccessFlags
Wscript.Echo "Access No Physical Directory: " & _
    objIIS.AccessNoPhysicalDir
Wscript.Echo "Access No Remote Execute: " & _
    objIIS.AccessNoRemoteExecute
Wscript.Echo "Access No Remote Read: " & objIIS.AccessNoRemoteRead
Wscript.Echo "Access No Remote Script: " & objIIS.AccessNoRemoteScript
Wscript.Echo "Access No Remote Write: " & objIIS.AccessNoRemoteWrite
Wscript.Echo "Access Read: " & objIIS.AccessRead
Wscript.Echo "Access Script: " & objIIS.AccessScript
Wscript.Echo "Access Source: " & objIIS.AccessSource
Wscript.Echo "Access SSL: " & objIIS.AccessSSL
Wscript.Echo "Access SSL 128: " & objIIS.AccessSSL128
Wscript.Echo "Access SSL Flags: " & objIIS.AccessSSLFlags
Wscript.Echo "Access SS LMapCert: " & objIIS.AccessSSLMapCert
Wscript.Echo "Access SSL Negotiate Certificate: " & _
    objIIS.AccessSSLNegotiateCert
Wscript.Echo "Access SSL Require Certificate: " & _
    objIIS.AccessSSLRequireCert
Wscript.Echo "Access Write: " & objIIS.AccessWrite
Wscript.Echo "Admin ACL Bin: " & objIIS.AdminACLBin
Wscript.Echo "Authentication Anonymous: " & objIIS.AuthAnonymous
Wscript.Echo "Authentication Basic: " & objIIS.AuthBasic
Wscript.Echo "Authentication Flags: " & objIIS.AuthFlags
Wscript.Echo "Authentication MD5: " & objIIS.AuthMD5
Wscript.Echo "Authentication NTLM: " & objIIS.AuthNTLM
Wscript.Echo "Authentication Passport: " & objIIS.AuthPassport
Wscript.Echo "Az Enable: " & objIIS.AzEnable
Wscript.Echo "Az Scope Name: " & objIIS.AzScopeName
Wscript.Echo "Az Store Name: " & objIIS.AzStoreName
Wscript.Echo "Caption: " & objIIS.Caption
Wscript.Echo "Connection Timeout: " & objIIS.ConnectionTimeout
Wscript.Echo "Default Logon Domain: " & objIIS.DefaultLogonDomain
Wscript.Echo "Description: " & objIIS.Description
Wscript.Echo "Don't Log: " & objIIS.DontLog
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
Wscript.Echo "Log Ext File Time Taken: " & _
    objIIS.LogExtFileTimeTaken
Wscript.Echo "Log Ext File URI Query: " & objIIS.LogExtFileUriQuery
Wscript.Echo "Log Ext File URI Stem: " & objIIS.LogExtFileUriStem
Wscript.Echo "Log Ext File User Agent: " & _
    objIIS.LogExtFileUserAgent
Wscript.Echo "Log Ext File User Name: " & objIIS.LogExtFileUserName
Wscript.Echo "Log Ext File Win32 Status: " & _
    objIIS.LogExtFileWin32Status
Wscript.Echo "Log File Directory: " & objIIS.LogFileDirectory
Wscript.Echo "Log File Period: " & objIIS.LogFilePeriod
Wscript.Echo "Log File Truncate Size: " & _
    objIIS.LogFileTruncateSize
Wscript.Echo "Log Module List: " & objIIS.LogModuleList
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
Wscript.Echo "Name: " & objIIS.Name
Wscript.Echo "NT Authentication Providers: " & _
    objIIS.NTAuthenticationProviders
Wscript.Echo "Pop3 Clear Text Provider: " & _
    objIIS.Pop3ClearTextProvider
Wscript.Echo "Pop3 Default Domain: " & objIIS.Pop3DefaultDomain
Wscript.Echo "Pop3 Expire Delay: " & objIIS.Pop3ExpireDelay
Wscript.Echo "Pop3 Expire Mail: " & objIIS.Pop3ExpireMail
Wscript.Echo "Pop3 Expire Start: " & objIIS.Pop3ExpireStart
Wscript.Echo "Pop3 Mail Expiration Time: " & _
    objIIS.Pop3MailExpirationTime
Wscript.Echo "Pop3 Routing Dll: " & objIIS.Pop3RoutingDll
Wscript.Echo "Server AutoStart: " & objIIS.ServerAutoStart
Wscript.Echo "Server Comment: " & objIIS.ServerComment
Wscript.Echo "Server Listen Backlog: " & _
    objIIS.ServerListenBacklog
Wscript.Echo "Server Listen Timeout: " & _
    objIIS.ServerListenTimeout
Wscript.Echo "Setting ID: " & objIIS.SettingID

