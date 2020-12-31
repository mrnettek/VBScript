' Description: Enumerates the properties of an IMAP server named IMAP4SVC/1.


strComputer = "LocalHost"
 
Set objIIS = GetObject("IIS://" & strComputer & "/IMAP4SVC/1")
 
Wscript.Echo "Access Execute: " & objIIS.AccessExecute
Wscript.Echo "Access Flags: " & objIIS.AccessFlags
Wscript.Echo "Access No Physical Directory: " & _
    objIIS.AccessNoPhysicalDir
Wscript.Echo "Access No Remote Execute: " & _
    objIIS.AccessNoRemoteExecute
Wscript.Echo "Access No Remote Read: " & _
    objIIS.AccessNoRemoteRead
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
Wscript.Echo "Authentication Anonymous: " & _
    objIIS.AuthAnonymous
Wscript.Echo "Authentication Basic: " & objIIS.AuthBasic
Wscript.Echo "Authentication Flags: " & objIIS.AuthFlags
Wscript.Echo "Authentication MD5: " & objIIS.AuthMD5
Wscript.Echo "Authentication NTLM: " & objIIS.AuthNTLM
Wscript.Echo "Authentication Passport: " & _
    objIIS.AuthPassport
Wscript.Echo "Az Enable: " & objIIS.AzEnable
Wscript.Echo "Az Scope Name: " & objIIS.AzScopeName
Wscript.Echo "Az Store Name: " & objIIS.AzStoreName
Wscript.Echo "Connection Timeout: " & _
    objIIS.ConnectionTimeout
Wscript.Echo "Default Logon Domain: " & _
    objIIS.DefaultLogonDomain
Wscript.Echo "Don't Log: " & objIIS.DontLog
Wscript.Echo "IMAP Clear Text Provider: " & _
    objIIS.ImapClearTextProvider
Wscript.Echo "IMAP Default Domain: " & objIIS.ImapDefaultDomain
Wscript.Echo "IMAP Expire Delay: " & objIIS.ImapExpireDelay
Wscript.Echo "IMAP Expire Mail: " & objIIS.ImapExpireMail
Wscript.Echo "IMAP Expire Start: " & objIIS.ImapExpireStart
Wscript.Echo "IMAP Mail Expiration Time: " & _
    objIIS.ImapMailExpirationTime
Wscript.Echo "IMAP Routing Dll: " & objIIS.ImapRoutingDll
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
Wscript.Echo "Log Ext File Win32 Status: " & _
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
Wscript.Echo "Name: " & objIIS.Name
Wscript.Echo "NT Authentication Providers: " & _
    objIIS.NTAuthenticationProviders
For Each strBinding in objIIS.SecureBindings
    Wscript.Echo "Secure Bindings: " & strBinding
Next
Wscript.Echo "Server AutoStart: " & objIIS.ServerAutoStart
For Each strBinding in objIIS.ServerBindings
    Wscript.Echo "Server Bindings: " & strBinding
Next
Wscript.Echo "Server Comment: " & objIIS.ServerComment
Wscript.Echo "Server Listen Backlog: " & _
    objIIS.ServerListenBacklog
Wscript.Echo "Server Listen Timeout: " & _
    objIIS.ServerListenTimeout
Wscript.Echo "Win32 Error: " & objIIS.Win32Error

