' Description: Returns information about an SMTP server named SMTPSVC/1.


strComputer = "LocalHost"
Set objIIS = GetObject("IIS://" & strComputer & "/SMTPSVC/1")
 
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
Wscript.Echo "Add No Headers: " & objIIS.AddNoHeaders
For Each strACL in objIIS.AdminACLBin
    Wscript.Echo "Admin ACL Bin: " & strACL
Next
Wscript.Echo "Allow Anonymous: " & objIIS.AllowAnonymous
Wscript.Echo "Always Use SSl: " & objIIS.AlwaysUseSsl
Wscript.Echo "Anonymous Only: " & objIIS.AnonymousOnly
Wscript.Echo "Anonymous Password Sync: " & _
    objIIS.AnonymousPasswordSync
Wscript.Echo "Anonymous User Name: " & objIIS.AnonymousUserName
Wscript.Echo "Authentication Anonymous: " & objIIS.AuthAnonymous
Wscript.Echo "Authentication Basic: " & objIIS.AuthBasic
Wscript.Echo "Authentication Flags: " & objIIS.AuthFlags
Wscript.Echo "Authentication MD5: " & objIIS.AuthMD5
Wscript.Echo "Authentication NTLM: " & objIIS.AuthNTLM
Wscript.Echo "Authentication Passport: " & objIIS.AuthPassport
Wscript.Echo "Az Enable: " & objIIS.AzEnable
Wscript.Echo "Az Scope Name: " & objIIS.AzScopeName
Wscript.Echo "Az Store Name: " & objIIS.AzStoreName
Wscript.Echo "Bad Mail Directory: " & objIIS.BadMailDirectory
Wscript.Echo "Cluster Enabled: " & objIIS.ClusterEnabled
Wscript.Echo "Connection Timeout: " & objIIS.ConnectionTimeout
Wscript.Echo "Connect Response: " & objIIS.ConnectResponse
Wscript.Echo "Default Domain: " & objIIS.DefaultDomain
Wscript.Echo "Default Logon Domain: " & _
    objIIS.DefaultLogonDomain
Wscript.Echo "Disable Socket Pooling: " & _
    objIIS.DisableSocketPooling
Wscript.Echo "Do Masquerade: " & objIIS.DoMasquerade
Wscript.Echo "Don't Log: " & objIIS.DontLog
Wscript.Echo "Drop Directory: " & objIIS.DropDirectory
Wscript.Echo "Enable Reverse DNS Lookup: " & _
    objIIS.EnableReverseDnsLookup
Wscript.Echo "Etrn Days: " & objIIS.EtrnDays
Wscript.Echo "Etrn Subdomains: " & objIIS.EtrnSubdomains
Wscript.Echo "Fully Qualified Domain Name: " & _
    objIIS.FullyQualifiedDomainName
Wscript.Echo "Hop Count: " & objIIS.HopCount
Wscript.Echo "Limit Remote Connections: " & _
    objIIS.LimitRemoteConnections
Wscript.Echo "Local Retry Attempts: " & _
    objIIS.LocalRetryAttempts
Wscript.Echo "Local Retry Interval: " & _
    objIIS.LocalRetryInterval
Wscript.Echo "Log Ext File Bytes Received: " & _
    objIIS.LogExtFileBytesRecv
Wscript.Echo "Log Ext File Bytes Sent: " & _
    objIIS.LogExtFileBytesSent
Wscript.Echo "Log Ext File Client IP: " & _
    objIIS.LogExtFileClientIp
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
Wscript.Echo "Log Odbc Data Source: " & objIIS.LogOdbcDataSource
Wscript.Echo "Log Odbc Password: " & objIIS.LogOdbcPassword
Wscript.Echo "Log Odbc Table Name: " & objIIS.LogOdbcTableName
Wscript.Echo "Log Odbc User Name: " & objIIS.LogOdbcUserName
Wscript.Echo "Log Plugin Clsid: " & objIIS.LogPluginClsid
Wscript.Echo "LogT ype: " & objIIS.LogType
Wscript.Echo "Masquerade Domain: " & objIIS.MasqueradeDomain
Wscript.Echo "Maximum Bandwidth: " & objIIS.MaxBandwidth
Wscript.Echo "Maximum Batched Messages: " & _
    objIIS.MaxBatchedMessages
Wscript.Echo "Maximum Connections: " & objIIS.MaxConnections
Wscript.Echo "Maximum Directory Change IO Size: " & _
    objIIS.MaxDirChangeIOSize
Wscript.Echo "Maximum Endpoint Connections: " & _
    objIIS.MaxEndpointConnections
Wscript.Echo "Maximum Mail Objects: " & objIIS.MaxMailObjects
Wscript.Echo "Maximum Message Size: " & objIIS.MaxMessageSize
Wscript.Echo "Maximum Out Connections: " & _
    objIIS.MaxOutConnections
Wscript.Echo "Maximum Out Connections Per Domain: " & _
    objIIS.MaxOutConnectionsPerDomain
Wscript.Echo "Maximum Recipients: " & objIIS.MaxRecipients
Wscript.Echo "Maximum Session Size: " & objIIS.MaxSessionSize
Wscript.Echo "Maximum SMTP Errors: " & objIIS.MaxSmtpErrors
Wscript.Echo "Name: " & objIIS.Name
Wscript.Echo "Name Resolution Type: " & _
    objIIS.NameResolutionType
Wscript.Echo "NT Authentication Providers: " & _
    objIIS.NTAuthenticationProviders
Wscript.Echo "Pickup Directory: " & objIIS.PickupDirectory
Wscript.Echo "Queue Directory: " & objIIS.QueueDirectory
Wscript.Echo "Realm: " & objIIS.Realm
Wscript.Echo "Relay For Authentication: " & objIIS.RelayForAuth
Wscript.Echo "Remote Retry Attempts: " & _
    objIIS.RemoteRetryAttempts
Wscript.Echo "Remote Retry Interval: " & _
    objIIS.RemoteRetryInterval
Wscript.Echo "Remote SMTP Port: " & objIIS.RemoteSmtpPort
Wscript.Echo "Remote SMTP Secure Port: " & _
    objIIS.RemoteSmtpSecurePort
Wscript.Echo "Remote Timeout: " & objIIS.RemoteTimeout
Wscript.Echo "Route Action: " & objIIS.RouteAction
Wscript.Echo "Route Password: " & objIIS.RoutePassword
Wscript.Echo "Route User Name: " & objIIS.RouteUserName
Wscript.Echo "Routing Dll: " & objIIS.RoutingDll
Wscript.Echo "Sasl Logon Domain: " & objIIS.SaslLogonDomain
Wscript.Echo "Send Bad To: " & objIIS.SendBadTo
Wscript.Echo "Send Ndr To: " & objIIS.SendNdrTo
Wscript.Echo "Server AutoStart: " & objIIS.ServerAutoStart
Wscript.Echo "Server Comment: " & objIIS.ServerComment
Wscript.Echo "Server Listen Backlog: " & _
    objIIS.ServerListenBacklog
Wscript.Echo "Server Listen Timeout: " & _
    objIIS.ServerListenTimeout
'Wscript.Echo "Setting ID: " & objIIS.SettingID
Wscript.Echo "Should Deliver: " & objIIS.ShouldDeliver
Wscript.Echo "Should Pickup Mail: " & objIIS.ShouldPickupMail
Wscript.Echo "Should Pipeline In: " & objIIS.ShouldPipelineIn
Wscript.Echo "Should Pipeline Out: " & _
    objIIS.ShouldPipelineOut
Wscript.Echo "Smart Host: " & objIIS.SmartHost
Wscript.Echo "Smart Host Type: " & objIIS.SmartHostType
Wscript.Echo "SMTP Aqueue Wait: " & objIIS.SmtpAqueueWait
Wscript.Echo "SMTP Authentication Timeout: " & _
    objIIS.SmtpAuthTimeout
Wscript.Echo "SMTP Bdat Timeout: " & objIIS.SmtpBdatTimeout
Wscript.Echo "SMTP Clear Text Provider: " & _
    objIIS.SmtpClearTextProvider
Wscript.Echo "SMTP Connect Timeout: " & objIIS.SmtpConnectTimeout
Wscript.Echo "SMTP Data Timeout: " & objIIS.SmtpDataTimeout
Wscript.Echo "SMTP Disable Relay: " & objIIS.SmtpDisableRelay
Wscript.Echo "SMTP Domain Validation Flags: " & _
    objIIS.SmtpDomainValidationFlags
Wscript.Echo "SMTP Dot Stuff Pickup Directory Files: " & _
    objIIS.SmtpDotStuffPickupDirFiles
Wscript.Echo "SMTP DSN Language ID: " & objIIS.SmtpDSNLanguageID
Wscript.Echo "SMTP DSN Options: " & objIIS.SmtpDSNOptions
Wscript.Echo "SMTP Eventlog Level: " & objIIS.SmtpEventlogLevel
Wscript.Echo "SMTP Helo No Domain: " & objIIS.SmtpHeloNoDomain
Wscript.Echo "SMTP Helo Timeout: " & objIIS.SmtpHeloTimeout
Wscript.Echo "SMTP Inbound Command Support Options: " & o_
    bjIIS.SmtpInboundCommandSupportOptions
Wscript.Echo "SMTP IP Restriction Flag: " & _
    objIIS.SmtpIpRestrictionFlag
Wscript.Echo "SMTP Local Delay Expire Minutes: " & _
    objIIS.SmtpLocalDelayExpireMinutes
Wscript.Echo "SMTP Local NDR Expire Minutes: " & _
    objIIS.SmtpLocalNDRExpireMinutes
Wscript.Echo "SMTP Mail From Timeout: " & _
    objIIS.SmtpMailFromTimeout
Wscript.Echo "SMTP Mail No Helo: " & objIIS.SmtpMailNoHelo
Wscript.Echo "SMTP Maximum Remote Q Threads: " & _
    objIIS.SmtpMaxRemoteQThreads
Wscript.Echo "SMTP Outbound Command Support Options: " & _
    objIIS.SmtpOutboundCommandSupportOptions
Wscript.Echo "SMTP Rcpt To Timeout: " & objIIS.SmtpRcptToTimeout
Wscript.Echo "SMTP Remote Delay Expire Minutes: " & _
    objIIS.SmtpRemoteDelayExpireMinutes
Wscript.Echo "SMTP Remote NDR Expire Minutes: " & _
    objIIS.SmtpRemoteNDRExpireMinutes
Wscript.Echo "SMTP Remote Progressive Retry: " & _
    objIIS.SmtpRemoteProgressiveRetry
Wscript.Echo "SMTP Remote Retry Threshold: " & _
    objIIS.SmtpRemoteRetryThreshold
Wscript.Echo "SMTP Rset Timeout: " & objIIS.SmtpRsetTimeout
Wscript.Echo "SMTP Sasl Timeout: " & objIIS.SmtpSaslTimeout
Wscript.Echo "SMTP SSL Certifcate Hostname Validation: " & _
    objIIS.SmtpSSLCertHostnameValidation
Wscript.Echo "SMTP SSL Require Trusted CA: " & _
    objIIS.SmtpSSLRequireTrustedCA
Wscript.Echo "SMTP Turn Timeout: " & objIIS.SmtpTurnTimeout
Wscript.Echo "SMTP Use Tcp Dns: " & objIIS.SmtpUseTcpDns
Wscript.Echo "Updated Default Domain: " & _
    objIIS.UpdatedDefaultDomain
Wscript.Echo "Updated FQDN: " & objIIS.UpdatedFQDN
Wscript.Echo "Win32 Error: " & objIIS.Win32Error

