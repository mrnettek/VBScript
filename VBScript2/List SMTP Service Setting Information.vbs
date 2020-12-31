' Description: Returns global SMTP server metabase property values for a computer running IIS 6.0.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsSmtpServiceSetting")
 
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
    Wscript.Echo "Access SSL Negotiate Certificate: " & _
        objItem.AccessSSLNegotiateCert
    Wscript.Echo "Access SSL Require Certificate: " & _
        objItem.AccessSSLRequireCert
    Wscript.Echo "Access Write: " & objItem.AccessWrite
    Wscript.Echo "Add No Headers: " & objItem.AddNoHeaders
    Wscript.Echo "Admin ACL Bin: " & objItem.AdminACLBin
    Wscript.Echo "Allow Anonymous: " & objItem.AllowAnonymous
    Wscript.Echo "Always Use Ssl: " & objItem.AlwaysUseSsl
    Wscript.Echo "Anonymous Only: " & objItem.AnonymousOnly
    Wscript.Echo "Anonymous Password Sync: " & _
        objItem.AnonymousPasswordSync
    Wscript.Echo "Anonymous User Name: " & objItem.AnonymousUserName
    Wscript.Echo "Anonymous User Password: " & _
        objItem.AnonymousUserPass
    Wscript.Echo "Authentication Anonymous: " & objItem.AuthAnonymous
    Wscript.Echo "Authentication Basic: " & objItem.AuthBasic
    Wscript.Echo "Authentication Flags: " & objItem.AuthFlags
    Wscript.Echo "Authentication MD5: " & objItem.AuthMD5
    Wscript.Echo "Authentication NTLM: " & objItem.AuthNTLM
    Wscript.Echo "Authentication Passport: " & objItem.AuthPassport
    Wscript.Echo "Az Enable: " & objItem.AzEnable
    Wscript.Echo "Az Scope Name: " & objItem.AzScopeName
    Wscript.Echo "Az Store Name: " & objItem.AzStoreName
    Wscript.Echo "Bad Mail Directory: " & objItem.BadMailDirectory
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Connection Timeout: " & objItem.ConnectionTimeout
    Wscript.Echo "Connect Response: " & objItem.ConnectResponse
    Wscript.Echo "Default Domain: " & objItem.DefaultDomain
    Wscript.Echo "Default Logon Domain: " & _
        objItem.DefaultLogonDomain
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Disable Socket Pooling: " & _
        objItem.DisableSocketPooling
    Wscript.Echo "Do Masquerade: " & objItem.DoMasquerade
    Wscript.Echo "Don't Log: " & objItem.DontLog
    Wscript.Echo "Drop Directory: " & objItem.DropDirectory
    Wscript.Echo "Enable Reverse DNS Lookup: " & _
        objItem.EnableReverseDnsLookup
    Wscript.Echo "Etrn Days: " & objItem.EtrnDays
    Wscript.Echo "Etrn Subdomains: " & objItem.EtrnSubdomains
    Wscript.Echo "Fully Qualified Domain Name: " & _
        objItem.FullyQualifiedDomainName
    Wscript.Echo "Hop Count: " & objItem.HopCount
    Wscript.Echo "Limit Remote Connections: " & _
        objItem.LimitRemoteConnections
    Wscript.Echo "Local Retry Attempts: " & objItem.LocalRetryAttempts
    Wscript.Echo "Local Retry Interval: " & objItem.LocalRetryInterval
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
    Wscript.Echo "Log Ext File Server IP: " & objItem.LogExtFileServerIp
    Wscript.Echo "Log Ext File Server Port: " & _
        objItem.LogExtFileServerPort
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
    Wscript.Echo "Masquerade Domain: " & objItem.MasqueradeDomain
    Wscript.Echo "Maximum Bandwidth: " & objItem.MaxBandwidth
    Wscript.Echo "Maximum Batched Messages: " & objItem.MaxBatchedMessages
    Wscript.Echo "Maximum Connections: " & objItem.MaxConnections
    Wscript.Echo "Maximum Diectory rChange IO Size: " & _
        objItem.MaxDirChangeIOSize
    Wscript.Echo "Maximum Endpoint Connections: " & _
        objItem.MaxEndpointConnections
    Wscript.Echo "Maximum Mail Objects: " & objItem.MaxMailObjects
    Wscript.Echo "Maximum Message Size: " & objItem.MaxMessageSize
    Wscript.Echo "Maximum Out Connections: " & objItem.MaxOutConnections
    Wscript.Echo "Maximum Out Connections Per Domain: " & _
        objItem.MaxOutConnectionsPerDomain
    Wscript.Echo "Maximum Recipients: " & objItem.MaxRecipients
    Wscript.Echo "Maximum Session Size: " & objItem.MaxSessionSize
    Wscript.Echo "Maximum SMTP Errors: " & objItem.MaxSmtpErrors
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Name Resolution Type: " & objItem.NameResolutionType
    Wscript.Echo "NT Authentication Providers: " & _
        objItem.NTAuthenticationProviders
    Wscript.Echo "Pickup Directory: " & objItem.PickupDirectory
    Wscript.Echo "Queue Directory: " & objItem.QueueDirectory
    Wscript.Echo "Realm: " & objItem.Realm
    Wscript.Echo "Relay For Authentication : " & objItem.RelayForAuth
    For Each strIP in objItem.RelayIpList
        Wscript.Echo "Relay IP List: " & strIP
    Next
    Wscript.Echo "Remote Retry Attempts: " & _
        objItem.RemoteRetryAttempts
    Wscript.Echo "Remote Retry Interval: " & _
        objItem.RemoteRetryInterval
    Wscript.Echo "Remote SMTP Port: " & objItem.RemoteSmtpPort
    Wscript.Echo "Remote SMTP Secure Port: " & _
        objItem.RemoteSmtpSecurePort
    Wscript.Echo "Remote Timeout: " & objItem.RemoteTimeout
    Wscript.Echo "Route Action: " & objItem.RouteAction
    Wscript.Echo "Route Password: " & objItem.RoutePassword
    Wscript.Echo "Route User Name: " & objItem.RouteUserName
    Wscript.Echo "Routing Dll: " & objItem.RoutingDll
    Wscript.Echo "Sasl Logon Domain: " & objItem.SaslLogonDomain
    Wscript.Echo "Send Bad To: " & objItem.SendBadTo
    Wscript.Echo "Send Ndr To: " & objItem.SendNdrTo
    Wscript.Echo "Server AutoStart: " & objItem.ServerAutoStart
    Wscript.Echo "Server Comment: " & objItem.ServerComment
    Wscript.Echo "Server Listen Timeout: " & _
        objItem.ServerListenTimeout
    Wscript.Echo "Setting ID: " & objItem.SettingID
    Wscript.Echo "Should Deliver: " & objItem.ShouldDeliver
    Wscript.Echo "Should Pickup Mail: " & objItem.ShouldPickupMail
    Wscript.Echo "Should Pipeline In: " & objItem.ShouldPipelineIn
    Wscript.Echo "Should Pipeline Out: " & objItem.ShouldPipelineOut
    Wscript.Echo "Smart Host: " & objItem.SmartHost
    Wscript.Echo "Smart Host Type: " & objItem.SmartHostType
    Wscript.Echo "SMTP Adv Queue Dll: " & objItem.SmtpAdvQueueDll
    Wscript.Echo "SMTP Aqueue Wait: " & objItem.SmtpAqueueWait
    Wscript.Echo "SMTP Authetication Timeout: " & _
        objItem.SmtpAuthTimeout
    Wscript.Echo "SMTP Bdat Timeout: " & objItem.SmtpBdatTimeout
    Wscript.Echo "SMTP Clear Text Provider: " & _
        objItem.SmtpClearTextProvider
    Wscript.Echo "SMTP Command Log Mask: " & _
        objItem.SmtpCommandLogMask
    Wscript.Echo "SMTP Connect Timeout: " & _
        objItem.SmtpConnectTimeout
    Wscript.Echo "SMTP Data Timeout: " & objItem.SmtpDataTimeout
    Wscript.Echo "SMTP Disable Relay: " & objItem.SmtpDisableRelay
    Wscript.Echo "SMTP Domain Validation Flags: " & _
        objItem.SmtpDomainValidationFlags
    Wscript.Echo "SMTP Dot Stuff Pickup Directory Files: " & _
        objItem.SmtpDotStuffPickupDirFiles
    Wscript.Echo "SMTP DSN Language ID: " & _
        objItem.SmtpDSNLanguageID
    Wscript.Echo "SMTP DSN Options: " & objItem.SmtpDSNOptions
    Wscript.Echo "SMTP Eventlog Level: " & objItem.SmtpEventlogLevel
    Wscript.Echo "SMTP Flush Mail File: " & _
        objItem.SmtpFlushMailFile
    Wscript.Echo "SMTP Helo No Domain: " & objItem.SmtpHeloNoDomain
    Wscript.Echo "SMTP Helo Timeout: " & objItem.SmtpHeloTimeout
    Wscript.Echo "SMTP Inbound Command Support Options: " & _
        objItem.SmtpInboundCommandSupportOptions
    Wscript.Echo "SMTP IP Restriction Flag: " & _
        objItem.SmtpIpRestrictionFlag
    Wscript.Echo "SMTP Local Delay Expire Minutes: " & _
        objItem.SmtpLocalDelayExpireMinutes
    Wscript.Echo "SMTP Local NDR Expire Minutes: " & _
        objItem.SmtpLocalNDRExpireMinutes
    Wscript.Echo "SMTP Mail From Timeout: " & _
        objItem.SmtpMailFromTimeout
    Wscript.Echo "SMTP Mail No Helo: " & objItem.SmtpMailNoHelo
    Wscript.Echo "SMTP Maximum Remote Q Threads: " & _
        objItem.SmtpMaxRemoteQThreads
    Wscript.Echo "SMTP Outbound Command Support Options: " & _
        objItem.SmtpOutboundCommandSupportOptions
    Wscript.Echo "SMTP Rcpt To Timeout: " & objItem.SmtpRcptToTimeout
    Wscript.Echo "SMTP Remote Delay Expire Minutes: " & _
        objItem.SmtpRemoteDelayExpireMinutes
    Wscript.Echo "SMTP Remote NDR Expire Minutes: " & _
        objItem.SmtpRemoteNDRExpireMinutes
    Wscript.Echo "SMTP Remote Progressive Retry: " & _
        objItem.SmtpRemoteProgressiveRetry
    Wscript.Echo "SMTP Remote Retry Threshold: " & _
        objItem.SmtpRemoteRetryThreshold
    Wscript.Echo "SMTP Rset Timeout: " & objItem.SmtpRsetTimeout
    Wscript.Echo "SMTP Sasl Timeout: " & objItem.SmtpSaslTimeout
    Wscript.Echo "SMTP SSL Certificate Hostname Validation: " & _
        objItem.SmtpSSLCertHostnameValidation
    Wscript.Echo "SMTP SSL Require Trusted CA: " & _
        objItem.SmtpSSLRequireTrustedCA
    Wscript.Echo "SMTP Turn Timeout: " & objItem.SmtpTurnTimeout
    Wscript.Echo "SMTP Use Tcp Dns: " & objItem.SmtpUseTcpDns
    Wscript.Echo "Updated Default Domain: " & _
        objItem.UpdatedDefaultDomain
    Wscript.Echo "Updated FQDN: " & objItem.UpdatedFQDN
Next

