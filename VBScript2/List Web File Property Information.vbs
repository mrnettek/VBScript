' Description: Returns global Web file metabase property values for a computer running IIS 6.0.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery("Select * from IIsWebFileSetting")
 
For Each objItem in colItems
    Wscript.Echo "Access Execute: " & objItem.AccessExecute
    Wscript.Echo "Access Flags: " & objItem.AccessFlags
    Wscript.Echo "Access No Physical Directory : " & _
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
    Wscript.Echo "Access SSL Map Certificate: " & objItem.AccessSSLMapCert
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
    Wscript.Echo "Authentication Anonymous: " & objItem.AuthAnonymous
    Wscript.Echo "Authentication Basic: " & objItem.AuthBasic
    Wscript.Echo "Authentication Flags: " & objItem.AuthFlags
    Wscript.Echo "Authentication MD5: " & objItem.AuthMD5
    Wscript.Echo "Authentication NTLM: " & objItem.AuthNTLM
    Wscript.Echo "AuthenticationPassport: " & objItem.AuthPassport
    Wscript.Echo "Authentication Persistence: " & objItem.AuthPersistence
    Wscript.Echo "Authentication Persist Single Request: " & _
        objItem.AuthPersistSingleRequest
    Wscript.Echo "Az Enable: " & objItem.AzEnable
    Wscript.Echo "Az Impersonation Level: " & objItem.AzImpersonationLevel
    Wscript.Echo "Az Scope Name: " & objItem.AzScopeName
    Wscript.Echo "Az Store Name: " & objItem.AzStoreName
    Wscript.Echo "Cache Control Custom: " & objItem.CacheControlCustom
    Wscript.Echo "Cache Control Maximum Age: " & objItem.CacheControlMaxAge
    Wscript.Echo "Cache Control No Cache: " & objItem.CacheControlNoCache
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "CGI Timeout: " & objItem.CGITimeout
    Wscript.Echo "Create CGI With New Console: " & _
        objItem.CreateCGIWithNewConsole
    Wscript.Echo "Create Process As User: " & objItem.CreateProcessAsUser
    Wscript.Echo "Default Doc Footer: " & objItem.DefaultDocFooter
    Wscript.Echo "Default Logon Domain: " & objItem.DefaultLogonDomain
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Disable Static File Cache: " & _
        objItem.DisableStaticFileCache
    Wscript.Echo "Do Dynamic Compression: " & objItem.DoDynamicCompression
    Wscript.Echo "Don't Log: " & objItem.DontLog
    Wscript.Echo "Do Static Compression: " & objItem.DoStaticCompression
    Wscript.Echo "Enable Doc Footer: " & objItem.EnableDocFooter
    Wscript.Echo "Enable Reverse Dns: " & objItem.EnableReverseDns
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
    Wscript.Echo "Pool Idc Timeout: " & objItem.PoolIdcTimeout
    Wscript.Echo "Realm: " & objItem.Realm
    Wscript.Echo "Setting ID: " & objItem.SettingID
    Wscript.Echo "SSI Exec Disable: " & objItem.SSIExecDisable
    Wscript.Echo "Upload Read Ahead Size: " & _
        objItem.UploadReadAheadSize
    Wscript.Echo "Use Digest SSP: " & objItem.UseDigestSSP
    Wscript.Echo
Next

