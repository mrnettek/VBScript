' Description: Returns the property values of a Web file named W3SVC/2142295254/root/iisstart.htm. If the file is not already in the IIS metabase, the script “creates” it, a process that adds the file to the metabase,  allowing its properties to be modified.


On Error Resume Next
 
strComputer = "LocalHost"
Set objX = GetObject _
    ("IIS://" & strComputer & "/W3SVC/2142295254/root/iisstart.htm")

If Err.Number <> 0 Then
    Set objIIS = GetObject _
        ("IIS://" & strComputer & "/W3SVC/2142295254/root")
    Set objFile = objIIS.Create("IIsWebFile", "iisstart.htm")
    objFile.SetInfo
    objIIS.SetInfo
End If
 
Set objFile = GetObject _
    ("IIS://" & strComputer & "/W3SVC/2142295254/root/iisstart.htm")
 
Wscript.Echo "Access Execute: " & objFile.AccessExecute
Wscript.Echo "Access Flags: " & objFile.AccessFlags
Wscript.Echo "Access No Physical Directory : " & _ 
    objFile.AccessNoPhysicalDir
Wscript.Echo "Access No Remote Execute: " & _ 
    objFile.AccessNoRemoteExecute
Wscript.Echo "Access No Remote Read: " & objFile.AccessNoRemoteRead
Wscript.Echo "Access No Remote Script: " & _ 
    objFile.AccessNoRemoteScript
Wscript.Echo "Access No Remote Write: " & objFile.AccessNoRemoteWrite
Wscript.Echo "Access Read: " & objFile.AccessRead
Wscript.Echo "Access Script: " & objFile.AccessScript
Wscript.Echo "Access Source: " & objFile.AccessSource
Wscript.Echo "Access SSL: " & objFile.AccessSSL
Wscript.Echo "Access SSL 128: " & objFile.AccessSSL128
Wscript.Echo "Access SSL Flags: " & objFile.AccessSSLFlags
Wscript.Echo "Access SSL Map Certificate: " & _ 
    objFile.AccessSSLMapCert
Wscript.Echo "Access SSL Negotiate Certificate: " & _ 
    objFile.AccessSSLNegotiateCert
Wscript.Echo "Access SSL Require Certificate: " & _ 
    objFile.AccessSSLRequireCert
Wscript.Echo "Access Write: " & objFile.AccessWrite
For Each strACL in objFile.AdminACLBin
    Wscript.Echo "Admin ACL Bin: " & strACL
Next
Wscript.Echo "Anonymous Password Sync: " & _ 
    objFile.AnonymousPasswordSync
Wscript.Echo "Anonymous User Name: " & objFile.AnonymousUserName
Wscript.Echo "Anonymous User Password: " & objFile.AnonymousUserPass
Wscript.Echo "Authentication Anonymous: " & objFile.AuthAnonymous
Wscript.Echo "Authentication Basic: " & objFile.AuthBasic
Wscript.Echo "Authentication Flags: " & objFile.AuthFlags
Wscript.Echo "Authentication MD5: " & objFile.AuthMD5
Wscript.Echo "Authentication NTLM: " & objFile.AuthNTLM
Wscript.Echo "Authentication Passport: " & objFile.AuthPassport
Wscript.Echo "Authentication Persistence: " & objFile.AuthPersistence
Wscript.Echo "Authentication Persist Single Request: " & _ 
    objFile.AuthPersistSingleRequest
Wscript.Echo "Az Enable: " & objFile.AzEnable
Wscript.Echo "Az Impersonation Level: " & _ 
    objFile.AzImpersonationLevel
Wscript.Echo "Az Scope Name: " & objFile.AzScopeName
Wscript.Echo "Az Store Name: " & objFile.AzStoreName
Wscript.Echo "Cache Control Custom: " & objFile.CacheControlCustom
Wscript.Echo "Cache Control Maximum Age: " & _ 
    objFile.CacheControlMaxAge
Wscript.Echo "Cache Control No Cache: " & objFile.CacheControlNoCache
Wscript.Echo "CGI Timeout: " & objFile.CGITimeout
Wscript.Echo "Create CGI With New Console: " & _ 
    objFile.CreateCGIWithNewConsole
Wscript.Echo "Create Process As User: " & objFile.CreateProcessAsUser
Wscript.Echo "Default Doc Footer: " & objFile.DefaultDocFooter
Wscript.Echo "Default Logon Domain: " & objFile.DefaultLogonDomain
Wscript.Echo "Disable Static File Cache: " & _ 
    objFile.DisableStaticFileCache
Wscript.Echo "Do Dynamic Compression: " & _ 
    objFile.DoDynamicCompression
Wscript.Echo "Don't Log: " & objFile.DontLog
Wscript.Echo "Do Static Compression: " & objFile.DoStaticCompression
Wscript.Echo "Enable Doc Footer: " & objFile.EnableDocFooter
Wscript.Echo "Enable Reverse Dns: " & objFile.EnableReverseDns
For Each strHeader in objFile.HttpCustomHeaders
    Wscript.Echo "Http Custom Header: " & strHeader
Next
For Each strError in objFile.HttpErrors
    Wscript.Echo "Http Error: " & strError
Next
Wscript.Echo "Http Expires: " & objFile.HttpExpires
For Each strPic in objFile.HttpPics
    Wscript.Echo "Http Pic: " & strPic
Next
Wscript.Echo "Http Redirect: " & objFile.HttpRedirect
Wscript.Echo "Logon Method: " & objFile.LogonMethod
Wscript.Echo "Maximum Request Entity Allowed: " & _ 
    objFile.MaxRequestEntityAllowed
For Each strMap in objFile.MimeMap
    Wscript.Echo "Mime Map: " & strMap
Next
Wscript.Echo "Name: " & objFile.Name
Wscript.Echo "NT Authentication Providers: " & _ 
    objFile.NTAuthenticationProviders
Wscript.Echo "Passport Requires AD Mapping: " & _ 
    objFile.PassportRequireADMapping
Wscript.Echo "Pool IDC Timeout: " & objFile.PoolIdcTimeout
Wscript.Echo "Realm: " & objFile.Realm
For Each strHeader in objFile.RedirectHeaders
     Wscript.Echo "Redirect Header: " & strHeader
Next
For Each strMap in objFile.ScriptMaps
    Wscript.Echo "Script Map: " & strMap
Next
Wscript.Echo "SSI Exec Disable: " & objFile.SSIExecDisable
Wscript.Echo "Upload Read Ahead Size: " & objFile._ 
    UploadReadAheadSize
Wscript.Echo "Use Digest SSP: " & objFile.UseDigestSSP

