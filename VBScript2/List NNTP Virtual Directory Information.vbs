' Description: Returns information about all the NNTP virtual directories on an IIS server.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsNntpVirtualDirSetting")
 
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
    Wscript.Echo "Admin ACL Bin: " & objItem.AdminACLBin
    Wscript.Echo "Content Indexed: " & objItem.ContentIndexed
    Wscript.Echo "Don't Log: " & objItem.DontLog
    Wscript.Echo "Ex Mdb Guid: " & objItem.ExMdbGuid
    Wscript.Echo "Fs Property Path: " & objItem.FsPropertyPath
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Path: " & objItem.Path
    Wscript.Echo "Setting ID: " & objItem.SettingID
    Wscript.Echo "UNC Password: " & objItem.UNCPassword
    Wscript.Echo "UNC User Name: " & objItem.UNCUserName
    Wscript.Echo "Vr Do Expire: " & objItem.VrDoExpire
    Wscript.Echo "Vr Driver Clsid: " & objItem.VrDriverClsid
    Wscript.Echo "Vr Driver Progid: " & objItem.VrDriverProgid
    Wscript.Echo "Vr Own Moderator: " & objItem.VrOwnModerator
    Wscript.Echo "Vr Use Account: " & objItem.VrUseAccount
    Wscript.Echo "Win32 Error: " & objItem.Win32Error
Next

