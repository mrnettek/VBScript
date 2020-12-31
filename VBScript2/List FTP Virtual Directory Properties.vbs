' Description: Lists the properties of all the virtual FTP directories on an IIS server.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsFtpVirtualDirSetting")

For Each objItem in colItems
    Wscript.Echo "Access Execute: " & objItem.AccessExecute
    Wscript.Echo "Access Flags: " & objItem.AccessFlags
    Wscript.Echo "Access No Physical Directory: " &  _
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
    Wscript.Echo "Access Write: " & objItem.AccessWrite
    Wscript.Echo "Don't Log: " & objItem.DontLog
    Wscript.Echo "FTP Directory Browse Show Long Date: " & _
        objItem.FtpDirBrowseShowLongDate
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Path: " & objItem.Path
    Wscript.Echo "Setting ID: " & objItem.SettingID
    Wscript.Echo "UNC Password: " & objItem.UNCPassword
    Wscript.Echo "UNC User Name: " & objItem.UNCUserName
    Wscript.Echo "Win32 Error: " & objItem.Win32Error
Next

