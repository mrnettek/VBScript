' Description: Displays mapper settings for Services for UNIX.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\sfuadmin")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Mapper_Settings")

For Each objItem in colItems
    Wscript.Echo "Additional Map Definitions: " & _
        objItem.AdditionalMapDefinitions
    Wscript.Echo "Advanced Group Maps: " & _
        objItem.AdvancedGroupMaps
    Wscript.Echo "Advanced User Maps: " & _
        objItem.AdvancedUserMaps
    Wscript.Echo "Anonymous GID: " & objItem.AnonymousGid
    Wscript.Echo "Anonymous UID: " & objItem.AnonymousUid
    Wscript.Echo "Anonymous Unix User: " & _
        objItem.AnonymousUnixUser
    Wscript.Echo "Authentication Type: " & objItem.AuthType
    Wscript.Echo "Backup File Name: " & objItem.BackupFileName
    Wscript.Echo "Default: " & objItem.Default
    Wscript.Echo "Group File Name: " & objItem.GroupFileName
    Wscript.Echo "Key Name: " & objItem.KeyName
    Wscript.Echo "Logging Level: " & objItem.LoggingLevel
    Wscript.Echo "Map File Name: " & objItem.MapFileName
    Wscript.Echo "NIS Domain: " & objItem.NisDomain
    Wscript.Echo "NIS Server: " & objItem.NisServer
    Wscript.Echo "NT Domain: " & objItem.NTDomain
    Wscript.Echo "NT Domain2: " & objItem.NTDomain2
    Wscript.Echo "Password File Name: " & objItem._
        PasswdFileName
    Wscript.Echo "Refresh Interval: " & _
        objItem.RefreshInterval
    Wscript.Echo "Restore File Name: " & _
        objItem.RestoreFileName
    Wscript.Echo "Security: " & objItem.Security
    Wscript.Echo "Server Type: " & objItem.ServerType
    Wscript.Echo "Write Block: " & objItem.WriteBlock
    Wscript.Echo
Next

