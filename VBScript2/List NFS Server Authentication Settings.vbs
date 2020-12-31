' Description: Displays Services for UNIX NFS server authentication settings.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\sfuadmin")

Set colItems = objWMIService.ExecQuery _
    ("Select * from NFSServer_Authenticate")

For Each objItem in colItems
    Wscript.Echo "Authentication Type: " & objItem.AuthType
    Wscript.Echo "Key Name: " & objItem.KeyName
    Wscript.Echo "NIS Domain: " & objItem.NISDomain
    Wscript.Echo "NIS Server: " & objItem.NISServer
    Wscript.Echo "NT Domain: " & objItem.NTDomain
    Wscript.Echo "PC NFSD Server: " & objItem.PCNFSDServer
    Wscript.Echo
Next

