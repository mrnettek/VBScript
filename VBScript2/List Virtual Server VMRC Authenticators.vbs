' Description: Lists Virtual Server VMRC authenticators.


On Error Resume Next

Set objVS = CreateObject("VirtualServer.Application")
Set colAuthenticators = objVS.VMRCAuthenticators
For Each objAuthenticator in colAuthenticators
    Wscript.Echo "Name: " & objAuthenticator.Name
    Wscript.Echo "Description: " & objAuthenticator.Description
    Wscript.Echo "Type: " & objAuthenticator.Type
    Wscript.Echo
Next

