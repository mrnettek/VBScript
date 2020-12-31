' Description: Locates Virtual Server access rights for the Administrators group.


On Error Resume Next

Set objVS = CreateObject("VirtualServer.Application")
Set objSecurity = objVS.Security

Set objAccessRights = objSecurity.FindEntry("Administrators",0)
Wscript.Echo objAccessRights.Name

