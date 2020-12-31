' Description: Lists security information for Virtual Server.


On Error Resume Next

Set objVS = CreateObject("VirtualServer.Application")
Set objSecurity = objVS.Security

Wscript.Echo "Group name: " & objSecurity.GroupName
Wscript.Echo "Group SID: " & objSecurity.GroupSID
Wscript.Echo "Owner name: " & objSecurity.OwnerName
Wscript.Echo "Owner name: " & objSecurity.OwnerSID

