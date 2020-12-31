Const WbemAuthenticationLevelPktPrivacy = 6

strComputer = "atl-ws-01"
strNamespace = "root\cimv2"
strUser = "Administrator"
strPassword = "4rTGh2#1"

Set objWbemLocator = CreateObject("WbemScripting.SWbemLocator")
Set objWMIService = objwbemLocator.ConnectServer _
    (strComputer, strNamespace, strUser, strPassword)
objWMIService.Security_.authenticationLevel = WbemAuthenticationLevelPktPrivacy

Set colItems = objWMIService.ExecQuery _
    ("Select * From Win32_OperatingSystem")
For Each objItem in ColItems
    Wscript.Echo strComputer & ": " & objItem.Caption
Next
  


