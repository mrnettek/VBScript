' Description: Displays NNTP expiration setting information.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsNntpExpireSetting")
 
For Each objItem in colItems
    For Each strGroup in objItem.ExpireNewsgroups
        Wscript.Echo "Expire Newsgroups: " & strGroup
    Next
    Wscript.Echo "Expire Policy Name: " & objItem.ExpirePolicyName
    Wscript.Echo "Expire Space: " & objItem.ExpireSpace
    Wscript.Echo "Expire Time: " & objItem.ExpireTime
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Setting ID: " & objItem.SettingID
Next

