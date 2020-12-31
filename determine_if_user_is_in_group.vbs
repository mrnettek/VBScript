' MrNetTek
' eddiejackson.net/blog
' 10/14/2019
' free for public use 
' free to claim as your own

Option Explicit
 
Dim arrMemberNames(), initialSize, i, strHolder, objADGroup
Dim strMemberName, strADUser, objADUser, strBool, objNetwork, strComputerUser
Dim strGroupMember
initialSize = 0
strBool = "FALSE"
 
Set objNetwork = CreateObject("Wscript.Network")
 
  
'SET AD GROUP
 Set objADGroup = GetObject("LDAP://CN=TheGroupName,DC=DomainName,DC=com")
 
'SET USER
strComputerUser = objNetwork.UserName
strComputerUser = LCase(strComputerUser)
 
 
'DETERMINE WHETHER USER IS IN GROUP 
For Each strADUser in objADGroup.Member
    Set objADUser = GetObject("LDAP://" & strADUser)    
    strGroupMember = LCase(objADUser.sAMAccountName)
    if strComputerUser = strGroupMember then
        strBool = "TRUE"
    end if  
    ReDim Preserve arrMemberNames(initialSize)
    arrMemberNames(initialSize) = objADUser.CN  
    initialSize = initialSize + 1
Next
 
If strBool = "TRUE" then msgbox "User was found in group!"
 
 
'CLEAR SESSION
initialSize = ""
i = ""
strHolder = ""
objADGroup = ""
strMemberName = ""
strADUser = ""
objADUser = ""
strBool = ""
 
'EXIT
WScript.Quit(0)