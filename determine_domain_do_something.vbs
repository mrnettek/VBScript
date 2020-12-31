' MrNetTek
' eddiejackson.net/blog
' 10/14/2019
' free for public use 
' free to claim as your own

Option Explicit
 
Dim objWMISvc, ColItems, objItem, strComputerDomain
 
Set objWMISvc = GetObject( "winmgmts:\\.\root\cimv2" )
Set colItems = objWMISvc.ExecQuery( "Select * from Win32_ComputerSystem", ,48 )
For Each objItem in colItems
    strComputerDomain = objItem.Domain
        If objItem.PartOfDomain Then
            'DOMAIN 1
            if strComputerDomain = "DOMAIN1.com" then
                'do something here
                msgbox "Domain 1"
            'DOMAIN 2
            elseif strComputerDomain = "DOMAIN2.com" then
                'do something here
                msgbox "Domain 2"
            'EVERYONE ELSE
            else
                'do something here
                msgbox "Everyone Else"
            end if
        end if
next
 
'CLEAR SESSION
Set objWMISvc = Nothing
Set colItems = Nothing
objItem = ""
strComputerDomain = ""