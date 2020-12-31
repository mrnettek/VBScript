' Description: Returns information about the local user accounts found on a computer.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_UserAccount Where LocalAccount = True")

For Each objItem in colItems
    Wscript.Echo "Account Type: " & objItem.AccountType
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Disabled: " & objItem.Disabled
    Wscript.Echo "Domain: " & objItem.Domain
    Wscript.Echo "Full Name: " & objItem.FullName
    Wscript.Echo "Local Account: " & objItem.LocalAccount
    Wscript.Echo "Lockout: " & objItem.Lockout
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Password Changeable: " & objItem.PasswordChangeable
    Wscript.Echo "Password Expires: " & objItem.PasswordExpires
    Wscript.Echo "Password Required: " & objItem.PasswordRequired
    Wscript.Echo "SID: " & objItem.SID
    Wscript.Echo "SID Type: " & objItem.SIDType
    Wscript.Echo "Status: " & objItem.Status
    Wscript.Echo
Next

