' Description: Returns network login information for all the users of a computer.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_NetworkLoginProfile")

For Each objItem in colItems
    dtmWMIDate = objItem.AccountExpires
    strReturn = WMIDateStringToDate(dtmWMIDate)
    Wscript.Echo "Account Expires: " & strReturn
    Wscript.Echo "Authorization Flags: " & objItem.AuthorizationFlags
    Wscript.Echo "Bad Password Count: " & objItem.BadPasswordCount
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "CodePage: " & objItem.CodePage
    Wscript.Echo "Comment: " & objItem.Comment
    Wscript.Echo "Country Code: " & objItem.CountryCode
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Flags: " & objItem.Flags
    Wscript.Echo "Full Name: " & objItem.FullName
    Wscript.Echo "Home Directory: " & objItem.HomeDirectory
    Wscript.Echo "Home Directory Drive: " & objItem.HomeDirectoryDrive
    dtmWMIDate = objItem.LastLogoff
    strReturn = WMIDateStringToDate(dtmWMIDate)
    Wscript.Echo "Last Logoff: " & strReturn
    dtmWMIDate = objItem.LastLogon
    strReturn = WMIDateStringToDate(dtmWMIDate)
    Wscript.Echo "Last Logon: " & strReturn
    Wscript.Echo "Logon Hours: " & objItem.LogonHours
    Wscript.Echo "Logon Server: " & objItem.LogonServer
    Wscript.Echo "Maximum Storage: " & objItem.MaximumStorage
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Number Of Logons: " & objItem.NumberOfLogons
    Wscript.Echo "Password Age: " & objItem.PasswordAge
    dtmWMIDate = objItem.PasswordExpires
    strReturn = WMIDateStringToDate(dtmWMIDate)
    Wscript.Echo "Password Expires: " & strReturn
    Wscript.Echo "Primary Group ID: " & objItem.PrimaryGroupId
    Wscript.Echo "Privileges: " & objItem.Privileges
    Wscript.Echo "Profile: " & objItem.Profile
    Wscript.Echo "Script Path: " & objItem.ScriptPath
    Wscript.Echo "Setting ID: " & objItem.SettingID
    Wscript.Echo "Units Per Week: " & objItem.UnitsPerWeek
    Wscript.Echo "User Comment: " & objItem.UserComment
    Wscript.Echo "User Id: " & objItem.UserId
    Wscript.Echo "User Type: " & objItem.UserType
    Wscript.Echo "Workstations: " & objItem.Workstations
    Wscript.Echo
Next
 
Function WMIDateStringToDate(dtmWMIDate)
    If Not IsNull(dtmWMIDate) Then
    WMIDateStringToDate = CDate(Mid(dtmWMIDate, 5, 2) & "/" & _
         Mid(dtmWMIDate, 7, 2) & "/" & Left(dtmWMIDate, 4) _
             & " " & Mid (dtmWMIDate, 9, 2) & ":" & _
                 Mid(dtmWMIDate, 11, 2) & ":" & Mid(dtmWMIDate, 13, 2))
    End If
End Function

