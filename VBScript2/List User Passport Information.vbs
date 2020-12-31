' Description: Returns information about the .NET passport for the user currently logged-on to a computer.


Set objUser = CreateObject("UserAccounts.PassportManager")
Wscript.Echo "Current Passport: " & objUser.CurrentPassport
Wscript.Echo "Member services URL: " & objUser.MemberServicesURL

