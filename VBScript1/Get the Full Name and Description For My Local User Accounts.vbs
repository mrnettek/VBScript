strComputer = "."

Set colAccounts = GetObject("WinNT://" & strComputer & "")
colAccounts.Filter = Array("user")

For Each objUser In colAccounts
    Wscript.Echo objUser.Name 
    Wscript.Echo objUser.FullName 
    Wscript.Echo Description 
    Wscript.Echo  
Next
  


