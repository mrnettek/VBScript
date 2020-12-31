On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_TSEnvironmentSetting",,48)
For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "ClientWallPaper: " & objItem.ClientWallPaper
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "InitialProgramPath: " & objItem.InitialProgramPath
    Wscript.Echo "InitialProgramPolicy: " & objItem.InitialProgramPolicy
    Wscript.Echo "SettingID: " & objItem.SettingID
    Wscript.Echo "Startin: " & objItem.Startin
    Wscript.Echo "TerminalName: " & objItem.TerminalName
Next

