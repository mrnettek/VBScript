' Description: Returns information about the Terminal Service client environment settings configured on a computer.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
("Select * from Win32_TSEnvironmentSetting")

For Each objItem in colItems
    Wscript.Echo "Client wallpaper: " & objItem.ClientWallpaper
    Wscript.Echo "Initial program path: " & objItem.InitialProgramPath
    Wscript.Echo "Initial program policy: " & objItem.InitialProgramPolicy
    Wscript.Echo "Setting ID: " & objItem.SettingID
    Wscript.Echo "Start in: " & objItem.Startin
    Wscript.Echo "Terminal name: " & objItem.TerminalName
    Wscript.Echo
Next

