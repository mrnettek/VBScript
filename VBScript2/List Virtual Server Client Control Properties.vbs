' Description: Lists Virtual Server client control properties.


On Error Resume Next

Set objClientControl = CreateObject("MSVMRCActiveXClient.VMRCClientControl")
Wscript.Echo "Administrator mode: " & objClientControl.AdministratorMode
Wscript.Echo "Autosize: " & objClientControl.AutoSize
Wscript.Echo "Enabled: " & objClientControl.Enabled
Wscript.Echo "Host key: " & objClientControl.HostKey
Wscript.Echo "Menu back color: " & objClientControl.MenuBackColor
Wscript.Echo "Menu enabled: " & objClientControl.MenuEnabled
Wscript.Echo "Menu font color: " & objClientControl.MenuFontColor
Wscript.Echo "Menu font size: " & objClientControl.MenuFontSize
Wscript.Echo "Menu height: " & objClientControl.MenuHeight
Wscript.Echo "Ready state: " & objClientControl.ReadyState
Wscript.Echo "Reduced-colors mode: " & objClientControl.ReducedColorsMode
Wscript.Echo "Server address: " & objClientControl.ServerAddress
Wscript.Echo "Server display height: " & objClientControl.ServerDisplayHeight
Wscript.Echo "Server display name: " & objClientControl.ServerDisplayName
Wscript.Echo "Server display width: " & objClientControl.ServerDisplayWidth
Wscript.Echo "Server port: " & objClientControl.ServerPort
Wscript.Echo "State: " & objClientControl.State
Wscript.Echo "View-only mode: " & objClientControl.ViewOnlyMode

