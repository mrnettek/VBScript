' Description: Enumerates Internet Connection Firewall connection properties.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & _
        "\root\microsoft\homenet")

Set colItems = objWMIService.ExecQuery _
    ("Select * from HNet_ConnectionProperties")

For Each objItem in colItems
    Wscript.Echo "Connection: " & objItem.Connection
    Wscript.Echo "Is Bridge: " & objItem.IsBridge
    Wscript.Echo "Is Bridge Member: " & objItem.IsBridgeMember
    Wscript.Echo "Is Firewalled: " & objItem.IsFirewalled
    Wscript.Echo "Is ICS Private: " & objItem.IsICSPrivate
    Wscript.Echo "Is ICS Public: " & objItem.IsICSPublic
Next

