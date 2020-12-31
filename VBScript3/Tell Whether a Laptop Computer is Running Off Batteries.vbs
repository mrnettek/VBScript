strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\wmi")

Set colItems = objWMIService.ExecQuery("Select * From BatteryStatus Where Voltage > 0")

For Each objItem in colItems
    Wscript.Echo "Battery: " & objItem.InstanceName
    Wscript.Echo "On AC Power: " & objItem.PowerOnline
    Wscript.Echo "Battery is Discharging: " & objItem.Discharging
    Wscript.Echo "Battery is Charging: " & objItem.Charging
    Wscript.Echo "Remaining capacity: " & objItem.RemainingCapacity
Next

