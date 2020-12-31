' Description: Configures a Terminal Services client so that mapped drives and printers are automatically connected at logon; also enables print jobs to be sent directly to the client’s local printer. These settings can be disabled by setting the appropriate constant value to 0 rather than 1.


Const CONNECT_CLIENT_DRIVES = 1
Const CONNECT_PRINTER = 1
Const USE_CLIENT_PRINTER = 1
 
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_TSClientSetting")

For Each objItem in colItems
    errResult = objItem.ConnectionSettings _
        (CONNECT_CLIENT_DRIVES, CONNECT_PRINTER, USE_CLIENT_PRINTER)
Next

