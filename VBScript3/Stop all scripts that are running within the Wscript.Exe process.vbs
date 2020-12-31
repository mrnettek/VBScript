
Dim oWMIService : Set oWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Dim cProcessList : Set cProcessList = oWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Name = 'wscript.exe'")

WScript.Echo "Stopping all your scripts that are running within the Wscript.Exe process"

For Each oProcess in cProcessList
    oProcess.Terminate()
Next
