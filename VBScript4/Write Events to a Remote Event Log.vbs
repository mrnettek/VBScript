' Description: Writes an event to the Application event log on a remote computer named PrimaryServer.


Const EVENT_SUCCESS = 0

Set objShell = Wscript.CreateObject("Wscript.Shell")

objShell.LogEvent EVENT_SUCCESS, _
    "Payroll application successfully installed." , "\\PrimaryServer"

