' Description: Reports the site name for the local computer.


Set objADSysInfo = CreateObject("ADSystemInfo")

WScript.Echo "Current site name: " & objADSysInfo.SiteName

