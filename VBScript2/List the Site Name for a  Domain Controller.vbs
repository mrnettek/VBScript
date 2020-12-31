' Description: Reports the site name for a specified  computer.


strDcName = "atl-dc-01"
Set objADSysInfo = CreateObject("ADSystemInfo")

strDcSiteName = objADSysInfo.GetDCSiteName(strDcName)
WScript.Echo "DC Site Name: " & strDcSiteName

