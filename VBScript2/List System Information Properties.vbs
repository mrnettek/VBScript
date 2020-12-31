' Description: Displays basic system information properties for Software Update Services.


Set objSysInfo = CreateObject("Microsoft.Update.SystemInfo")

Wscript.Echo "OEM hardware support link: " & objSysInfo.OEMHardwareSupportLink
Wscript.Echo "Reboot required: " & objSysInfo.RebootRequired

