' Description: Installs a hypothetical software program (using a Windows Installer package) on a remote computer. Requires delegation for the computer and user accounts involved in the procedure.


Const wbemImpersonationLevelDelegate = 4

Set objWbemLocator = CreateObject("WbemScripting.SWbemLocator")
Set objConnection = objwbemLocator.ConnectServer _
    ("WebServer", "root\cimv2", "fabrikam\administrator", _
         "password", , "kerberos:WebServer")
objConnection.Security_.ImpersonationLevel = wbemImpersonationLevelDelegate

Set objSoftware = objConnection.Get("Win32_Product")
errReturn = objSoftware.Install("\\atl-dc-02\scripts\1561_lab.msi",,True)

