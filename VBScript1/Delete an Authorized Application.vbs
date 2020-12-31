' Description: Deletes Freecell.exe from the list of authorized applications in the Windows Firewall current profile.


Set objFirewall = CreateObject("HNetCfg.FwMgr")
Set objPolicy = objFirewall.LocalPolicy.CurrentProfile

Set colApplications = objPolicy.AuthorizedApplications

errReturn = colApplications.Remove("c:\windows\system32\freecell.exe")

