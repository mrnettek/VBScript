' Description: Enables File and Printer Sharing on a computer running Windows XP Service Pack 2.


Set objFirewall = CreateObject("HNetCfg.FwMgr")
Set objPolicy = objFirewall.LocalPolicy.CurrentProfile

Set colServices = objPolicy.Services
Set objService = colServices.Item(0)
objService.Enabled = TRUE

