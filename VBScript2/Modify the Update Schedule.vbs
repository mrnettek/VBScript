' Description: Demonstration script that modifies the software update schedule.


Set objAutoUpdate = CreateObject("Microsoft.Update.AutoUpdate")
Set objSettings = objAutoUpdate.Settings

objSettings.ScheduledInstallationDay = 3
objSettings.ScheduledInstallationTime = 4

objSettings.Save

