' Description: Modifies the system restore configuration values on a computer, setting the global interval to 100,000 seconds; the life interval to 8,000,000 seconds; and the session interval to 500,000 seconds.


Const GLOBAL_INTERVAL_IN_SECONDS = 100000
Const LIFE_INTERVAL_IN_SECONDS = 8000000
Const SESSION_INTERVAL_IN_SECONDS = 500000
 
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\default")

Set objItem = objWMIService.Get("SystemRestoreConfig='SR'")
objItem.DiskPercent = 10
objItem.RPGlobalInterval = GLOBAL_INTERVAL_IN_SECONDS
objItem.RPLifeInterval = LIFE_INTERVAL_IN_SECONDS
objItem.RPSessionInterval = SESSION_INTERVAL_IN_SECONDS
objItem.Put_

