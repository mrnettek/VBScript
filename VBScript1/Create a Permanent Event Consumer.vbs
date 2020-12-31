' Description: Creates a permanent event consumer for monitoring changes in service status.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\default")

Set objConsumerType = objWMIService.get("SMTPEventConsumer")
Set objConsumer = objConsumerType.SpawnInstance_
objConsumer.Name = "Service Monitor Consumer"
objConsumer.Message = "A service has changed state."
objConsumer.SMTPServer = "mailserver.fabrikam.com"
objConsumer.Subject = "Service state change"
objConsumer.ToLine = "administrator@fabrikam.com"
objConsumer.Put_

