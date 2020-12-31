' Description: Returns information about current activity on the fax server atl-dc-02.


Set objFaxServer = CreateObject("FaxComEx.FaxServer")
objFaxServer.Connect "atl-dc-02"

Set objfaxActivity = objFaxServer.Activity

Wscript.Echo "Incoming messages: " & objFaxActivity.IncomingMessages
Wscript.Echo "Outgoing messages: " & objFaxActivity.OutgoingMessages
Wscript.Echo "Queued messages: " & objFaxActivity.QueuedMessages
Wscript.Echo "Routing messages: " & objFaxActivity.RoutingMessages

