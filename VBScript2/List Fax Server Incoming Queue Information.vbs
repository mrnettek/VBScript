' Description: Indicates whether the incoming queue is blocked on the fax server atl-dc-02.


Set objFaxServer = CreateObject("FaxComEx.FaxServer")
objFaxServer.Connect "atl-dc-02"

Set objFolder = objFaxServer.Folders

Set objIncomingQueue = objFolder.IncomingQueue
Wscript.Echo "Blocked: " & objIncomingQueue.Blocked

