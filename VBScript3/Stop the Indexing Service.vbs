' Description: Stops the Indexing Service on the local computer.


On Error Resume Next

Set objAdminIS = CreateObject("Microsoft.ISAdm")
objAdminIS.Stop()

