' Description: Demonstration script that unloads the application in the W3SVC/1/ROOT/tsweb virtual Web directory.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * From IIsWebVirtualDir Where Name = " & _
        "'W3SVC/1/ROOT/tsweb'")

For Each objItem in colItems
   objItem.AppUnload(True)
Next

