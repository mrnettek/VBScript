' Description: Types the string "Test string." within a virtual machine named Windows 2000 Server.


On Error Resume Next

Set objVS = CreateObject("VirtualServer.Application")
Set objVM = objVS.FindVirtualMachine("Windows 2000 Server")

Set objKeyboard = objVM.Keyboard
errReturn = objKeyboard.TypeASCIIText("Test string.")

