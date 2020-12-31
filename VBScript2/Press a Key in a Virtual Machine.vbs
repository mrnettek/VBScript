' Description: Presses the left CTRL and the Escape key within a virtual machine named Windows 2000 Server.


On Error Resume Next

Set objVS = CreateObject("VirtualServer.Application")
Set objVM = objVS.FindVirtualMachine("Windows 2000 Server")

Set objKeyboard = objVM.Keyboard
errReturn = objKeyboard.PressKey("Key_LeftCtrl")
errReturn = objKeyboard.PressKey("Key_Escape")

