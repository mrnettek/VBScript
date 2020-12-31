On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_SerialPortConfiguration",,48)
For Each objItem in colItems
    Wscript.Echo "AbortReadWriteOnError: " & objItem.AbortReadWriteOnError
    Wscript.Echo "BaudRate: " & objItem.BaudRate
    Wscript.Echo "BinaryModeEnabled: " & objItem.BinaryModeEnabled
    Wscript.Echo "BitsPerByte: " & objItem.BitsPerByte
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "ContinueXMitOnXOff: " & objItem.ContinueXMitOnXOff
    Wscript.Echo "CTSOutflowControl: " & objItem.CTSOutflowControl
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "DiscardNULLBytes: " & objItem.DiscardNULLBytes
    Wscript.Echo "DSROutflowControl: " & objItem.DSROutflowControl
    Wscript.Echo "DSRSensitivity: " & objItem.DSRSensitivity
    Wscript.Echo "DTRFlowControlType: " & objItem.DTRFlowControlType
    Wscript.Echo "EOFCharacter: " & objItem.EOFCharacter
    Wscript.Echo "ErrorReplaceCharacter: " & objItem.ErrorReplaceCharacter
    Wscript.Echo "ErrorReplacementEnabled: " & objItem.ErrorReplacementEnabled
    Wscript.Echo "EventCharacter: " & objItem.EventCharacter
    Wscript.Echo "IsBusy: " & objItem.IsBusy
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Parity: " & objItem.Parity
    Wscript.Echo "ParityCheckEnabled: " & objItem.ParityCheckEnabled
    Wscript.Echo "RTSFlowControlType: " & objItem.RTSFlowControlType
    Wscript.Echo "SettingID: " & objItem.SettingID
    Wscript.Echo "StopBits: " & objItem.StopBits
    Wscript.Echo "XOffCharacter: " & objItem.XOffCharacter
    Wscript.Echo "XOffXMitThreshold: " & objItem.XOffXMitThreshold
    Wscript.Echo "XOnCharacter: " & objItem.XOnCharacter
    Wscript.Echo "XOnXMitThreshold: " & objItem.XOnXMitThreshold
    Wscript.Echo "XOnXOffInFlowControl: " & objItem.XOnXOffInFlowControl
    Wscript.Echo "XOnXOffOutFlowControl: " & objItem.XOnXOffOutFlowControl
Next

