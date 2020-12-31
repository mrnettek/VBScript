strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_OperatingSystem")

For Each objItem in colItems
    If InStr(objItem.Caption, "Server 2003") Then
        strText = objItem.Caption 
        If objItem.ServicePackMajorVersion = "1" Then
            strText = strText & ", Service Pack 1"
        End If
        Wscript.Echo strText
    Else
        Wscript.Echo "This computer is not running Windows Server 2003."
    End If
Next
  


