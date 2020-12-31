On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
      & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

strFile = "C:\Scripts\My_script.vbs"

Set colItems = objWMIService.ExecQuery _
    ("ASSOCIATORS OF {Win32_LogicalFileSecuritySetting='" & strFile & "'}" _ 
        & " WHERE AssocClass=Win32_LogicalFileOwner ResultRole=Owner")

For Each objItem in colItems
    Wscript.Echo objItem.ReferencedDomainName
    Wscript.Echo objItem.AccountName
Next
  


