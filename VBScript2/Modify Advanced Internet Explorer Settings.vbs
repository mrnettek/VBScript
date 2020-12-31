' Description: Demonstration script that disables the Download Complete message box that typically appears after downloading a file using Internet Explorer.


On Error Resume Next

Const HKEY_CURRENT_USER = &H80000001

strComputer = "."
strValue = "no"

Set objReg = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}\\" & strComputer & _
        "\root\default:StdRegProv")

strKeyPath = "Software\Microsoft\Internet Explorer\Main"
objReg.SetStringValue HKEY_CURRENT_USER, strKeyPath, _
    "NotifyDownloadComplete",strValue

