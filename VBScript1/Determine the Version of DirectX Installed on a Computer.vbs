Const HKEY_LOCAL_MACHINE = &H80000002

strComputer = "."
 
Set objRegistry=GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
 
strKeyPath = "Software\Microsoft\DirectX"
strValueName = "Version"

objRegistry.GetStringValue HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strValue

Select Case strValue
    Case "4.02.0095"
        strVersion = "1.0"
    Case "4.03.00.1096"
        strVersion = "2.0"
    Case "4.04.0068"
        strVersion = "3.0"
    Case "4.04.0069"
        strVersion = "3.0"
    Case "4.05.00.0155"
        strVersion = "5.0"
    Case "4.05.01.1721"
        strVersion = "5.0"
    Case "4.05.01.1998"
        strVersion = "5.0"
    Case "4.06.02.0436"
        strVersion = "6.0"
    Case "4.07.00.0700"
        strVersion = "7.0"
    Case "4.07.00.0716"
        strVersion = "7.0a"
    Case "4.08.00.0400"
        strVersion = "8.0"
    Case "4.08.01.0881"
        strVersion = "8.1"
    Case "4.08.01.0810"
        strVersion = "8.1"
    Case "4.09.0000.0900"
        strVersion = "9.0"
    Case "4.09.00.0900"
        strVersion = "9.0"
    Case "4.09.0000.0901"
        strVersion = "9.0a"
    Case "4.09.00.0901"
        strVersion = "9.0a"
    Case "4.09.0000.0902"
        strVersion = "9.0b"
    Case "4.09.00.0902"
        strVersion = "9.0b"
    Case "4.09.00.0904"
        strVersion = "9.0c"
    Case "4.09.0000.0904"
        strVersion = "9.0c"
End Select

Wscript.Echo strVersion
  


