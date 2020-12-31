' Description: Retrieves NTFS file system settings from the registry.


On Error Resume Next

Set objShell = WScript.CreateObject("WScript.Shell")

strRegKey =  objShell.RegRead _
    ("HKLM\System\CurrentControlSet\Control\FileSystem\" _
        & "NtfsDisable8dot3NameCreation")
If IsNull(strRegKey) Then
    Wscript.Echo "No value set for disabling 8.3 file name creation."
ElseIf strRegKey = 1 Then
    WScript.Echo "No 8.3 file names will be created for new files."
ElseIf strRegKey = 0 Then
    Wscript.Echo "8.3 file names will be created for new files."
End If

strRegKey = Null
strRegKey =  objShell.RegRead _
    ("HKLM\System\CurrentControlSet\Control\FileSystem\" _
        & "NtfsAllowExtendedCharacterIn8Dot3Name")
If IsNull(strRegKey) Then
    Wscript.Echo "No value set for allowing extended characters in " _
       & " 8.3 file names."
ElseIf strRegKey = 1 Then
    WScript.Echo "Extended characters are permitted in 8.3 file names."
ElseIf strRegKey = 0 Then
    Wscript.Echo "Extended characters not permitted in 8.3 file names."
End If

strRegKey = Null
strRegKey =  objShell.RegRead _
    ("HKLM\System\CurrentControlSet\Control\FileSystem\" _
        & "NtfsMftZoneReservation")
If IsNull(strRegKey) Then
    Wscript.Echo "No value set for reserving the MFT zone."
ElseIf strRegKey = 1 Then
    WScript.Echo _
        "One-eighth of the disk has been reserved for the MFT zone."
ElseIf strRegKey = 2 Then
    Wscript.Echo "One-fourth of the disk reserved for the MFT zone."
ElseIf strRegKey = 3 Then
    Wscript.Echo "Three-eighths of the disk reserved for the MFT zone."
ElseIf strRegKey = 4 Then
    Wscript.Echo "One half of the disk reserved for the MFT zone."
End If

strRegKey = Null
strRegKey =  objShell.RegRead _
    ("HKLM\System\CurrentControlSet\Control\FileSystem\" _
        & "NtfsDisableLastAccessUpdate")
If IsNull(strRegKey) Then
    Wscript.Echo "No value set for disabling the last access update " _
        & "for files and folder."
ElseIf strRegKey = 1 Then
    WScript.Echo "The last access timestamp will not be updated on files " _
        & "and folders."
ElseIf strRegKey = 0 Then
    Wscript.Echo "The last access timestamp updated on files and " _
         & "folders."
End If

strRegKey = Null
strRegKey =  objShell.RegRead _
    ("HKLM\System\CurrentControlSet\Control\FileSystem\Win31FileSystem")
If IsNull(strRegKey) Then
    Wscript.Echo "No value set for using long file names and " _
        & "timestamps."
ElseIf strRegKey = 1 Then
    WScript.Echo "Long file names and extended timestamps are used."
ElseIf strRegKey = 0 Then
    Wscript.Echo "Long file names and extended timestamps are not used."
End If

