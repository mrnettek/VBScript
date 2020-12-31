' Description: Disables the updating of the last access time for files and folders.


Set objShell = WScript.CreateObject("WScript.Shell")
strRegKey =  objShell.RegWrite _
    ("HKLM\System\CurrentControlSet\Control\FileSystem\" _
        & "NtfsDisableLastAccessUpdate" , 1, "REG_DWORD")

