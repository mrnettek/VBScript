' Description: Demonstration script that modifies an FTP virtual directory property (FtpDirBrowseShowLongDate) in the IIS metabase.


strComputer = "LocalHost"
Set objDirectory = GetObject("IIS://" & strComputer & _
    "/MSFTPSVC/1012388136/root")

objDirectory.FtpDirBrowseShowLongDate = TRUE
objDirectory.SetInfo

