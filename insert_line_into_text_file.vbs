' MrNetTek
' eddiejackson.net/blog
' 10/14/2019
' free for public use 
' free to claim as your own

Option Explicit
 
'call function
'file to modify, text to search for, text to add
AddLine "C:\temp\test.txt","[Logon Profiles]","XXXXXXXXXXXXXXXX"
 
WScript.Quit
 
 
Function AddLine(sourcePath, searchString, addText)
 
    Dim strFileSourcePath, strFileTargetPath, objSource, objTarget, objInput, objOutput, strLine, objShell
 
    Const ForReading = 1, ForWriting = 2
 
    Set objShell = WScript.CreateObject("WScript.Shell")
 
    ' clear session
 
    objShell.Run "cmd /c del /q c:\temp\test.new",0,true
 
    WScript.Sleep 1000
 
 
    strFileSourcePath = sourcePath
 
    strFileTargetPath = "c:\temp\test.new"
 
    Set objSource = CreateObject("scripting.filesystemobject")
 
    Set objTarget = CreateObject("scripting.filesystemobject")
 
    Set objInput = objSource.OpenTextFile(strFileSourcePath,ForReading,-1) 
 
    Set objOutput = objSource.OpenTextFile(strFileTargetPath,ForWriting,True,0) 
 
    Do While objInput.AtEndOfStream <> true
 
        strLine = objInput.ReadLine
 
        ' look for this
        if strLine = searchString then
   
            ' if found, do this
            objOutput.WriteLine strLine
 
            objOutput.WriteLine addText
        else
 
            ' if not found, do this
            objOutput.WriteLine strLine
 
        end if   
 
    Loop
 
    objInput.Close
 
    objOutput.Close
 
 
    ' create updated text file
    objShell.Run "cmd /c copy /y c:\temp\test.new " & chr(34) & sourcePath & chr(34),0,true
 
    WScript.Sleep 1000
 
    ' clear session
    objShell.Run "cmd /c del /q c:\temp\test.new",0,true
 
    Set objSource = Nothing
     
    Set objTarget = Nothing
 
    Set strLine = Nothing
 
    Set objInput = Nothing
 
    Set objOutput = Nothing
 
end function