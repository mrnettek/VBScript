' MrNetTek
' eddiejackson.net/blog
' 10/14/2019
' free for public use 
' free to claim as your own

on error resume next

Dim varHTTP, varBinaryString, varFileName, varLink
 
set objShell = CreateObject("WScript.Shell")
  
Set varHTTP = CreateObject("Microsoft.XMLHTTP")
Set varBinaryString = CreateObject("Adodb.Stream")
 
varFileName = "Foo.exe"
varLink = "http://TheWebURL/" & varFileName
varHTTP.Open "GET", varLink, False
varHTTP.Send
 
 
'Sequencing...
CheckFile()
DownloadFile()
 
'add other stuff to do here
 
 
sub CheckFile()
    Select Case Cint(varHTTP.status)
        Case 200, 202, 302 
            'it exists
            msgbox "File exists!"
            Exit Sub
        Case Else
            'does not exist         
            msgbox "File does not exist!"
            WScript.quit
    End Select
end sub
 
sub DownloadFile()
        With varBinaryString
            .Type = 1 'my type has been set to binary
            .Open
            .Write varHTTP.responseBody
            .SaveToFile ".\" & varFileName, 2 'if exist, overwrite
        End With
        varBinaryString.close
        msgbox "Download complete!"
end sub