' Description: Retrieves the HTML source for the Web page http://www.microsoft.com. This script contributed by Maxim Stepin of Microsoft.


url="http://www.microsoft.com"
Set objHTTP = CreateObject("MSXML2.XMLHTTP")

Call objHTTP.Open("GET", url, FALSE)
objHTTP.Send

WScript.Echo(objHTTP.ResponseText)

