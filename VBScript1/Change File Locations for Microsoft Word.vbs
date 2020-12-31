Const wdUserTemplatesPath = 2
Const wdWorkgroupTemplatesPath = 3

Set objWord = CreateObject("Word.Application")
Set objOptions = objWord.Options

objOptions.DefaultFilePath(wdUserTemplatesPath) = "C:\Templates"
objOptions.DefaultFilePath(wdWorkgroupTemplatesPath) = "C:\Workgroup\Templates"

objWord.Quit
  


