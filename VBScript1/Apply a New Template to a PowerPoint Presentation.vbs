Set objPPT = CreateObject("PowerPoint.Application")
objPPT.Visible = True

Set objPresentation = objPPT.Presentations.Open("C:\Presentations\Test.ppt")

objPresentation.ApplyTemplate _
    ("C:\Program Files\Microsoft Office\Templates\Presentation Designs\Ocean.pot")

objPresentation.Save

objPPT.Quit
  


