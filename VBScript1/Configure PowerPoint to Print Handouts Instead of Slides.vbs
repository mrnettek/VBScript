Const ppPrintOutputThreeSlideHandouts = 3

Set objPPT = CreateObject("PowerPoint.Application")
objPPT.Visible = True

Set objPresentation = objPPT.Presentations.Add
Set objSlide = objPresentation.Slides.Add(1,1)

Set objOptions = objPresentation.PrintOptions
objOptions.OutputType = ppPrintOutputThreeSlideHandouts
  


