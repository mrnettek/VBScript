' Description: Sample HTML code for adding a timer to a Web page, This timer calls a function named RunScript every 5 seconds (5000 milliseconds).


Sub Window_OnLoad
    iTimerID = window.setInterval("RunScript", 5000, "VBScript")
End Sub

