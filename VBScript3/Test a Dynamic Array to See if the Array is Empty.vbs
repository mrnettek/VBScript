On Error Resume Next

Dim arrTest()
ReDim Preserve arrTest(0)

intUpper = Ubound(arrTest)

If Err = 0 Then
    Wscript.Echo "This array is not empty."
Else
    Wscript.Echo "This array is empty."
    Err.Clear
End If

Dim arrTest2()

intUpper2 = Ubound(arrTest2)

If Err = 0 Then
    Wscript.Echo "This array is not empty."
Else
    Wscript.Echo "This array is empty."
    Err.Clear
End If
  


