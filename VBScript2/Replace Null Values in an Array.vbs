Dim arrNumbers(2, 2)

arrNumbers(0,1) = 100
arrNumbers(1,2) = 200
arrNumbers(2,1) = 300
arrNumbers(2,2) = 400

For x = 0 to 2
    For y = 0 to 2
        If arrNumbers(x, y) = "" Then
            arrNumbers(x, y) = "-"
        End If
    Next
Next
   
For x = 0 to 2
    For y = 0 to 2
        Wscript.Echo x & "," & y & ": " & arrNumbers(x, y)
    Next
Next
  


