
Dim Test : Test = "Hello World"

'Encode with ROT13
Test = ROT13(Test)
msgbox Test

'Decode with ROT13
Test = ROT13(Test)
msgbox Test


' --------------------------
Public Function ROT13(sText)

Const cAlphabet = "abcdefghijklmnopqrstuvwxyzabcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZABCDEFGHIJKLMNOPQRSTUVWXYZ"
Dim sROT13, lPos

    For lPos = 1 To Len(sText)
        iChar = Instr(cAlphabet, Mid(sText, lPos, 1))
        If iChar = 0 Then
            sROT13 = sROT13 & Mid(sText, lPos, 1)
        Else
            sROT13 = sROT13 & Mid(cAlphabet, iChar + 13, 1)
        End If
    Next

    ROT13 = sROT13

End Function
