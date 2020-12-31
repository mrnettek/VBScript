Private lngTrack
Private arrLongConversion(4)
Private arrSplit64(63)

Private Const OFFSET_4 = 4294967296
Private Const MAXINT_4 = 2147483647

Private Const S11 = 7
Private Const S12 = 12
Private Const S13 = 17
Private Const S14 = 22
Private Const S21 = 5
Private Const S22 = 9
Private Const S23 = 14
Private Const S24 = 20
Private Const S31 = 4
Private Const S32 = 11
Private Const S33 = 16
Private Const S34 = 23
Private Const S41 = 6
Private Const S42 = 10
Private Const S43 = 15
Private Const S44 = 21


' ------ Call MD5 Hash Function for your file here !
msgbox MD5FileHash("C:\Windows\system32\cmd.exe")


' ---------------------------
Public Function MD5FileHash(strFile)

	Dim strMD5 : strMD5 = ""
	Dim ofso : Set ofso = CreateObject("Scripting.FileSystemObject")

	If ofso.FileExists(strFile) then

		strMD5 = BinaryToString(ReadTextFile(strFile, ""))

		MD5FileHash = CalculateMD5(strMD5)

	Else

		MD5FileHash = strFile & VbCrLf & "Error: File not found"

	End if

End Function


' --------------------------------------
Function ReadTextFile(FileName, CharSet)

	Const adTypeText = 2
	Dim BinaryStream : Set BinaryStream = CreateObject("ADODB.Stream")

	BinaryStream.Type = adTypeText

	If Len(CharSet) > 0 Then

		BinaryStream.CharSet = CharSet

	End If

	BinaryStream.Open
	BinaryStream.LoadFromFile FileName

	ReadTextFile = BinaryStream.ReadText

End Function

' -----------------------------
Function BinaryToString(Binary)

Dim cl1, cl2, cl3, pl1, pl2, pl3
Dim L
	cl1 = 1
	cl2 = 1
	cl3 = 1
	L = LenB(Binary)

	Do While cl1<=L

		pl3 = pl3 & Chr(AscB(MidB(Binary,cl1,1)))
		cl1 = cl1 + 1
		cl3 = cl3 + 1

		If cl3>300 Then
			pl2 = pl2 & pl3
			pl3 = ""
			cl3 = 1
			cl2 = cl2 + 1

			If cl2>200 Then

				pl1 = pl1 & pl2
				pl2 = ""
				cl2 = 1

			End If
		End If
	Loop

	BinaryToString = pl1 & pl2 & pl3

End Function

' -------------------------------------------------------
Private Function MD5Round(strRound, a, b, C, d, X, S, ac)

	Select Case strRound

		Case "FF"
			a = MD5LongAdd4(a, (b And C) Or (Not (b) And d), X, ac)
			a = MD5Rotate(a, S)
			a = MD5LongAdd(a, b)
        
		Case "GG"
			a = MD5LongAdd4(a, (b And d) Or (C And Not (d)), X, ac)
			a = MD5Rotate(a, S)
			a = MD5LongAdd(a, b)
            
		Case "HH"
			a = MD5LongAdd4(a, b Xor C Xor d, X, ac)
			a = MD5Rotate(a, S)
			a = MD5LongAdd(a, b)
            
		Case "II"
			a = MD5LongAdd4(a, C Xor (b Or Not (d)), X, ac)
			a = MD5Rotate(a, S)
			a = MD5LongAdd(a, b)
	End Select

End Function

' -------------------------------------------
Private Function MD5Rotate(lngValue, lngBits)

	Dim lngSign
	Dim lngI
    
	lngBits = (lngBits Mod 32)

	If lngBits = 0 Then MD5Rotate = lngValue: Exit Function

	For lngI = 1 To lngBits
		lngSign = lngValue And &HC0000000
		lngValue = (lngValue And &H3FFFFFFF) * 2
		lngValue = lngValue Or ((lngSign < 0) And 1) Or (CBool(lngSign And &H40000000) And &H80000000)
	Next

	MD5Rotate = lngValue

End Function

' ---------------------
Private Function TRID()

	Dim sngNum, lngnum
	Dim strResult

	sngNum = Rnd(2147483648)
	strResult = CStr(sngNum)

	strResult = Replace(strResult, "0.", "")
	strResult = Replace(strResult, ".", "")
	strResult = Replace(strResult, "E-", "")

	TRID = strResult

End Function

' -------------------------------------------------
Private Function MD564Split(lngLength, bytBuffer())

	Dim lngBytesTotal, lngBytesToAdd
	Dim intLoop, intLoop2, lngTrace
	Dim intInnerLoop, intLoop3

	lngBytesTotal = lngTrack Mod 64
	lngBytesToAdd = 64 - lngBytesTotal
	lngTrack = (lngTrack + lngLength)

	If lngLength >= lngBytesToAdd Then

		For intLoop = 0 To lngBytesToAdd - 1
			arrSplit64(lngBytesTotal + intLoop) = bytBuffer(intLoop)
		Next

		MD5Conversion arrSplit64

		lngTrace = (lngLength) Mod 64

		For intLoop2 = lngBytesToAdd To lngLength - intLoop - lngTrace Step 64

			For intInnerLoop = 0 To 63
				arrSplit64(intInnerLoop) = bytBuffer(intLoop2 + intInnerLoop)
			Next

		MD5Conversion arrSplit64

		Next

		lngBytesTotal = 0
	Else

		intLoop2 = 0

	End If

	For intLoop3 = 0 To lngLength - intLoop2 - 1

		arrSplit64(lngBytesTotal + intLoop3) = bytBuffer(intLoop2 + intLoop3)

	Next

End Function

' ---------------------------------------
Private Function MD5StringArray(strInput)

	Dim intLoop
	Dim bytBuffer()
	ReDim bytBuffer(Len(strInput))
    
	For intLoop = 0 To Len(strInput) - 1
		bytBuffer(intLoop) = Asc(Mid(strInput, intLoop + 1, 1))
	Next

	MD5StringArray = bytBuffer

End Function

' ------------------------------------
Private Sub MD5Conversion(bytBuffer())

	Dim X(16), a
	Dim b, C
	Dim d

	a = arrLongConversion(1)
	b = arrLongConversion(2)
	C = arrLongConversion(3)
	d = arrLongConversion(4)

	MD5Decode 64, X, bytBuffer

	MD5Round "FF", a, b, C, d, X(0), S11, -680876936
	MD5Round "FF", d, a, b, C, X(1), S12, -389564586
	MD5Round "FF", C, d, a, b, X(2), S13, 606105819
	MD5Round "FF", b, C, d, a, X(3), S14, -1044525330
	MD5Round "FF", a, b, C, d, X(4), S11, -176418897
	MD5Round "FF", d, a, b, C, X(5), S12, 1200080426
	MD5Round "FF", C, d, a, b, X(6), S13, -1473231341
	MD5Round "FF", b, C, d, a, X(7), S14, -45705983
	MD5Round "FF", a, b, C, d, X(8), S11, 1770035416
	MD5Round "FF", d, a, b, C, X(9), S12, -1958414417
	MD5Round "FF", C, d, a, b, X(10), S13, -42063
	MD5Round "FF", b, C, d, a, X(11), S14, -1990404162
	MD5Round "FF", a, b, C, d, X(12), S11, 1804603682
	MD5Round "FF", d, a, b, C, X(13), S12, -40341101
	MD5Round "FF", C, d, a, b, X(14), S13, -1502002290
	MD5Round "FF", b, C, d, a, X(15), S14, 1236535329

	MD5Round "GG", a, b, C, d, X(1), S21, -165796510
	MD5Round "GG", d, a, b, C, X(6), S22, -1069501632
	MD5Round "GG", C, d, a, b, X(11), S23, 643717713
	MD5Round "GG", b, C, d, a, X(0), S24, -373897302
	MD5Round "GG", a, b, C, d, X(5), S21, -701558691
	MD5Round "GG", d, a, b, C, X(10), S22, 38016083
	MD5Round "GG", C, d, a, b, X(15), S23, -660478335
	MD5Round "GG", b, C, d, a, X(4), S24, -405537848
	MD5Round "GG", a, b, C, d, X(9), S21, 568446438
	MD5Round "GG", d, a, b, C, X(14), S22, -1019803690
	MD5Round "GG", C, d, a, b, X(3), S23, -187363961
	MD5Round "GG", b, C, d, a, X(8), S24, 1163531501
	MD5Round "GG", a, b, C, d, X(13), S21, -1444681467
	MD5Round "GG", d, a, b, C, X(2), S22, -51403784
	MD5Round "GG", C, d, a, b, X(7), S23, 1735328473
	MD5Round "GG", b, C, d, a, X(12), S24, -1926607734

	MD5Round "HH", a, b, C, d, X(5), S31, -378558
	MD5Round "HH", d, a, b, C, X(8), S32, -2022574463
	MD5Round "HH", C, d, a, b, X(11), S33, 1839030562
	MD5Round "HH", b, C, d, a, X(14), S34, -35309556
	MD5Round "HH", a, b, C, d, X(1), S31, -1530992060
	MD5Round "HH", d, a, b, C, X(4), S32, 1272893353
	MD5Round "HH", C, d, a, b, X(7), S33, -155497632
	MD5Round "HH", b, C, d, a, X(10), S34, -1094730640
	MD5Round "HH", a, b, C, d, X(13), S31, 681279174
	MD5Round "HH", d, a, b, C, X(0), S32, -358537222
	MD5Round "HH", C, d, a, b, X(3), S33, -722521979
	MD5Round "HH", b, C, d, a, X(6), S34, 76029189
	MD5Round "HH", a, b, C, d, X(9), S31, -640364487
	MD5Round "HH", d, a, b, C, X(12), S32, -421815835
	MD5Round "HH", C, d, a, b, X(15), S33, 530742520
	MD5Round "HH", b, C, d, a, X(2), S34, -995338651

	MD5Round "II", a, b, C, d, X(0), S41, -198630844
	MD5Round "II", d, a, b, C, X(7), S42, 1126891415
	MD5Round "II", C, d, a, b, X(14), S43, -1416354905
	MD5Round "II", b, C, d, a, X(5), S44, -57434055
	MD5Round "II", a, b, C, d, X(12), S41, 1700485571
	MD5Round "II", d, a, b, C, X(3), S42, -1894986606
	MD5Round "II", C, d, a, b, X(10), S43, -1051523
	MD5Round "II", b, C, d, a, X(1), S44, -2054922799
	MD5Round "II", a, b, C, d, X(8), S41, 1873313359
	MD5Round "II", d, a, b, C, X(15), S42, -30611744
	MD5Round "II", C, d, a, b, X(6), S43, -1560198380
	MD5Round "II", b, C, d, a, X(13), S44, 1309151649
	MD5Round "II", a, b, C, d, X(4), S41, -145523070
	MD5Round "II", d, a, b, C, X(11), S42, -1120210379
	MD5Round "II", C, d, a, b, X(2), S43, 718787259
	MD5Round "II", b, C, d, a, X(9), S44, -343485551

	arrLongConversion(1) = MD5LongAdd(arrLongConversion(1), a)
	arrLongConversion(2) = MD5LongAdd(arrLongConversion(2), b)
	arrLongConversion(3) = MD5LongAdd(arrLongConversion(3), C)
	arrLongConversion(4) = MD5LongAdd(arrLongConversion(4), d)

End Sub

' -------------------------------------------
Private Function MD5LongAdd(lngVal1, lngVal2)

	Dim lngHighWord
	Dim lngLowWord
	Dim lngOverflow

	lngLowWord = (lngVal1 And &HFFFF&) + (lngVal2 And &HFFFF&)
	lngOverflow = lngLowWord \ 65536
	lngHighWord = (((lngVal1 And &HFFFF0000) \ 65536) + ((lngVal2 And &HFFFF0000) \ 65536) + lngOverflow) And &HFFFF&

	MD5LongAdd = MD5LongConversion((lngHighWord * 65536) + (lngLowWord And &HFFFF&))

End Function

' --------------------------------------------------------------
Private Function MD5LongAdd4(lngVal1, lngVal2, lngVal3, lngVal4)

	Dim lngHighWord
	Dim lngLowWord
	Dim lngOverflow

	lngLowWord = (lngVal1 And &HFFFF&) + (lngVal2 And &HFFFF&) + (lngVal3 And &HFFFF&) + (lngVal4 And &HFFFF&)
	lngOverflow = lngLowWord \ 65536
	lngHighWord = (((lngVal1 And &HFFFF0000) \ 65536) + ((lngVal2 And &HFFFF0000) \ 65536) + ((lngVal3 And &HFFFF0000) \ 65536) + ((lngVal4 And &HFFFF0000) \ 65536) + lngOverflow) And &HFFFF&
	MD5LongAdd4 = MD5LongConversion((lngHighWord * 65536) + (lngLowWord And &HFFFF&))

End Function

' -------------------------------------------------------------
Private Sub MD5Decode(intLength, lngOutBuffer(), bytInBuffer())

	Dim intDblIndex
	Dim intByteIndex
	Dim dblSum

	intDblIndex = 0
    
	For intByteIndex = 0 To intLength - 1 Step 4

		dblSum = bytInBuffer(intByteIndex) + bytInBuffer(intByteIndex + 1) * 256 + bytInBuffer(intByteIndex + 2) * 65536 + bytInBuffer(intByteIndex + 3) * 16777216
		lngOutBuffer(intDblIndex) = MD5LongConversion(dblSum)
		intDblIndex = (intDblIndex + 1)

	Next

End Sub

' ------------------------------------------
Private Function MD5LongConversion(dblValue)

	If dblValue < 0 Or dblValue >= OFFSET_4 Then Error 6

	If dblValue <= MAXINT_4 Then
		MD5LongConversion = dblValue
	Else
		MD5LongConversion = dblValue - OFFSET_4
	End If

End Function

' ---------------------
Private Sub MD5Finish()

	Dim dblBits
	Dim arrPadding(72)
	Dim lngBytesBuffered

	arrPadding(0) = &H80

	dblBits = lngTrack * 8

	lngBytesBuffered = lngTrack Mod 64
    
	If lngBytesBuffered <= 56 Then
		MD564Split (56 - lngBytesBuffered), arrPadding
	Else
		MD564Split (120 - lngTrack), arrPadding
	End If

	arrPadding(0) = MD5LongConversion(dblBits) And &HFF&
	arrPadding(1) = MD5LongConversion(dblBits) \ 256 And &HFF&
	arrPadding(2) = MD5LongConversion(dblBits) \ 65536 And &HFF&
	arrPadding(3) = MD5LongConversion(dblBits) \ 16777216 And &HFF&
	arrPadding(4) = 0
	arrPadding(5) = 0
	arrPadding(6) = 0
	arrPadding(7) = 0

	MD564Split 8, arrPadding

End Sub

' --------------------------------------
Private Function MD5StringChange(lngnum)

	Dim bytA
	Dim bytB
	Dim bytC
	Dim bytD

	bytA = lngnum And &HFF&
	If bytA < 16 Then
		MD5StringChange = "0" & Hex(bytA)
	Else
		MD5StringChange = Hex(bytA)
	End If

	bytB = (lngnum And &HFF00&) \ 256
	If bytB < 16 Then
		MD5StringChange = MD5StringChange & "0" & Hex(bytB)
	Else
		MD5StringChange = MD5StringChange & Hex(bytB)
	End If

	bytC = (lngnum And &HFF0000) \ 65536
	If bytC < 16 Then
		MD5StringChange = MD5StringChange & "0" & Hex(bytC)
	Else
		MD5StringChange = MD5StringChange & Hex(bytC)
	End If

	If lngnum < 0 Then
		bytD = ((lngnum And &H7F000000) \ 16777216) Or &H80&
	Else
		bytD = (lngnum And &HFF000000) \ 16777216
	End If

	If bytD < 16 Then
		MD5StringChange = MD5StringChange & "0" & Hex(bytD)
	Else
		MD5StringChange = MD5StringChange & Hex(bytD)
	End If

End Function

' -------------------------
Private Function MD5Value()

	MD5Value = LCase(MD5StringChange(arrLongConversion(1)) & MD5StringChange(arrLongConversion(2)) & MD5StringChange(arrLongConversion(3)) & MD5StringChange(arrLongConversion(4)))

End Function

' ---------------------------------------------------
Public Function CalculateMD5(strMessage)

	Dim bytBuffer

	bytBuffer = MD5StringArray(strMessage)

	MD5Start

		MD564Split Len(strMessage), bytBuffer

	MD5Finish

	CalculateMD5 = MD5Value

End Function

' --------------------
Private Sub MD5Start()

	lngTrack = 0
	arrLongConversion(1) = MD5LongConversion(1732584193)
	arrLongConversion(2) = MD5LongConversion(4023233417)
	arrLongConversion(3) = MD5LongConversion(2562383102)
	arrLongConversion(4) = MD5LongConversion(271733878)

End Sub

