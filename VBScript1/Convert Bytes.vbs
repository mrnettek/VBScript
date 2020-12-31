' This functions converts bytes to the specific type

' inputs -----------------------------
' intRaw: 	the raw Value
' strStype: 	Type of the raw value
' strDType: 	Type of the output
' intprec: 	Decimal percision
' Returns: 	numceric value

' Type Definition --------------------
'      bits - Bits
'         b - Byte
'         k - Kilobyte
'         m - megabyte
'         g - gigabyte
'         t - terabyte
'         p - petabyte
'         e - exobyte


wscript.echo ConvertBytes(1024,"k","m",0)

' ----------------------------------------------------------------
Private Function ConvertBytes(intRaw, strSType, strDtype, intPrec)

	Const BITS_PER_BYTE		= 8
	Const BYTES_PER_KILOBYTE	= 1024
	Const BYTES_PER_MEGABYTE	= 1048576
	Const BYTES_PER_GIGABYTE	= 1073741824
	Const BYTES_PER_TERABYTE	= 1099511627776
	Const BYTES_PER_PETABYTE	= 1125899906842624
	Const BYTES_PER_EXABYTE		= 1152921504606846976

	strSType=LCase(strSType)
	strDType=LCase(strDType)

	Select Case strSType
		Case "bits"
			intRaw=intRaw/BITS_PER_BYTE
		Case "b"
			intRaw=intRaw
		Case "k"
			intRaw=intRaw*BYTES_PER_KILOBYTE
		Case "m"
			intRaw=intRaw*BYTES_PER_MEGABYTE
		Case "g"
			intRaw=intRaw*BYTES_PER_GIGABYTE
		Case "t"
			intRaw=intRaw*BYTES_PER_TERABYTE
		Case "p"
			intRaw=intRaw*BYTES_PER_PETABYTE
		Case "e"
			intRaw=intRaw*BYTES_PER_EXABYTE
	End Select

	Select Case strDType
		Case "bits"
			intRaw=intRaw*BITS_PER_BYTE
		Case "b"
			intRaw=intRaw
		Case "k"
			intRaw=intRaw/BYTES_PER_KILOBYTE
		Case "m"
			intRaw=intRaw/BYTES_PER_MEGABYTE
		Case "g"
			intRaw=intRaw/BYTES_PER_GIGABYTE
		Case "t"
			intRaw=intRaw/BYTES_PER_TERABYTE
		Case "p"
			intRaw=intRaw/BYTES_PER_PETABYTE
		Case "e"
			intRaw=intRaw/BYTES_PER_EXABYTE
	End Select

	ConvertBytes = int(intRaw * (10^intPrec))/(10^intPrec)
	
End Function

