strFirstName = "kEn"
strLastName = "MYEr"

intFirstName = Len(strFirstName)
strFirstLetter = UCase(Left(strFirstName, 1))
strRemainingLetters = LCase(Right(strFirstName, intFirstName - 1))

strFirstName = strFirstLetter & strRemainingLetters

intLastName = Len(strLastName)
strFirstLetter = UCase(Left(strLastName, 1))
strRemainingLetters = LCase(Right(strLastName, intLastName - 1))

strLastName = strFirstLetter & strRemainingLetters

Wscript.Echo strFirstName, strLastName
  


