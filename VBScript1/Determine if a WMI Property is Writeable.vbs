strComputer = "."
strNamespace = "\root\cimv2"
strClass = "Win32_Printer"
strProperty = "PortName"

blnWriteable = False

Set objClass = GetObject("winmgmts:\\" & strComputer & strNameSpace & ":" & strClass)

For Each objClassProperty In objClass.Properties_
    If objClassProperty.Name = strProperty Then
        If objClassProperty.IsArray Then
            Wscript.Echo "This property is an array."
        End If

        Select Case objClassProperty.CIMType
            Case 8
                Wscript.Echo "This is a string property."
            Case 11
                Wscript.Echo "This is a Boolean property."
            Case 13
                Wscript.Echo "This is an Object property."
            Case 101
                Wscript.Echo "This is a date-time property."
            Case 102
            Wscript.Echo "This is a Reference property."
            Case 103
                Wscript.Echo "This is a string-type property."
            Case Else
                Wscript.Echo "This is a numeric property."
        End Select

        For Each objQualifier in ObjClassProperty.Qualifiers_
            If objQualifier.Name = "write" Then
                blnWriteable = True
            End If
        Next
    End If
Next

If blnWriteable = True Then
    Wscript.Echo "This property is read-write."
Else
    Wscript.Echo "This property is read-only."
End If
  


