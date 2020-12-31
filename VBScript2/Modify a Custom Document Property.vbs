' Description: Modifies a custom property (TestProperty, setting the new value to "New value") found in the summary information properties for a document named C:\Scripts\Test.doc.


Set objPropertyReader = CreateObject("DSOleFile.PropertyReader")
Set objDocument = objPropertyReader.GetDocumentProperties _
    ("C:\Scripts\Test.doc")

Set colCustomProperties = objDocument.CustomProperties
For Each strProperty in colCustomProperties
    If strProperty.Name = "TestProperty" Then
        strProperty.Value = "New value"
    End If
Next

