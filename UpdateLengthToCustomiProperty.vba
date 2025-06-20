Sub UpdateLengthToCustomiProperty()
    Dim invApp As Inventor.Application
    Set invApp = ThisApplication
    
    Dim doc As Document
    Dim partDoc As PartDocument
    Dim userParams As Parameters
    Dim param As Parameter
    Dim lengthValue As String
    Dim propSet As PropertySet
    Dim customProp As Property
    Dim foundLength As Boolean
    Dim i As Integer

    ' Loop through all open documents
    For i = 1 To invApp.Documents.Count
        If invApp.Documents.Item(i).DocumentType = kPartDocumentObject Then
            Set partDoc = invApp.Documents.Item(i)
            
            ' Access user parameters
            Set userParams = partDoc.ComponentDefinition.Parameters
            
            foundLength = False
            For Each param In userParams
                If LCase(param.Name) = "length" Then
                    lengthValue = CStr(param.Value * 10) ' convert to mm if internal units are cm
                    foundLength = True
                    Exit For
                End If
            Next param
            
            If foundLength Then
                Set propSet = partDoc.PropertySets.Item("Inventor User Defined Properties")
                On Error Resume Next
                Set customProp = propSet.Item("Długość")
                If Err.Number <> 0 Then
                    Err.Clear
                    ' Create the property if it doesn't exist
                    Set customProp = propSet.Add(lengthValue, "Długość")
                Else
                    ' Update existing property
                    customProp.Value = lengthValue
                End If
                On Error GoTo 0
            Else
                MsgBox "Length parameter not found in document: " & partDoc.DisplayName
            End If
        End If
    Next i

    MsgBox "Custom property 'Długość' updated in all open .ipt files."
End Sub
