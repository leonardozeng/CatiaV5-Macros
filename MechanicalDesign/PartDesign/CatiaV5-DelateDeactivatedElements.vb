Sub CATMain()
On Error Resume Next

Dim partDocument As Document
Set partDocument = CATIA.ActiveDocument

Dim myPart As Part
Set myPart = partDocument1.Part

'If Err.Number = 0 Then

    Dim selection1 As Selection
    Set selection1 = partDocument.Selection
    selection1.Search "CATPrtSearch.PartDesign Feature.Activity=FALSE"
    
    If selection1.Count = 0 Then
        MsgBox "Nie ma deaktywowanych elementów"
        Exit Sub
        
    Else
        MsgBox ("Liczba deaktywowanych elementów to: " & selection1.Count & ". Kliknij Tak aby potwierdzić usuwanie lub Nie aby wyjść.")
        selection1.Delete
        part1.Update
    End If

    
'Else
'MsgBox "Otwary dokument nie jest dokumentem typu PartDesign!"
'End If
End Sub

