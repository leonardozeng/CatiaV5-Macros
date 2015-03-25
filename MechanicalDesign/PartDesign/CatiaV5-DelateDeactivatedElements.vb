'Copyright (c) 2014-2015 Krzysztof Gorzynski <gorzynskikrzysztof@gmail.com>
'
'Permission to use, copy, modify, and distribute this software for any
'purpose with or without fee is hereby granted, provided that the above
'copyright notice and this permission notice appear in all copies.
'
'THE SOFTWARE IS PROVIDED "AS IS" AND THE AUTHOR DISCLAIMS ALL WARRANTIES
'WITH REGARD TO THIS SOFTWARE INCLUDING ALL IMPLIED WARRANTIES OF
'MERCHANTABILITY AND FITNESS. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR
'ANY SPECIAL, DIRECT, INDIRECT, OR CONSEQUENTIAL DAMAGES OR ANY DAMAGES
'WHATSOEVER RESULTING FROM LOSS OF USE, DATA OR PROFITS, WHETHER IN AN
'ACTION OF CONTRACT, NEGLIGENCE OR OTHER TORTIOUS ACTION, ARISING OUT OF
'OR IN CONNECTION WITH THE USE OR PERFORMANCE OF THIS SOFTWARE.
'----------------------------------------------------------------------------
' Macro:    CatiaV5-DelateDeactivatedElements.catvbs
' Version:  0.0
' Code:     Catia VBS
' Purpose:  
' Autor:    Krzysztof Górzyński
' Datum:    24/03/2015
'----------------------------------------------------------------------------
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

