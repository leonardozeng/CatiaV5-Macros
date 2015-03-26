'Copyright (c) 2015 Krzysztof Gorzynski <gorzynskikrzysztof@gmail.com>
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
' Macro: CatiaV5-TextToUpper.catvbs
' Version: 0.0
' Code: Catia VBS
' Purpose: 
' Autor: Krzysztof Górzyński
' Datum: 26/03/2015
'----------------------------------------------------------------------------

Option Explicit

Sub CATMain()
On Error Resume Next

Set drawingDocument1 = CATIA.ActiveDocument
Set drawingSheets1 = drawingDocument1.Sheets

If TypeName(drawingDocument1) = "DrawingDocument" Then
    For l = 1 To drawingSheets1.Count
        Set mySheet = drawingSheets1.Item(l)
        For i = 1 To mySheet.Views.Count
            Set myView = mySheet.Views.Item(i)
            Set myTexts = myView.Texts
            For k = 1 To myTexts.Count
                Set myText = myTexts.Item(k)
                myText.Text = VBA.UCase(myText.Text)
            Next
        Next
    Next

End Sub
