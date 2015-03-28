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
' Macro: CatiaV5-AddBomNumbersFromTxt.catvbs
' Version: 0.1
' Code: Catia VBS
' Purpose: 
' Autor: Krzysztof Górzyński
' Datum: 29/03/2015
'----------------------------------------------------------------------------
Sub CATMain()

Dim CATIA As Object
Set CATIA = GetObject(, "CATIA.Application")

If TypeName(CATIA.ActiveDocument) <> "DrawingDocument" Then
    MsgBox "Aktywny dokument nie jest dokumentem typu 2D!" & vbCrLf & _
    "Wykonaj ten skrypt na objekcie 2D durniu!", , msgboxtext2
    Exit Sub
End If

Dim MyDrawingDoc As DrawingDocument
Set MyDrawingDoc = CATIA.ActiveDocument

Dim MyDrawingSheets As DrawingSheets
Set MyDrawingSheets = MyDrawingDoc.Sheets

Dim strFilePath As String
Dim objFile As File
Dim objTextStream As TextStream
Dim strLine As String
Dim counter As Integer
Dim strTabel() As Variant 'lista czesci
Dim posTable() As Variant ' nr pozycji

counter = 0
strLine = "Default"

'otwiera okno dialogowe
strFilePath = CATIA.FileSelectionBox("Select Text File", "*.txt", 0)

'jesli klikniesz Anuluj
If strFilePath = "" Then Exit Sub

Set objFile = CATIA.FileSystem.getFile(strFilePath)
Set objTextStream = objFile.OpenAsTextStream("ForReading")

'AtEndOfStream - nie konczaca sie petla
Do While Not objTextStream.AtEndOfStream
    
    strLine = Replace(objTextStream.ReadLine, "-", "")
    
    'z powodu "AtEndOfStream - nie konczaca sie petla" warunkowe wyjscie z petli
    If strLine = "" Then Exit Do
    
    ReDim Preserve strTabel(counter)
    strTabel(counter) = CStr(strLine)
  
    counter = counter + 1
   
Loop

objTextStream.Close

'loop przez wszystkie arkusze
For numsheet = 1 To MyDrawingSheets.Count

    Dim CurrentSheet As DrawingSheet
    Set CurrentSheet = MyDrawingSheets.Item(numsheet)
    Dim DrwViews As DrawingViews
    Set DrwViews = CurrentSheet.Views
    
        'loop przez wszyskie widoki
        For numview = 1 To DrwViews.Count
            Dim DrwView As DrawingView
            Set DrwView = DrwViews.Item(numview)

            DrwView.Activate

            Dim DrwTexts     As DrawingTexts
            Set DrwTexts = DrwView.Texts
            Dim iTxt As Integer
            Dim nTxt As Integer
            Dim strText As String
            Dim dwgText As DrawingText
            Dim txtLeaders As DrawingLeaders
            ReDim posTable(counter) ' lista pozycji w formacie "XXX" np "005"
            
            iTxt = 1
            nTxt = 1
            
            'dopóki jest >= 1 text na widoku
            While iTxt <= DrwTexts.Count
                Set dwgText = DrwTexts.Item(iTxt)
                
                'interesuja nas tylko "Referenzkreis"
                If Left(dwgText.Name, 13) = "Referenzkreis" Then
                
                    'jesli nr jest np taki "4010040010_A"
                    For i = 1 To Len(dwgText.Text)
                        If Mid(dwgText.Text, i, 1) = "_" Then
                            Pos = i
                            dwgText.Text = Left(dwgText.Text, i - 1)
                        End If
                    Next i
                    
                    'jesli nr z 2D jest w liscie "strTabel" - ostatnie 10 znaków w rekordzie
                    'to z tego rekordu piersze 3 to "dwgText.Text"
                    For i = 1 To UBound(strTabel)
                        If dwgText.Text = Right(strTabel(i), 10) Then
                            dwgText.Text = Left(strTabel(i), 3)
                            posTable(i) = dwgText.Text
                        End If
                    Next i
                    
                    nTxt = nTxt + 1
                End If
                iTxt = iTxt + 1
                    
            Wend
        Next numview
Next numsheet

'Dim missingPos() As Variant ' lista brakujących pozycji w formacie "XXX XXXXXXXXXX" np "005 3014221510"

'For x = LBound(strTabel) To UBound(strTabel)
'    For i = LBound(posTable) To UBound(posTable)
'
'        If posTable(i) = Left(strTabel(x), 3) Then
'            MsgBox strTabel(x)
'            strTabel(x).Delete
'        End If
''        If posTable(i) <> Left(strTabel(x), 3) And i = UBound(posTable) Then
'
'    Next i
'Next x




'i = 1
'While strTabel(i) <> ""
'
'    iTxt = 1
'    While iTxt <= DrwTexts.Count
'
'        Set dwgText = DrwTexts.Item(iTxt)
'        If Left(dwgText.Name, 13) = "Referenzkreis" Then
'            For x = LBound(strTabel) To UBound(strTabel)
'
'                Dim posNr As String
'                posNr = Left(strTabel(x), 3)
'                If dwgText.Text = posNr Then
'                    Exit For
'
'                Else
'                    If posNr = "" Then
'
'                        MsgBox "Brak nr pozycji: " & posNr & dwgText.Text
'
'                    End If
'
'                End If
'
'            Next x
'
'        End If
'        iTxt = iTxt + 1
'    Wend
'
'    i = i + 1
'Wend
    

End Sub
