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
' Macro:    CatiaV5-SaveItWhenYouWant.catvbs
' Version:  0.1
' Code:     Catia VBS
' Purpose:  Put the time when you want to save your file 
' Autor:    Krzysztof Górzyński
' Datum:    25/03/2015
'----------------------------------------------------------------------------

Public TimeOnOFF As Boolean
Sub CATMain()
On Error Resume Next

Dim myTime As Date
myTime = "14:38:01"

TimeOnOFF = Not TimeOnOFF

If TimeOnOFF Then

    Dim S As Integer
        While TimeOnOFF = True

            If Second(Now) > S Or Second(Now) = 0 Then
    
                S = Second(Now)
                If Time = myTime Then
                
                    CATIA.ActiveDocument.Save
                    MsgBox CATIA.ActiveDocument.Name & " was saved."
                    
                    Exit Sub
                    
                End If
                
            End If
            DoEvents
        Wend
End If
End Sub
