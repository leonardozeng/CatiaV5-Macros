Public TimeOnOFF As Boolean
Sub Autosave()

Dim deadOnTime As Date
deadOnTime = "14:38:01"

TimeOnOFF = Not TimeOnOFF

If TimeOnOFF Then

    Dim S As Integer
        While TimeOnOFF = True

            If Second(Now) > S Or Second(Now) = 0 Then
    
                
                frmClock.Label1.Caption = Time()
                
                S = Second(Now)
                If Time = deadOnTime Then
                
'                    frmClock.Label2.Caption = "Save..."
                    CATIA.ActiveDocument.Save
                    MsgBox CATIA.ActiveDocument.Name & " was saved."
                    
                    Exit Sub
                    
                    
'
'                Else
                
'                    frmClock.Label2.Caption = deadOnTime - Time()
                
                End If
                
            End If
            
            DoEvents
        Wend
End If

End Sub
