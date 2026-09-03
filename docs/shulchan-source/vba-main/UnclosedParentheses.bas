VBA MACRO UnclosedParentheses.bas 
in file: word/vbaProject.bin - OLE stream: 'VBA/UnclosedParentheses'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Dim rng As Range
Sub Search(entireDocument As Boolean)
    
    Dim doc As Document
    Dim paraCount As Integer
    Dim paraRange As Range
    Dim opens As String
    Dim msg As String
    Dim i As Integer
    
    Set doc = ActiveDocument
    If rng Is Nothing Then Set rng = doc.Range
    If rng.End - rng.Start = 0 Then Set rng = doc.Range
    paraCount = rng.Paragraphs.Count
    
    FormUnclosedParentheses.Label1.Caption = "סורק את המסמך, אנא המתן..."
    
    For i = 1 To paraCount
        DoEvents
        If entireDocument Then
            Set paraRange = rng
            i = paraCount
        Else
            Set paraRange = rng.Paragraphs(i).Range
        End If
        
        With paraRange
        
            If rng.Start > .Start Then
                .Start = rng.Start
                msg = " - ייתכן ששגיאה זו קשורה לשגיאה הקודמת"
            Else
                msg = ""
            End If
            opens = ""
            .Collapse wdCollapseStart
            Do
                DoEvents
                
                If Len(opens) = 0 Then
                    .MoveUntil ")(][" & Chr(13)
                Else
                    .MoveEndUntil ")(][" & Chr(13)
                End If
                
                .MoveEnd wdCharacter, 1
                
                If Right(.text, 1) = Chr(13) And Len(opens) = 0 Then
                    Exit Do
                ElseIf Right(.text, 1) = Chr(13) And Len(opens) > 0 Then
                    .Select
                    FormUnclosedParentheses.Label1.Caption = "פותח ללא סוגר" & msg
                    GoTo Ending
                ElseIf Right(.text, 1) = "(" Or Right(.text, 1) = "[" Then
                    opens = opens & Right(.text, 1)
                ElseIf Right(.text, 1) = ")" Then
                    If Right(opens, 1) = "(" Then
                        opens = Left(opens, Len(opens) - 1)
                    ElseIf Right(opens, 1) = "[" Then
                        .Select
                        FormUnclosedParentheses.Label1.Caption = "סוגר לא תואם" & msg
                        GoTo Ending
                    Else
                        .Start = .End - 1
                        .Select
                        FormUnclosedParentheses.Label1.Caption = "סוגר ללא פותח" & msg
                        GoTo Ending
                    End If
                ElseIf Right(.text, 1) = "]" Then
                    If Right(opens, 1) = "[" Then
                        opens = Left(opens, Len(opens) - 1)
                    ElseIf Right(opens, 1) = "(" Then
                        .Select
                        FormUnclosedParentheses.Label1.Caption = "סוגר לא תואם" & msg
                        GoTo Ending
                    Else
                        .Start = .End - 1
                        .Select
                        FormUnclosedParentheses.Label1.Caption = "סוגר ללא פותח" & msg
                        GoTo Ending
                    End If
                End If
            Loop
            
        End With
    Next i
    
    FormUnclosedParentheses.Label1.Caption = "הסריקה הושלמה"
    FormUnclosedParentheses.CBNext.Caption = "סיים"
    FormUnclosedParentheses.CBNext.Visible = True
    Set rng = Nothing
    Exit Sub
Ending:
    FormUnclosedParentheses.CBNext.Visible = True
    rng.Start = paraRange.End

End Sub
Sub ResetRange()
    Set rng = Nothing
End Sub
-------------------------------------------------------------------------------
