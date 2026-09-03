VBA MACRO Footnotes.bas 
in file: word/vbaProject.bin - OLE stream: 'VBA/Footnotes'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit
Option Private Module
Sub Add()
    
    Dim selectionText As String
    Dim SelectionRng As Range
    Dim FootnotesRng As Footnote
    Dim FirstChar As String
    
    On Error GoTo Ending
    Application.ScreenUpdating = False
    Application.UndoRecord.StartCustomRecord "הערת שולים ללא מספר"
    
    ' שמירת הטקסט ומחיקת התו האחרון במידה והוא רווח
    Set SelectionRng = Selection.Range
    selectionText = SelectionRng.text
    selectionText = RemoveChar(selectionText)
        
    ' טיפול בטקסט הראשי
    With SelectionRng
        If SelectionRng.text = "" Then .MoveEnd
        FirstChar = Left(.text, 1)
        .Collapse wdCollapseStart
        .MoveEnd 1
        .Delete
        Set FootnotesRng = ActiveDocument.Footnotes.Add(Range:=SelectionRng, Reference:=FirstChar)
        .MoveEnd 1
        .Font.Superscript = False
    End With
    
    ' טיפול בהערת השולים
    
    With FootnotesRng.Range
        .Expand wdParagraph
        .Delete
        .text = selectionText
        .Font.Bold = True
        .Font.BoldBi = True
        .Collapse wdCollapseEnd
        .Font.Bold = False
        .Font.BoldBi = False
        .Select
    End With
    
Ending:
    Application.ScreenUpdating = True
    Application.UndoRecord.EndCustomRecord
End Sub
Sub ReplaceSelection()
    
    Dim selectionText As String
    Dim SelectionRng As Range
    Dim ftn As Footnote
    
    On Error GoTo Ending
    Application.ScreenUpdating = False
    Application.UndoRecord.StartCustomRecord "מחיקת מספר מהערת שולים"
    
    Set SelectionRng = Selection.Range
    selectionText = SelectionRng.text
    selectionText = RemoveChar(selectionText)

    Set ftn = SelectionRng.Footnotes(1)
    With ftn
        If .Reference.Font.Hidden = False Then
            .Reference.Font.Hidden = True
            With .Range
                .Collapse
                .Start = .Paragraphs(1).Range.Start
                If .text Like ftn.Reference.text & "*" Then .Delete
                .InsertAfter selectionText
                .Bold = True
                .BoldBi = True
            End With
        End If
    End With
Ending:
    Application.ScreenUpdating = True
    Application.UndoRecord.EndCustomRecord
End Sub
Sub ReplaceAll()
    
    Dim ftn As Footnote
    Dim selectionText As String
    Dim wordsNum As Integer
    Dim text As String
    Dim i As Integer
    
    text = InputBox("הכנס את מספר המילים המשמשות לד''ה" & vbNewLine & "ניתן להכניס מספר שלילי כדי לקבוע את טווח המילים אחורה" & vbNewLine & "וכן ניתן להכניס 0")
    If IsNumeric(text) Then
        wordsNum = text
    Else
        MsgBox "המספר שהוכנס אינו חוקי"
        Exit Sub
    End If
    
    On Error GoTo Ending
    Application.ScreenUpdating = False
    Application.UndoRecord.StartCustomRecord "מחיקת מספר מכל הערות השולים"
        
    FormRunning.Show vbModeless
    For i = 1 To ActiveDocument.Footnotes.Count
        Set ftn = ActiveDocument.Footnotes(i)
        FormRunning.Label1.Caption = "מסתיר מספר הערת שולים " & i & " מתוך " & ActiveDocument.Footnotes.Count: DoEvents
        With ftn
            If .Reference.Font.Hidden = False And .Reference.Font.Superscript = True Then
                With .Reference
                    .Font.Hidden = True
                    If wordsNum < 0 Then
                        .MoveStart wdWord, wordsNum
                    Else
                        .MoveEnd wdWord, wordsNum
                    End If
                    selectionText = .text
                End With
                With .Range
                    .Collapse
                    .Select
                    .Start = .Paragraphs(1).Range.Start
                    If .text Like ftn.Reference.text & "*" Then .Delete
                    selectionText = RemoveChar(selectionText)
                    .InsertAfter selectionText
                    .Bold = True
                    .BoldBi = True
                End With
            End If
        End With
    Next i
Ending:
    Application.ScreenUpdating = True
    Application.UndoRecord.EndCustomRecord
    Unload FormRunning
End Sub
Function RemoveChar(text As String) As String
    
    Dim lastChar As String

    Do
        lastChar = Right(text, 1)
        If lastChar <> " " And lastChar <> Chr(9) And lastChar <> Chr(11) And lastChar <> Chr(13) Then Exit Do
        text = Left(text, Len(text) - 1)
    Loop
    Do
        lastChar = Left(text, 1)
        If lastChar <> " " And lastChar <> Chr(9) And lastChar <> Chr(11) And lastChar <> Chr(13) Then Exit Do
        text = Right(text, Len(text) - 1)
    Loop
    
    
    text = Replace(text, Chr(2), "")
    If text <> "" Then text = text & ". "
    RemoveChar = text

End Function
-------------------------------------------------------------------------------
