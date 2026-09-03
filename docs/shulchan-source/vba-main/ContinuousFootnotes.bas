VBA MACRO ContinuousFootnotes.bas 
in file: word/vbaProject.bin - OLE stream: 'VBA/ContinuousFootnotes'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit
Option Private Module
Dim marginHegiht As Double
Function GetTextBoxName() As String
    GetTextBoxName = "ContinuousFootnotesBox"
End Function
Function GetBmkName() As String
    GetBmkName = "_NotsBmk"
End Function
Sub Repair()
    
    Dim doc As Document
    Dim rng As Range
    Dim origNumberFormat As WdNoteNumberStyle
    Dim userText As String
    Dim para As Paragraph
    Dim box As Shape
    Dim boxHeight As Double
    Dim boxRange As Range
    Dim i, x As Integer
    Dim pageRange As Range
    Dim note As Endnote
    Dim noteRange As Range
    Dim noteBmk() As Bookmark
    Dim fldNoteNum() As Field
    Dim fld() As Field
    Dim boxCanToLink As Boolean
    Dim linesCount As Integer
    
    userText = InputBox("הכנס את המרחק בין הטקסט להערות (מ''מ)" & vbNewLine & "כולל גובה החוצץ")
    If Not TextToNumIsInRange(userText, 1, 15, marginHegiht, True) Then Exit Sub
    marginHegiht = CentimetersToPoints(marginHegiht) / 10
    
    On Error GoTo Ending
    Application.UndoRecord.StartCustomRecord "החלת הערות ברצף"
    Application.ScreenUpdating = False
    ActiveWindow.ActivePane.View.ShowAll = False

    FormRunning.Show vbModeless
    
    Set doc = ActiveDocument
    Set rng = doc.Range
    ReDim noteBmk(doc.Footnotes.Count) As Bookmark
    ReDim fld(doc.Footnotes.Count) As Field
    ReDim fldNoteNum(doc.Footnotes.Count) As Field
    x = 1
    
    ' המרת הערות השוליים להערות סיום
    FormRunning.Label1.Caption = "ממיר הערות שולים להערות סיום": DoEvents
    With doc
        If .Footnotes.Count = 0 Then
            MsgBox "לא נמצאו הערות שולים במסמך", vbOKOnly, "שגיאה"
            GoTo Ending
        ElseIf .Endnotes.Count > 0 Then
            MsgBox "לא ניתן לבצע את הפעולה במסמך המכיל הערות סיום", vbOKOnly, "שגיאה"
            GoTo Ending
        End If
        origNumberFormat = .Footnotes.NumberStyle
        .Footnotes.Convert
        .Endnotes.NumberStyle = origNumberFormat
        .Range(.Content.End - 1, .Content.End).InsertBreak Type:=wdPageBreak
    End With
    
    ' מניעת בקרת שורות מיותמות
    FormRunning.Label1.Caption = "מונע בקרת שורות מיותמות": DoEvents
    For Each para In rng.Paragraphs
        para.Format.WidowControl = False
    Next para
    
    For i = 1 To doc.ComputeStatistics(wdStatisticPages) * 2
        
        Application.ScreenRefresh
        doc.Repaginate
        If i > rng.Information(wdActiveEndPageNumber) Then Exit For
        Selection.GoTo what:=wdGoToPage, Count:=i
        Set pageRange = Selection.Bookmarks("\Page").Range
        
        With pageRange
            If .Endnotes.Count > 0 Then
                FormRunning.Label1.Caption = "מוסיף תיבת טקסט": DoEvents
                Set box = CreateBox(pageRange)
            Else
                GoTo NextPage
            End If
        End With
        
        For x = x To doc.Endnotes.Count
            doc.Repaginate
            Application.ScreenRefresh
            FormRunning.Label1.Caption = "הערה " & x & " מתוך " & doc.Endnotes.Count
            DoEvents
            If stopCode Then GoTo Ending
            boxHeight = box.Height
            box.TextFrame.TextRange.Select
            linesCount = Selection.Range.ComputeStatistics(wdStatisticLines)
            
            Set note = doc.Endnotes(x)
            If note.Reference.Information(wdActiveEndPageNumber) > i Then Exit For
             
            Set noteRange = note.Range
            With noteRange
'                .Expand wdParagraph
'                .End = .End - 1
            End With
            Set noteBmk(note.Index) = doc.Bookmarks.Add(GetBmkName & note.Index, noteRange)
            box.TextFrame.TextRange.Characters.Last.Font.Superscript = True
            box.TextFrame.TextRange.Characters.Last.InsertCrossReference _
                ReferenceType:="הערת סיום", _
                ReferenceKind:=wdEndnoteNumber, _
                ReferenceItem:=note.Index, _
                InsertAsHyperlink:=False, _
                IncludePosition:=False, _
                SeparateNumbers:=False
            box.TextFrame.TextRange.InsertAfter Chr(160)
            
            Set boxRange = box.TextFrame.TextRange
            boxRange.Collapse wdCollapseEnd
            boxRange.MoveStart wdCharacter, -1
            boxRange.Font.Superscript = False
            Set fld(note.Index) = box.TextFrame.TextRange.Fields.Add( _
                Range:=box.TextFrame.TextRange.Characters.Last, _
                Type:=wdFieldRef, _
                text:=noteBmk(note.Index), _
                PreserveFormatting:=False)
            fld(note.Index).Update
            fld(note.Index).Result.Select
            Selection.ClearFormatting
            fld(note.Index).Result.Style = noteBmk(note.Index).Range.Style
            fld(note.Index).Update
            box.TextFrame.TextRange.InsertAfter " "
            box.TextFrame.TextRange.ParagraphFormat.Borders.DistanceFromTop = 8
            
            If box.name Like "*Linked" Then
                Call AutoBoxSizeManual(box)
            Else
                box.TextFrame.AutoSize = True
            End If
            
            If note.Reference.Information(wdActiveEndPageNumber) > i Then
                    box.TextFrame.AutoSize = False
                    box.Height = boxHeight
                    FormRunning.Label1.Caption = "מקשר תיבת טקסט לעמוד הבא": DoEvents
                    Call OrderBoxAndLinkToNext(box, note, linesCount)
                    x = x + 1
                    Exit For
            End If
        Next x

NextPage:
    Next i
    FormRunning.Label1.Caption = "מוסיף חוצץ להערות השולים"
    Call AddSeparator
Ending:
    
    Application.UndoRecord.EndCustomRecord
    Application.ScreenUpdating = True
    Unload FormRunning
    
End Sub
Sub Remove()
    Dim bmk As Bookmark
    Dim box As Shape
    Dim doc As Document
    Dim i As Integer
    
    On Error GoTo Ending
    Application.UndoRecord.StartCustomRecord "הסרת הערות ברצף"
    Application.ScreenUpdating = False
    
    Set doc = ActiveDocument
    
    For i = doc.Bookmarks.Count To 1 Step -1
        Set bmk = doc.Bookmarks(i)
        If bmk.name = GetBmkName Then
            bmk.Delete
        End If
    Next i
    
    For i = doc.Shapes.Count To 1 Step -1
        Set box = doc.Shapes(i)
        If box.name Like GetTextBoxName & "*" Then
            box.Delete
        End If
    Next i
    doc.Endnotes.Convert

Ending:
    
    Application.UndoRecord.EndCustomRecord
    Application.ScreenUpdating = True
End Sub

Function CreateBox(rng As Range, Optional pageNum As Integer = 0)
    
    Dim doc As Document
    Dim shp As Shape
    Dim box As Shape
    
    Set doc = ActiveDocument
    If pageNum = 0 Then
        pageNum = rng.Information(wdActiveEndPageNumber)
    Else
        Set rng = Selection.GoTo(what:=wdGoToPage, Count:=pageNum)
    End If
    
    For Each shp In doc.Shapes
        If shp.name Like GetTextBoxName & "*" And shp.Anchor.Information(wdActiveEndPageNumber) = pageNum Then
            Set box = shp
        ElseIf shp.Anchor.Information(wdActiveEndPageNumber) > pageNum Then
            Exit For
        End If
    Next shp
    
    If box Is Nothing Then
        Set box = doc.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=0, Top:=0, Width:=Helpers.GetPageWidth(rng), Height:=50)
        With box
            .name = GetTextBoxName
            .TextFrame.MarginBottom = 0
            .TextFrame.MarginTop = marginHegiht
            .TextFrame.MarginRight = 0
            .TextFrame.MarginLeft = 0
            .TextFrame.AutoSize = True
            .TextFrame.TextRange.ParagraphFormat.SpaceAfter = 0
            .line.Visible = msoFalse
            .RelativeHorizontalPosition = wdRelativeHorizontalPositionMargin
            .RelativeVerticalPosition = wdRelativeVerticalPositionMargin
            .Left = wdShapeRight
            .Top = wdShapeBottom
            .LayoutInCell = False
            .LockAnchor = True
            .WrapFormat.Type = wdWrapTopBottom
'            .TextFrame.TextRange.Columns.SetCount 2
        End With
    End If
    
    Set CreateBox = box

End Function
Function OrderBoxAndLinkToNext(box As Shape, note As Endnote, linesCount As Integer)
    Dim pageNum As Long
    Dim boxHeight As Double
    Dim boxMarginTop As Double
    Dim nextBox As Shape
    Dim fld As Field
    Dim i As Integer
    
    pageNum = note.Reference.Information(wdActiveEndPageNumber)
    i = linesCount
    Do While note.Reference.Information(wdActiveEndPageNumber) = pageNum
        i = i + 1
        boxHeight = box.Height
        Call AutoBoxSizeManual(box, i)
        If i >= 50 Then Exit Do
    Loop
    With box
        .Height = boxHeight
        .Height = .Height + 5
        .RelativeVerticalPosition = wdRelativeVerticalPositionBottomMarginArea
        .Top = -.Height + 5
        .TextFrame.TextRange.ParagraphFormat.SpaceAfter = 0
        .TextFrame.TextRange.Paragraphs(1).Format.WidowControl = False
        .name = GetTextBoxName & "Linked"
        
        Set nextBox = CreateBox(.Anchor, pageNum + 1)
        nextBox.name = nextBox.name & "Linked"
        .TextFrame.Next = nextBox.TextFrame
        Call AutoBoxSizeManual(nextBox)
    End With
End Function
Function ChekIfCanToLink(box As Shape, note As Endnote) As Boolean
    
    Dim orginalBoxHeight As Double
    Dim notePageNum As Integer
    Dim boxAutoSize As Boolean
    
    ' שמירת גובה התיבה בדיקת מיקום הערת שולים והגדלת התיבה
    orginalBoxHeight = box.Height
    boxAutoSize = box.TextFrame.AutoSize
    notePageNum = note.Reference.Information(wdActiveEndPageNumber)
    box.TextFrame.AutoSize = True
    box.Height = box.Height + 5
    
    ' בדיקה שהמיקום לא השתנה
    If notePageNum = note.Reference.Information(wdActiveEndPageNumber) Then
        ChekIfCanToLink = True
    End If
    
    ' החזרת גודל התיבה
    box.Height = orginalBoxHeight
    box.TextFrame.AutoSize = boxAutoSize
    
End Function
Function PrevCommentsInLine(note As Endnote)
    
    Dim noteIndex As Integer
    Dim prevNote As Endnote
    Dim lineStart, lineEnd As Long
    
    On Error GoTo Ending
    noteIndex = note.Index
    Set prevNote = ActiveDocument.Endnotes(noteIndex - 1)
    
    note.Reference.Select
    With Selection
        .Expand wdLine
        lineStart = .Start
        lineEnd = .End
    End With
    With prevNote.Reference
        If .Start >= lineStart And .End <= lineEnd Then
            PrevCommentsInLine = True
        End If
    End With
    
Ending:
    On Error GoTo 0

End Function
Function AutoBoxSizeManual(box As Shape, Optional lineCount As Integer = 0)
    
    box.TextFrame.TextRange.Select
    
    With Selection
        If lineCount > 0 Then
            .Collapse wdCollapseStart
            .Move wdLine, lineCount
        End If
        .InsertAfter Chr(11)
        .Collapse wdCollapseEnd
        Application.ScreenRefresh
        box.Height = .Information(wdVerticalPositionRelativeToTextBoundary) + _
            box.TextFrame.MarginBottom + _
            box.TextFrame.MarginTop
        .MoveStart wdCharacter, -1
        .Delete
    End With
    
    With box
        If .RelativeVerticalPosition = wdRelativeVerticalPositionBottomMarginArea Then
            .Top = -.Height + 5
        End If
    End With
    
End Function
Sub AddSeparator()
    Dim doc As Document
    Dim shp, box As Shape
    Dim i As Integer
    
    On Error GoTo Ending
    Application.UndoRecord.StartCustomRecord "הוספת חוצץ להערות שולים"
    Application.ScreenUpdating = False
    
    Set doc = ActiveDocument
    For i = doc.Shapes.Count To 1 Step -1
        Set shp = doc.Shapes(i)
        If shp.name Like GetTextBoxName & "*" Then
            If shp.name Like "*" & "Separator" Then
                shp.Delete
                GoTo NextShp
            End If
            Set box = doc.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=0, Top:=0, Width:=Helpers.GetPageWidth(doc.Range), Height:=20, Anchor:=shp.Anchor)
            With box
                .name = GetTextBoxName & "Separator"
                .Height = shp.TextFrame.MarginTop + 10
                .TextFrame.MarginBottom = 0
                .TextFrame.MarginTop = 0
                .TextFrame.MarginRight = 0
                .TextFrame.MarginLeft = 0
                .TextFrame.TextRange.ParagraphFormat.SpaceAfter = 0
                .line.Visible = msoFalse
                .RelativeHorizontalPosition = wdRelativeHorizontalPositionMargin
                .RelativeVerticalPosition = wdRelativeVerticalPositionMargin
                .Left = wdShapeRight
                .Top = Helpers.GetPageHeight(doc.Range) - shp.Height - 5
                If shp.RelativeVerticalPosition = wdRelativeVerticalPositionBottomMarginArea Then
                    .Top = Helpers.GetPageHeight(doc.Range) - Abs(shp.Top) - 5
                End If
                .LayoutInCell = False
                .LockAnchor = True
                .WrapFormat.Type = wdWrapNone
                .TextFrame.TextRange.text = "__________________"
            End With

        End If
NextShp:
    Next i
    box.name = GetTextBoxName & "FirstSeparator"
Ending:
    
    Application.UndoRecord.EndCustomRecord
    Application.ScreenUpdating = True
End Sub
Sub UpdateSeparator()
    
    Dim shp As Shape
    Dim firstShp As Shape
    
    On Error GoTo Ending
    Application.UndoRecord.StartCustomRecord "הוספת חוצץ להערות שולים"
    Application.ScreenUpdating = False
    
    For Each shp In ActiveDocument.Shapes
        If shp.name Like "*" & "FirstSeparator" Then
            Set firstShp = shp
            firstShp.TextFrame.TextRange.Select
            Selection.Copy
            Exit For
        End If
    Next shp
    
    For Each shp In ActiveDocument.Shapes
        shp.Select
        If shp.name Like "*" & "Separator" Then
            With shp
                .TextFrame.TextRange.Delete
                DoEvents
                .TextFrame.TextRange.Paste
                .Top = .Top + .Height - firstShp.Height
                .Height = firstShp.Height
                
            End With
        End If
NextShp:
    Next shp
Ending:
    
    Application.UndoRecord.EndCustomRecord
    Application.ScreenUpdating = True
End Sub
-------------------------------------------------------------------------------
