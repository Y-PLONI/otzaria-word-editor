VBA MACRO LineCommentBox.bas 
in file: word/vbaProject.bin - OLE stream: 'VBA/LineCommentBox'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit
Option Private Module
Sub Add()
    Dim doc As Document
    Dim rng As Range
    Dim para As Paragraph
    Dim userText As String
    Dim sizePercentage As Double
    Dim box() As Shape
    Dim prevLineSpacing As Double
    Dim prevLineSpacingRule As WdLineSpacing
    Dim lineHeight As Double
    Dim lineHeightSpaceSingle As Double
    Dim lineCount As Integer
    Dim posInPage As Double
    Dim i As Integer
    
    Set doc = ActiveDocument
    Set rng = Selection.Range
    userText = InputBox("הכנס את גודל הפונט באחוזים")
    If Not Helpers.TextToNumIsInRange(userText, 20, 100, sizePercentage, charsToRemove:="%", msg:=True) Then Exit Sub
    sizePercentage = sizePercentage / 100
    
    On Error GoTo Ending
    Application.UndoRecord.StartCustomRecord "הוסף פירוש מתחת המילים"
    Application.ScreenUpdating = False
    
    For Each para In rng.Paragraphs
        With para.Range
            prevLineSpacing = .ParagraphFormat.LineSpacing
            prevLineSpacingRule = .ParagraphFormat.LineSpacingRule
            .ParagraphFormat.LineSpacing = 12
            .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
            lineHeightSpaceSingle = Helpers.GetLineHeight(rng:=.Duplicate)
            lineCount = .ComputeStatistics(wdStatisticLines)
            .ParagraphFormat.LineSpacing = prevLineSpacing
            .ParagraphFormat.LineSpacingRule = prevLineSpacingRule
            lineHeight = Helpers.GetLineHeight(rng:=.Duplicate)
            .Select
            .MoveEnd wdCharacter, -1
            .InsertAfter Chr(9) & Chr(11)
        End With
        
        ReDim box(1 To lineCount) As Shape
        
        With Selection
            .Collapse wdCollapseStart
            For i = 1 To lineCount
                .EndKey unit:=wdLine
                .InsertAfter Chr(11) & "א"
                .Collapse wdCollapseEnd
                posInPage = .Information(wdVerticalPositionRelativeToPage)
                .MoveStart wdCharacter, -2
                .Delete
                .Expand wdLine
                Set box(i) = doc.Shapes.AddTextbox( _
                    Orientation:=msoTextOrientationHorizontal, _
                    Left:=CentimetersToPoints(0), _
                    Top:=posInPage, _
                    Width:=Helpers.GetPageWidth(rng), _
                    Height:=lineHeight * sizePercentage, _
                    Anchor:=.Range)
                With box(i)
                    .name = GetBoxName
                    .WrapFormat.Type = wdWrapSquare
                    .RelativeVerticalPosition = wdRelativeVerticalPositionLine
                    .RelativeHorizontalPosition = wdRelativeHorizontalPositionMargin
                    .Left = 0
                    .Top = (-lineHeight + lineHeightSpaceSingle) ' * sizePercentage
                    .line.Visible = msoFalse
                    .TextFrame.MarginBottom = (lineHeight - lineHeightSpaceSingle) * sizePercentage
                    .TextFrame.MarginTop = 0
                    .TextFrame.MarginLeft = 0
                    .TextFrame.MarginRight = 0
                    If i = 1 Then
                        Call Helpers.CopyFontFormat(para.Range, .TextFrame.TextRange)
                        With .TextFrame.TextRange.Font
                            .Size = .Size * sizePercentage
                            .SizeBi = .SizeBi * sizePercentage
                        End With
                    Else
                        On Error Resume Next
                        box(i - 1).TextFrame.Next = .TextFrame
                        On Error GoTo Ending
                    End If
                    
                    If i = lineCount Then
                        .WrapFormat.Type = wdWrapFront
                    .Top = (-lineHeight + lineHeightSpaceSingle) + lineHeight
                    End If
                End With
            Next i
        End With
    Next para
Ending:
    Application.UndoRecord.EndCustomRecord
    Application.ScreenUpdating = True
End Sub
Sub Remove()
    Dim rng As Range
    Dim shp As Shape
    Dim i As Integer
    
    On Error GoTo Ending
    Application.UndoRecord.StartCustomRecord "הסר פירוש מתחת המילים"
    
    Set rng = Selection.Range
    rng.Expand wdParagraph
    For i = rng.ShapeRange.Count To 1 Step -1
        Set shp = rng.ShapeRange(i)
        If shp.name = GetBoxName Then shp.Delete
    Next i
Ending:
    Application.UndoRecord.EndCustomRecord
End Sub
Sub ShowBoxLine(showOption As Boolean)
    Dim rng As Range
    Dim shp As Shape
    
    On Error GoTo Ending
    Application.UndoRecord.StartCustomRecord "הצג מיתאר של תיבות הטקסט"
    
    Set rng = Selection.Range
    rng.Expand wdParagraph
    For Each shp In rng.ShapeRange
        If shp.name = GetBoxName Then
            shp.line.Visible = showOption
            shp.line.ForeColor = RGB(255, 0, 255)
        End If
    Next shp
Ending:
    Application.UndoRecord.EndCustomRecord
End Sub
Function GetBoxName() As String
    GetBoxName = "LineCommentBox"
End Function
-------------------------------------------------------------------------------
