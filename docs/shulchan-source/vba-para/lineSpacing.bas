VBA MACRO lineSpacing.bas 
in file: word/vbaProject.bin - OLE stream: 'VBA/LineSpacing'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit
Option Private Module

Sub Repair(ByVal paraCollection As Collection)

    Dim selectionRange As Range
    Dim paraRange As Range
    Dim lineSpacing As Double
    Dim fontName As String
    Dim fontSize As Single
    Dim savedSpaceBefore As Double
    Dim i As Integer
    
    ' Begin undo recording
    On Error Resume Next
    Application.UndoRecord.StartCustomRecord "תיקון מרווח שורות"
    On Error GoTo Ending
    On Error GoTo 0
    
    FormRunning.LblLineSpacing.Caption = "אוסף מידע עבור תיקון מרווח שורות"
    FormRunning.LblLineSpacing.Font.Bold = True: DoEvents

    ' Set selection range
    Set selectionRange = Selection.Range
    Set paraCollection = Helpers.GetValidParagraphRangesForLineSpacing(paraCollection, False)
    
    ' Loop through paragraphs
    For i = 1 To paraCollection.Count
        Set paraRange = paraCollection(i).Duplicate
        
        FormRunning.LblLineSpacing.Caption = "מתקן מרווח שורות " & i & " מתוך " & paraCollection.Count
        If i Mod -Int(-paraCollection.Count / 40) = 0 Then DoEvents
        If stopCode Then GoTo Ending
        
        With paraRange
            If Not .ParagraphFormat.LineSpacingRule = wdLineSpaceExactly Then
                savedSpaceBefore = .Paragraphs(1).SpaceBefore
                .Paragraphs(1).SpaceBefore = 0
                GetFontProperties paraRange, fontName, fontSize
                .Collapse wdCollapseStart
                .InsertBefore Chr(11)
                With .Font
                    .NameBi = fontName
                    .Name = fontName
                    .SizeBi = fontSize
                    .Size = fontSize
                End With
                .Collapse wdCollapseEnd
'                Application.ScreenRefresh
                lineSpacing = .Information(wdVerticalPositionRelativeToTextBoundary)
                If lineSpacing > 0 Then
                    .ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
                    .ParagraphFormat.lineSpacing = lineSpacing
                End If
                .MoveStart wdCharacter, -1
                .Delete
                .Paragraphs(1).SpaceBefore = savedSpaceBefore
            End If
        End With
    
    Next i

    FormRunning.LblLineSpacing.Caption = "תיקון מרווח שורות - הושלם": DoEvents
    ' Restore original selection
    selectionRange.Select
    
Ending:
    ' End undo recording
    Application.UndoRecord.EndCustomRecord
End Sub

Sub remove(ByVal paraCollection As Collection)

    Dim selectionRange As Range
    Dim paraRange As Range
    Dim lineSpacingExactly As Double
    Dim lineSpacingSingle As Double
    Dim lineSpacing As Double
    Dim fontName As String
    Dim fontSize As Single
    Dim i As Integer
    
    ' Begin undo recording
    On Error Resume Next
    Application.UndoRecord.StartCustomRecord "ביטול תיקון מרווח שורות"
    On Error GoTo Ending
    
    FormRemoveRunning.LblLineSpacing.Font.Bold = True
    FormRemoveRunning.LblLineSpacing.Caption = "אוסף מידע עבור הסרת תיקון מרווח שורות"
    Set paraCollection = Helpers.GetValidParagraphRangesForLineSpacing(paraCollection, True)
    'set range
    Set selectionRange = Selection.Range
    
    ' Loop through paragraphs
    For i = 1 To paraCollection.Count
        FormRemoveRunning.LblLineSpacing.Caption = "מסיר תיקון מרווח שורות " & i & " מתוך " & paraCollection.Count
        If i Mod -Int(-paraCollection.Count / 40) = 0 Then DoEvents
        If stopCode Then GoTo Ending
        
        Set paraRange = paraCollection(i).Duplicate
        
        FormRemoveRunning.LblLineSpacing.Caption = "מסיר תיקון מרווח שורות " & i & " מתוך " & paraCollection.Count: DoEvents
        If stopCode Then GoTo Ending
            
        With paraRange
            If .ParagraphFormat.LineSpacingRule = wdLineSpaceExactly Then
                GetFontProperties paraRange, fontName, fontSize
                lineSpacingExactly = .ParagraphFormat.lineSpacing
                .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
                .Collapse wdCollapseStart
                .text = Chr(11)
                With .Font
                    .NameBi = fontName
                    .Name = fontName
                    .SizeBi = fontSize
                    .Size = fontSize
                End With
                .Collapse wdCollapseStart
                .MoveEnd wdCharacter, 1
                .Collapse wdCollapseEnd
                lineSpacingSingle = .Information(wdVerticalPositionRelativeToTextBoundary) - .Paragraphs(1).SpaceBefore
                .MoveStart wdCharacter, -1
                .Delete
                lineSpacing = lineSpacingExactly / lineSpacingSingle
                If lineSpacing > 0 Then
                    .ParagraphFormat.LineSpacingRule = wdLineSpaceMultiple
                    .ParagraphFormat.lineSpacing = LinesToPoints(lineSpacing)
                End If
            End If
        End With
        
    Next i
    
    FormRemoveRunning.LblLineSpacing.Caption = "הסרת תיקון מרווח שורות - הושלם"
    'reset selection
    selectionRange.Select
    
Ending:
    ' End undo recording
    Application.UndoRecord.EndCustomRecord
End Sub

Sub alignmentUpwards()
    
    Dim paraCollection As Collection
    Dim selectionRange As Range
    Dim paraRange As Range
    Dim firstWordStyle As style

    ' Begin undo recording
    On Error Resume Next
    Application.UndoRecord.StartCustomRecord "יישור כלפי מעלה"
    On Error GoTo Ending
    
    ' Set selection range
    Set selectionRange = Selection.Range
    Set paraCollection = Helpers.GetValidParagraphRangesForFirstWord(selectionRange)
    
    ' Loop through paragraphs
    For Each paraRange In paraCollection
            
        With paraRange
            .Collapse
            .MoveEndUntil Cset:=" " & Chr(13)
            .style = newPosition(paraRange)
        End With
        
    Next paraRange
    
    ' Restore original selection
    selectionRange.Select
        
Ending:
    ' End undo recording
    Application.UndoRecord.EndCustomRecord
End Sub
-------------------------------------------------------------------------------
