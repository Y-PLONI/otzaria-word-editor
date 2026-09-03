VBA MACRO HangingFirstWord.bas 
in file: word/vbaProject.bin - OLE stream: 'VBA/HangingFirstWord'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit
Option Private Module

Sub Execute(ByVal paraCollection As Collection, skip2Lines As Boolean)
    
    Dim selectionRange As Range
    Dim paraRange As Range
    Dim firstWordSpacing As Double
    Dim i As Integer
    
    ' Begin undo recording
    On Error Resume Next
    Application.UndoRecord.StartCustomRecord "עיצוב חלון"
    On Error GoTo Ending
    
    FormRunning.LblHangingFirstWord.Caption = "אוסף מידע עבור עיצוב חלון"
    FormRunning.LblHangingFirstWord.Font.Bold = True: DoEvents
    
    ' Set selection range
    Set selectionRange = Selection.Range
    Call remove(paraCollection)
    
    If skip2Lines = True Then
         Set paraCollection = Helpers.GetValidParagraphRanges(paraCollection, 3)
    Else
        Set paraCollection = Helpers.GetValidParagraphRanges(paraCollection, 2)
    End If
   
    
    If selectionRange.Information(wdInFootnote) Or selectionRange.Information(wdInEndnote) Then
        Call Helpers.PrepareFootnotes(paraCollection)
    End If
        
    ' Loop through paragraphs
    For i = 1 To paraCollection.Count
        Set paraRange = paraCollection(i).Duplicate
        
        FormRunning.LblHangingFirstWord.Caption = "מעצב חלון " & i & " מתוך " & paraCollection.Count: DoEvents
        If stopCode Then GoTo Ending
        
        With paraRange
            .Collapse
            .MoveUntil " "
            .Move
            firstWordSpacing = .Information(wdHorizontalPositionRelativeToTextBoundary)
            .Select
        End With
            
        Selection.EndKey

        With Selection.Range
            .text = vbVerticalTab & Chr(160)
            .Collapse (wdCollapseEnd)
'            Application.ScreenRefresh
            .Previous.Font.Spacing = .Information(wdHorizontalPositionRelativeToTextBoundary) - firstWordSpacing
        End With
                        
         
    Next i
    
    FormRunning.LblHangingFirstWord.Caption = "עיצוב חלון - הושלם": DoEvents
    ' Restore original selection
    selectionRange.Select
    
Ending:
    ' End undo recording
    Application.UndoRecord.EndCustomRecord
End Sub

Sub DoubleWindow(ByVal paraCollection As Collection)
    
    Dim selectionRange As Range
    Dim paraRange As Range
    Dim firstWordSpacing As Double
    Dim currentSpacing As Double
    Dim i As Integer
    
    ' Begin undo recording
    On Error Resume Next
    Application.UndoRecord.StartCustomRecord "עיצוב חלון"
    On Error GoTo Ending
    
    FormRunning.LblHangingFirstWord.Caption = "אוסף מידע עבור עיצוב חלון"
    FormRunning.LblHangingFirstWord.Font.Bold = True: DoEvents
    
    ' Set selection range
    Set selectionRange = Selection.Range
    Call remove(paraCollection)
    Set paraCollection = Helpers.GetValidParagraphRanges(paraCollection, 3)
    
    If selectionRange.Information(wdInFootnote) Or selectionRange.Information(wdInEndnote) Then
        Call Helpers.PrepareFootnotes(paraCollection)
    End If
        
    ' Loop through paragraphs
    For i = 1 To paraCollection.Count
        Set paraRange = paraCollection(i).Duplicate
        
        FormRunning.LblHangingFirstWord.Caption = "מעצב חלון " & i & " מתוך " & paraCollection.Count: DoEvents
        If stopCode Then GoTo Ending
            
        With paraRange
            .Collapse
            .MoveUntil " "
            .Move
             firstWordSpacing = .Information(wdHorizontalPositionRelativeToTextBoundary)
            .Select
        End With

        With Selection
            .EndKey
            .text = vbVerticalTab & Chr(160)
            .Collapse (wdCollapseEnd)
'            Application.ScreenRefresh
            firstWordSpacing = .Information(wdHorizontalPositionRelativeToTextBoundary) - firstWordSpacing
            .Previous.Font.Spacing = firstWordSpacing
            .EndKey
            .text = vbVerticalTab & Chr(160)
            .Collapse (wdCollapseEnd)
            .Previous.Font.Spacing = firstWordSpacing
        End With
                 
        
    Next i

    FormRunning.LblHangingFirstWord.Caption = "עיצוב חלון - הושלם": DoEvents
    ' Restore original selection
    selectionRange.Select
    
Ending:
    ' End undo recording
    Application.UndoRecord.EndCustomRecord
End Sub
Sub DoubleWindowThreeLines(ByVal paraCollection As Collection, skip2Lines As Boolean)
    
    Dim selectionRange As Range
    Dim paraRange As Range
    Dim firstWordSpacing As Double
    Dim currentSpacing As Double
    Dim linesCount As Integer
    Dim i As Integer
    
    ' Begin undo recording
    On Error Resume Next
    Application.UndoRecord.StartCustomRecord "עיצוב חלון"
    On Error GoTo Ending
    
    FormRunning.LblHangingFirstWord.Caption = "אוסף מידע עבור עיצוב חלון"
    FormRunning.LblHangingFirstWord.Font.Bold = True: DoEvents
    
    ' Set selection range
    Set selectionRange = Selection.Range
    Call remove(paraCollection)
    If skip2Lines = True Then
        Set paraCollection = Helpers.GetValidParagraphRanges(paraCollection, 3)
    Else
        Set paraCollection = Helpers.GetValidParagraphRanges(paraCollection, 2)
    End If
    
    If selectionRange.Information(wdInFootnote) Or selectionRange.Information(wdInEndnote) Then
        Call Helpers.PrepareFootnotes(paraCollection)
    End If
        
    ' Loop through paragraphs
    For i = 1 To paraCollection.Count
        Set paraRange = paraCollection(i).Duplicate
        
        FormRunning.LblHangingFirstWord.Caption = "מעצב חלון " & i & " מתוך " & paraCollection.Count: DoEvents
        If stopCode Then GoTo Ending
            
        With paraRange
            linesCount = .ComputeStatistics(wdStatisticLines)
            .Collapse
            .MoveUntil " "
            .Move
            firstWordSpacing = .Information(wdHorizontalPositionRelativeToTextBoundary)
            .Select
        End With

        With Selection
            .EndKey
            .text = vbVerticalTab & Chr(160)
            .Collapse (wdCollapseEnd)
            firstWordSpacing = .Information(wdHorizontalPositionRelativeToTextBoundary) - firstWordSpacing
            .Previous.Font.Spacing = firstWordSpacing
            If linesCount = 3 Then
                .EndKey
                .text = vbVerticalTab & Chr(160)
                .Collapse (wdCollapseEnd)
                .Previous.Font.Spacing = firstWordSpacing
                linesCount = paraRange.Paragraphs(1).Range.ComputeStatistics(wdStatisticLines)
                If linesCount > 3 Then
                    .MoveStart wdCharacter, -2
                    .Delete
                End If
            End If
            
        End With
                             
    Next i

    FormRunning.LblHangingFirstWord.Caption = "עיצוב חלון - הושלם": DoEvents
    ' Restore original selection
    selectionRange.Select
    
Ending:
    ' End undo recording
    Application.UndoRecord.EndCustomRecord
End Sub

Public Sub remove(Optional ByVal paraCollection As Collection = Nothing)
    Dim tempoRange As Range
    Dim paraRange As Range
    Dim i As Integer
    
    On Error GoTo Ending
    Application.UndoRecord.StartCustomRecord "הסרת עיצוב חלון"
    Application.ScreenUpdating = False

    If paraCollection Is Nothing Then Set paraCollection = Helpers.GetAllSelectedParagraphRanges
    
    For i = 1 To paraCollection.Count
        Set paraRange = paraCollection(i).Duplicate
        
        FormRemoveRunning.LblHangingFirstWord.Caption = "מסיר עיצוב חלון " & i & " מתוך " & paraCollection.Count
        If i Mod -Int(-paraCollection.Count / 40) = 0 Then DoEvents
        If stopCode Then GoTo Ending
        
        Set paraRange = paraCollection(i)
        Do
            Set tempoRange = paraRange.Duplicate
            With tempoRange.Find
                .text = vbVerticalTab & Chr(160)
                .Replacement.text = ""
                .Wrap = wdFindStop
                If .Execute Then tempoRange.Delete Else: Exit Do
            End With
        Loop
    Next i
Ending:
    Application.UndoRecord.EndCustomRecord
    Application.ScreenUpdating = True
End Sub
-------------------------------------------------------------------------------
