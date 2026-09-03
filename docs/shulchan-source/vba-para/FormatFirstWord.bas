VBA MACRO FormatFirstWord.bas 
in file: word/vbaProject.bin - OLE stream: 'VBA/FormatFirstWord'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit
Option Private Module

Sub ByHand(ByVal paraCollection As Collection, dinami As Boolean, alignmentUpwards As Boolean)

    Dim selectionRange As Range
    Dim paraRange As Range
    Dim i As Integer
    
    ' טעינת הדיאלוג
    If FontDialogueHelper.fontFormatDialog Is Nothing Then _
         FontDialogueHelper.LoadFontDialogSettings
    ' טעינת הדיאלוג הדינמי
    If dinami Then FontDinamiHelper.GetFirstWordDinamiOption
    
    ' Begin undo recording
    On Error Resume Next
    Application.UndoRecord.StartCustomRecord "עיצוב מילה ראשונה"
    Application.ScreenUpdating = False
    On Error GoTo Ending
    
    FormRunning.LblFirstWord.Caption = "אוסף מידע עבור עיצוב מילה ראשונה"
    FormRunning.LblFirstWord.Font.Bold = True: DoEvents
    
    ' Set selection range
    Set selectionRange = Selection.Range
    Set paraCollection = Helpers.GetValidParagraphRangesForFirstWord(paraCollection)
    
    ' Loop through paragraphs
    For i = 1 To paraCollection.Count
        Set paraRange = paraCollection(i).Duplicate
        
        FormRunning.LblFirstWord.Caption = "מעצב מילה ראשונה " & i & " מתוך " & paraCollection.Count
        If i Mod -Int(-paraCollection.Count / 40) = 0 Then DoEvents
        If stopCode Then GoTo Ending
        With paraRange
            .Collapse
            .MoveEndUntil Cset:=" "
            .Select
        End With
        
        FontDialogueHelper.fontFormatDialog.Execute
        
        If dinami Then Call FontDinamiHelper.ApplyDesign(paraRange)
        
        If alignmentUpwards Then paraRange.Font.Position = Helpers.newPosition(paraRange)
    Next i
    
    FormRunning.LblFirstWord.Caption = "עיצוב מילה ראשונה - הושלם": DoEvents
    ' Restore original selection
    selectionRange.Select
        
Ending:
    ' End undo recording
    Application.UndoRecord.EndCustomRecord
    Application.ScreenUpdating = True
End Sub

Sub ByStyle(ByVal paraCollection As Collection, dinami As Boolean, alignmentUpwards As Boolean)

    Dim selectionRange As Range
    Dim paraRange As Range
    Dim firstWordStyle As style
    Dim i As Integer
    
    ' Begin undo recording
    On Error Resume Next
    Application.UndoRecord.StartCustomRecord "עיצוב מילה ראשונה"
    Application.ScreenUpdating = False
    On Error GoTo Ending
    
    FormRunning.LblFirstWord.Caption = "אוסף מידע עבור עיצוב מילה ראשונה"
    FormRunning.LblFirstWord.Font.Bold = True: DoEvents
    
    ' Set selection range
    Set selectionRange = Selection.Range
    Set paraCollection = Helpers.GetValidParagraphRangesForFirstWord(paraCollection)
    
    'prepare style
    Set firstWordStyle = Helpers.GetFirstWordStyle()
    If dinami Then FontDinamiHelper.GetFirstWordDinamiOption
    
    ' Loop through paragraphs
    For i = 1 To paraCollection.Count
        Set paraRange = paraCollection(i).Duplicate
        FormRunning.LblFirstWord.Caption = "מעצב מילה ראשונה " & i & " מתוך " & paraCollection.Count
        If i Mod -Int(-paraCollection.Count / 40) = 0 Then DoEvents
        If stopCode Then GoTo Ending
        
        With paraRange
            .Collapse
            .MoveEndUntil Cset:=" "
            .Font.Reset
            .style = firstWordStyle
            If alignmentUpwards Then .Font.Position = Helpers.newPosition(paraRange)
            .Characters.First.Select
        End With
        
        If dinami Then Call FontDinamiHelper.ApplyDesign(paraRange)
        If alignmentUpwards Then paraRange.Font.Position = Helpers.newPosition(paraRange)
        
    Next i
    
    FormRunning.LblFirstWord.Caption = "עיצוב מילה ראשונה - הושלם": DoEvents
     
    ' Restore original selection
    selectionRange.Select
        
Ending:
    ' End undo recording
    Application.UndoRecord.EndCustomRecord
    Application.ScreenUpdating = True
End Sub

Sub remove(Optional ByVal paraCollection As Collection = Nothing)
    Dim selectionRange As Range
    Dim paraRange As Range
    Dim txt As String
    Dim i As Integer
    
    ' Begin undo recording
    On Error Resume Next
    Application.UndoRecord.StartCustomRecord "הסרת עיצוב מילה ראשונה"
    Application.ScreenUpdating = False
    On Error GoTo Ending
    
    FormRemoveRunning.LblFirstWord.Caption = "אוסף מידע עבור הסרת עיצוב מילה ראשונה"
    FormRemoveRunning.LblFirstWord.Font.Bold = True
    
    'set range
    Set selectionRange = Selection.Range
    If paraCollection Is Nothing Then
        Set paraCollection = Helpers.GetAllSelectedParagraphRanges
    End If
    
    ' Loop through paragraphs
    For i = 1 To paraCollection.Count
        Set paraRange = paraCollection(i).Duplicate
        
        FormRemoveRunning.LblFirstWord.Caption = "מסיר עיצוב מילה ראשונה " & i & " מתוך " & paraCollection.Count
        If i Mod -Int(-paraCollection.Count / 40) = 0 Then DoEvents
        If stopCode Then GoTo Ending
            
        With paraRange
            .Collapse
            .MoveEndUntil Cset:=" "
            .MoveEnd
             txt = .text
            .text = ""
            .Select
            .text = txt
        End With
        
    Next i
    
    FormRemoveRunning.LblFirstWord.Caption = "הסרת עיצוב מילה ראשונה - הושלם"
    
    'reset selection
    selectionRange.Select
    
Ending:
    ' End undo recording
    Application.UndoRecord.EndCustomRecord
    Application.ScreenUpdating = True
End Sub
-------------------------------------------------------------------------------
