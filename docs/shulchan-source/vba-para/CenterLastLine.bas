VBA MACRO CenterLastLine.bas 
in file: word/vbaProject.bin - OLE stream: 'VBA/CenterLastLine'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit
Option Private Module

Sub Execute(ByVal paraCollection As Collection, skip2Lines As Boolean)

    Dim selectionRange As Range
    Dim paraRange As Range
    Dim i As Integer
    ' Begin undo recording
    On Error Resume Next
    Application.UndoRecord.StartCustomRecord "מירכוז שורה אחרונה"
    Application.ScreenUpdating = False
    On Error GoTo Ending
    
    FormRunning.LblCenterLastLine.Caption = "אוסף מידע עבור מירכוז שורה אחרונה"
    FormRunning.LblCenterLastLine.Font.Bold = True: DoEvents
    ' Set selection range
    Set selectionRange = Selection.Range
    Call remove(paraCollection)
        
    If skip2Lines Then
         Set paraCollection = Helpers.GetValidParagraphRangesForLastLine(paraCollection, 3)
    Else
        Set paraCollection = Helpers.GetValidParagraphRangesForLastLine(paraCollection, 2)
    End If

    Application.ActiveDocument.DefaultTabStop = 0
    
    ' Loop through paragraphs
    For i = 1 To paraCollection.Count
        Set paraRange = paraCollection(i)
        
        FormRunning.LblCenterLastLine.Caption = "ממרכז שורה אחרונה " & i & " מתוך " & paraCollection.Count
        If i Mod -Int(-paraCollection.Count / 40) = 0 Then DoEvents
        If stopCode Then GoTo Ending

        With paraRange
            .ParagraphFormat.Alignment = wdAlignParagraphDistribute
            .End = .End - 1
            .InsertAfter Chr(9)
        End With
        
    Next i

    FormRunning.LblCenterLastLine.Caption = "מירכוז שורה אחרונה - הושלם": DoEvents

Ending:
    ' End undo recording
    Application.UndoRecord.EndCustomRecord
    Application.ScreenUpdating = True
    selectionRange.Select
End Sub

Sub remove(Optional ByVal paraCollection As Collection = Nothing)
    Dim paraRange As Range
    Dim tempoRange As Range
    Dim i As Integer
    
    On Error GoTo Ending
    Application.UndoRecord.StartCustomRecord "הסרת מירכוז שורה אחרונה"
    Application.ScreenUpdating = False
    
    If paraCollection Is Nothing Then
        Set paraCollection = Helpers.GetAllSelectedParagraphRanges
    End If
    
    For i = 1 To paraCollection.Count
        FormRemoveRunning.LblCenterLastLine.Caption = "מסיר מירכוז שורה אחרונה " & i & " מתוך " & paraCollection.Count
        If i Mod -Int(-paraCollection.Count / 40) = 0 Then DoEvents
        If stopCode Then GoTo Ending
        
        Set paraRange = paraCollection(i)
        
        paraRange.ParagraphFormat.TabStops.ClearAll
        Set tempoRange = paraRange.Duplicate
        With tempoRange.Find
            .ParagraphFormat.Alignment = wdAlignParagraphDistribute
            .text = "^t^p"
            .Wrap = wdFindStop
            If .Execute Then
                tempoRange.Characters.First.Delete
                paraRange.ParagraphFormat.Alignment = wdAlignParagraphJustify
            End If
        End With
        Set tempoRange = paraRange.Duplicate
        With tempoRange.Find
            .text = "^l^t"
            .Replacement.text = ""
            .Wrap = wdFindStop
            If .Execute Then tempoRange.Delete
        End With
    Next i
Ending:
    Application.UndoRecord.EndCustomRecord
    Application.ScreenUpdating = True
End Sub
-------------------------------------------------------------------------------
