VBA MACRO Helpers.bas 
in file: word/vbaProject.bin - OLE stream: 'VBA/Helpers'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit
Option Private Module

Sub PrepareFootnotes(ByVal paraCollection As Collection)
    Dim paraRange As Range
    
    For Each paraRange In paraCollection
        With paraRange.Find
            .text = "^f"
            .Replacement.text = "^&ֵ%%" & Chr(160)
            .Wrap = wdFindStop
            .Execute Replace:=wdReplaceAll
        End With
        
         With paraRange.Find
            .text = "%%" & Chr(160) & " "
            .Replacement.text = Chr(160)
            .Wrap = wdFindStop
            .Execute Replace:=wdReplaceAll
        End With
    Next paraRange
End Sub
Function GetAllSelectedParagraphRanges() As Collection
    
    Dim rng As Range
    Dim para As Paragraph
    Dim paraCollection As New Collection

    Set rng = Selection.Range
    For Each para In rng.Paragraphs
        paraCollection.Add para.Range
    Next para
    
    Set GetAllSelectedParagraphRanges = paraCollection

End Function
Function GetNonHeadingSelectedParagraphRanges() As Collection
    
    Dim rng As Range
    Dim para As Paragraph
    Dim paraCollection As New Collection
    Dim i As Integer
    
    Set rng = Selection.Range
    For i = 1 To rng.Paragraphs.Count
        FormRunning.Label7.Caption = "אוסף מידע " & i & " מתוך " & rng.Paragraphs.Count
        FormRemoveRunning.Label7.Caption = "אוסף מידע " & i & " מתוך " & rng.Paragraphs.Count
        DoEvents
        If stopCode Then GoTo Ending
        Set para = rng.Paragraphs(i)
        If para.OutlineLevel = 10 And _
            para.Alignment <> wdAlignParagraphCenter _
            Then
                paraCollection.Add para.Range
        End If
    Next i
Ending:
    Set GetNonHeadingSelectedParagraphRanges = paraCollection

End Function
Function GetSpecificSelectedParagraphRanges() As Collection
    
    Dim rng As Range
    Dim para As Paragraph
    Dim paraCollection As New Collection
    Dim stlName(0 To 14) As String
    Dim i As Integer
    
    Set rng = Selection.Range
    
    For i = LBound(stlName) To UBound(stlName)
        stlName(i) = SettingsHelper.GetSavedSetting(appName, "ParagraphFormat", "stlName" & i, "")
    Next i
    
    For i = 1 To rng.Paragraphs.Count
        FormRunning.Label7.Caption = "אוסף מידע " & i & " מתוך " & rng.Paragraphs.Count
        FormRemoveRunning.Label7.Caption = "אוסף מידע " & i & " מתוך " & rng.Paragraphs.Count
        DoEvents
        If stopCode Then GoTo Ending
        Set para = rng.Paragraphs(i)
        If IsStyleInArray(para, stlName) Then
            paraCollection.Add para.Range
        End If
    Next i
    
Ending:
    Set GetSpecificSelectedParagraphRanges = paraCollection

End Function
Function GetNonSpecificSelectedParagraphRanges() As Collection
    
    Dim rng As Range
    Dim para As Paragraph
    Dim paraCollection As New Collection
    Dim stlName(0 To 14) As String
    Dim i As Integer
    
    Set rng = Selection.Range
    
    For i = LBound(stlName) To UBound(stlName)
        stlName(i) = SettingsHelper.GetSavedSetting(appName, "ParagraphFormat", "stlName" & i, "")
    Next i
    
    For i = 1 To rng.Paragraphs.Count
        FormRunning.Label7.Caption = "אוסף מידע " & i & " מתוך " & rng.Paragraphs.Count
        FormRemoveRunning.Label7.Caption = "אוסף מידע " & i & " מתוך " & rng.Paragraphs.Count
        DoEvents
        If stopCode Then GoTo Ending
        Set para = rng.Paragraphs(i)
        If Not IsStyleInArray(para, stlName) Then
            paraCollection.Add para.Range
        End If
    Next i
    
Ending:
    Set GetNonSpecificSelectedParagraphRanges = paraCollection

End Function
Function GetValidParagraphRanges(ByVal prevParaCollection As Collection, targetLineCount As Integer) As Collection
     Dim paraRange As Range
     Dim paraCollection As New Collection
     
     For Each paraRange In prevParaCollection
        If paraRange.ComputeStatistics(wdStatisticLines) >= targetLineCount And _
            paraRange.ParagraphFormat.Alignment <> wdAlignParagraphCenter _
            Then
               paraCollection.Add paraRange
        End If
    Next paraRange
    
    Set GetValidParagraphRanges = paraCollection
End Function
Function GetValidParagraphRangesForLastLine(ByVal prevParaCollection As Collection, targetLineCount As Integer) As Collection
    
    Dim paraRange As Range
    Dim paraLineCount As Integer
    Dim paraTargetLineCount As Integer
    Dim paraCollection As New Collection
    
    For Each paraRange In prevParaCollection
        paraLineCount = paraRange.ComputeStatistics(wdStatisticLines)
        ' בדיקה אם הטווח מכיל עיצוב חלון
        paraTargetLineCount = targetLineCount
        If paraRange.text Like "*" & Chr(11) & Chr(160) & "*" & Chr(11) & Chr(160) & "*" Then
            If targetLineCount < 4 Then paraTargetLineCount = 4
        ElseIf paraRange.text Like "*" & Chr(11) & Chr(160) & "*" Then
            If targetLineCount < 3 Then paraTargetLineCount = 3
        End If
        
        If paraLineCount >= paraTargetLineCount And _
            paraRange.ParagraphFormat.Alignment <> wdAlignParagraphCenter _
            Then
                paraCollection.Add paraRange
        End If
    Next paraRange
    
    Set GetValidParagraphRangesForLastLine = paraCollection
End Function

Function GetValidParagraphRangesForFirstWord(ByVal prevParaCollection As Collection) As Collection
    
    Dim paraRange As Range
    Dim paraCollection As New Collection
    Dim isFootNotes As Boolean
     
    If prevParaCollection(1).Information(wdInFootnote) Or prevParaCollection(1).Information(wdInEndnote) Then isFootNotes = True
    
    For Each paraRange In prevParaCollection
        With paraRange
            If .text Like "* *" Then
                
                If isFootNotes = True And Asc(.Characters.First) = 2 Then ' adjustment for footnote references
                    .MoveStartUntil " "
                    .MoveStart
                End If
                
                paraCollection.Add paraRange
                
            End If
        End With
    Next paraRange
    
    Set GetValidParagraphRangesForFirstWord = paraCollection
End Function
Function GetValidParagraphRangesForLineSpacing(ByVal prevParaCollection As Collection, ByVal isRemove As Boolean) As Collection
    
    Dim paraRange As Range
    Dim paraCollection As New Collection
    Dim isFootNotes As Boolean
     
    If prevParaCollection(1).Information(wdInFootnote) Or prevParaCollection(1).Information(wdInEndnote) Then isFootNotes = True
    
    For Each paraRange In prevParaCollection
        With paraRange
            If .text Like "* *" And ( _
                .ParagraphFormat.LineSpacingRule <> wdLineSpaceExactly Or ( _
                isRemove And .ParagraphFormat.LineSpacingRule = wdLineSpaceExactly)) Then
                
                    If isFootNotes = True And Asc(.Characters.First) = 2 Then ' adjustment for footnote references
                        .MoveStartUntil " "
                        .MoveStart
                    End If
                    
                    paraCollection.Add paraRange
                
            End If
        End With
    Next paraRange
    
    Set GetValidParagraphRangesForLineSpacing = paraCollection
End Function
Public Function GetFirstWordStyle() As style
    Dim targetStyle As style
    Dim targetStyleName As String
    
    targetStyleName = GetSavedSetting(RibbonControl.appName, "ParagraphFormat", "FirstWordStyle", "מילה ראשונה")
    If targetStyleName = "" Then targetStyleName = "מילה ראשונה"
    
    For Each targetStyle In Application.ActiveDocument.Styles
        If targetStyle.NameLocal = targetStyleName Then
            Set GetFirstWordStyle = targetStyle
            Exit Function
        End If
    Next
        
    Set targetStyle = Application.ActiveDocument.Styles.Add(targetStyleName, wdStyleTypeCharacter)
    With targetStyle.Font
        .Bold = 1
        .BoldBi = 1
        .Size = .Size + 2
        .SizeBi = .SizeBi + 2
    End With
    targetStyle.QuickStyle = True
    
    Set GetFirstWordStyle = targetStyle
End Function

Function newPosition(paraRange As Range) As Double
    
    Dim fontSize As Double
    Dim sizeAfter As Double
    
    fontSize = paraRange.Font.SizeBi
    sizeAfter = ActiveDocument.Range(paraRange.End, paraRange.End + 1).Font.SizeBi
    newPosition = -(fontSize - sizeAfter) / 2

End Function
Function GetFontProperties(paraRange As Range, ByRef fontName As String, ByRef fontSize As Single)
    Dim tempoRange As Range
    Set tempoRange = paraRange.Duplicate
    With tempoRange
        .Collapse wdCollapseStart
        .MoveUntil " " & Chr(13)
        .MoveEnd wdCharacter, 1
        fontName = .Font.NameBi
        fontSize = .Font.SizeBi
    End With
End Function
Function IsStyleInArray(para As Paragraph, stlNameArray() As String) As Boolean
    Dim stl As String
    Dim i As Integer
    
    For i = LBound(stlNameArray) To UBound(stlNameArray)
        stl = stlNameArray(i)
        If para.style.NameLocal = stl Then
            IsStyleInArray = True
            Exit Function
        End If
    Next i
End Function

-------------------------------------------------------------------------------
