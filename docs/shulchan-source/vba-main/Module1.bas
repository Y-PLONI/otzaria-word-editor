VBA MACRO Module1.bas 
in file: word/vbaProject.bin - OLE stream: 'VBA/Module1'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit
Sub dafg()
    
    Dim rng As Range
    Dim paraRange As Range
    Dim prevParaRange As Range
    Dim nextParaRange As Range
    Dim tempoRange As Range
    Dim styleName As String
    Dim para As Paragraph
    Dim i As Integer
    
    On Error GoTo Ending
    Set rng = ActiveDocument.Range
    
    styleName = "כותרת 1"

    For i = ActiveDocument.Paragraphs.Count To 1 Step -1
        Set para = ActiveDocument.Paragraphs(i)
        para.Range.Select
        If para.Style = styleName And para.Range.text Like "*" & Chr(13) Then
            Set paraRange = para.Range
            
            On Error GoTo SecondAction
            Set prevParaRange = ActiveDocument.Paragraphs(i - 1).Range
            On Error GoTo Ending
            
            With prevParaRange
                .MoveEnd wdCharacter, 1
                If SearchAndReplace("^b", "False", rng:=prevParaRange.Duplicate, searchMatchWildcards:=False) Is Nothing Then
                    .MoveEnd wdCharacter, -2
'                    .InsertAfter Chr(13)
                    .Collapse direction:=wdCollapseEnd
                    .InsertBreak Type:=wdSectionBreakContinuous
                    .MoveEnd wdCharacter, 1
                    .Delete
                End If
            End With
            
SecondAction:
            On Error GoTo ThirdAction
            Set nextParaRange = ActiveDocument.Paragraphs(i + 1).Range
            On Error GoTo Ending
            nextParaRange.MoveEnd wdCharacter, 1
            With paraRange
                If SearchAndReplace("^b", "False", rng:=nextParaRange.Duplicate, searchMatchWildcards:=False) Is Nothing Then
                    .MoveEnd wdCharacter, -1
                    .InsertAfter Chr(13)
                    .Collapse direction:=wdCollapseEnd
                    .InsertBreak Type:=wdSectionBreakContinuous
                    .Move wdCharacter, -1
                    .Sections.PageSetup.TextColumns.SetCount 1
                    .Move wdCharacter, 1
                    .MoveEnd wdCharacter, 1
                    .Delete
                Else
ThirdAction:
                    On Error GoTo Ending
                    .Sections.PageSetup.TextColumns.SetCount 1
                End If
            End With
        End If
    Next i
Ending:
End Sub
Function SearchStyle(rng As Range, ByVal styleName As String) As Range
    
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .text = ""
        .Replacement.text = ""
        .Style = styleName
        .Wrap = wdFindStop
        .Format = True
        .MatchKashida = False
        .MatchWildcards = True
        If .Execute Then Set SearchStyle = rng.Duplicate
    End With

End Function
-------------------------------------------------------------------------------
