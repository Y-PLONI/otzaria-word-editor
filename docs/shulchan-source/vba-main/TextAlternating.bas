VBA MACRO TextAlternating.bas 
in file: word/vbaProject.bin - OLE stream: 'VBA/TextAlternating'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit
Option Private Module

Sub Starting(endChar As String, startChar As String)
    
    Dim rng As Range
    Dim para As Paragraph
    
    On Error GoTo Ending
    Application.UndoRecord.StartCustomRecord "עיצוב טקסט מתחלף"
    
    Set rng = Selection.Range
    
    For Each para In rng.Paragraphs
        
        With para.Range
            .Collapse direction:=wdCollapseStart
            .MoveEndUntil endChar & Chr(13)
            .MoveEnd unit:=wdCharacter, Count:=1
            If Not .text Like "*" & Chr(13) & "*" Then
                .Font.Bold = True
                .Font.BoldBi = True
            End If
            
            Do
                .Collapse direction:=wdCollapseEnd
                .MoveUntil startChar & Chr(13)
                .MoveEnd unit:=wdCharacter, Count:=2
                If .text Like "*" & Chr(13) & "*" Then Exit Do
                .Collapse direction:=wdCollapseEnd
                .MoveEndUntil endChar & Chr(13)
                .MoveEnd unit:=wdCharacter, Count:=1
                If .text Like "*" & Chr(13) & "*" Then Exit Do
                .Font.Bold = True
                .Font.BoldBi = True
            Loop
        
        End With
    
    Next para
    
Ending:
    
    Application.UndoRecord.EndCustomRecord

End Sub
-------------------------------------------------------------------------------
