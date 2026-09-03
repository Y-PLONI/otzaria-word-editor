VBA MACRO LastLineBalance.bas 
in file: word/vbaProject.bin - OLE stream: 'VBA/LastLineBalance'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit
Option Private Module

Sub Repair(ByVal paraCollection As Collection, skip2Lines As Boolean)

    Dim selectionRange As Range
    Dim paraRange As Range
    Dim minLineRng As Range
    Dim minLineNum As Integer
    Dim i, x As Integer
    
    On Error GoTo Ending
    Application.UndoRecord.StartCustomRecord "איזון שורה אחרונה"
    
    FormRunning.LblBalanceLastLine.Caption = "אוסף מידע עבור איזון שורה אחרונה"
    FormRunning.LblBalanceLastLine.Font.Bold = True: DoEvents
    
    Set selectionRange = Selection.Range
    
    If skip2Lines Then
        Set paraCollection = Helpers.GetValidParagraphRangesForLastLine(paraCollection, 3)
    Else
        Set paraCollection = Helpers.GetValidParagraphRangesForLastLine(paraCollection, 2)
    End If
    
    For i = 1 To paraCollection.Count
        Set paraRange = paraCollection(i).Duplicate
        
        FormRunning.LblBalanceLastLine.Caption = "מאזן שורה אחרונה " & i & " מתוך " & paraCollection.Count: DoEvents
        If stopCode Then GoTo Ending
       
        If OneWord(paraRange) Then
        
            If paraRange.text Like "*" & Chr(11) & Chr(160) & "*" & Chr(11) & Chr(160) & "*" Then
                minLineNum = 3
            ElseIf paraRange.text Like "*" & Chr(11) & Chr(160) & "*" Then
                minLineNum = 2
            Else
                minLineNum = 1
            End If
            
            For x = 1 To 10
                Set minLineRng = MinLineSearch(paraRange, minLineNum)
                Call SpaceExpansion(minLineRng)
                If Not OneWord(paraRange) Then Exit For
            Next x
        
        End If
    Next i
    
    FormRunning.LblBalanceLastLine.Caption = "איזון שורה אחרונה - הושלם": DoEvents

Ending:
    Application.UndoRecord.EndCustomRecord
    selectionRange.Select
End Sub

Function OneWord(paraRange As Range) As Boolean
    
    paraRange.Select
    With Selection
        .Collapse direction:=wdCollapseEnd
        Do While Mid(ActiveDocument.Range(.Start - 1, .Start).text, 1, 1) = " " Or Mid(ActiveDocument.Range(.Start - 1, .Start).text, 1, 1) = Chr(13) Or Mid(ActiveDocument.Range(.Start - 1, .Start).text, 1, 1) = Chr(9)
            .Move wdCharacter, -1
        Loop
        .MoveStart wdLine, -1
            
        If .text Like "* *" Then
            OneWord = False
        Else
            OneWord = True
        End If
        .MoveStart wdCharacter, -2
        ' דילוג במקרה שיש תו מעבר שורה
        If .text Like "*" & Chr(11) & "*" Then OneWord = False
    End With
End Function

Function MinLineSearch(paraRange As Range, ByRef minLineNum As Integer) As Range
    
    Dim i As Integer
    Dim lineCount As Integer
    Dim spacious As Double
    Dim startChar As Double
    
    paraRange.Select
    
    lineCount = paraRange.ComputeStatistics(wdStatisticLines)
    spacious = 1000
    
    For i = minLineNum To lineCount - 1
        paraRange.Select
        With Selection
            .Collapse direction:=wdCollapseStart
            .Move wdLine, i - 1
            .MoveUntil " "
            startChar = .Information(wdHorizontalPositionRelativeToPage)
            .Move Unit:=wdCharacter
            
            If spacious > startChar - .Information(wdHorizontalPositionRelativeToPage) Then
                spacious = startChar - .Information(wdHorizontalPositionRelativeToPage)
                .MoveStart wdLine, -1
                .MoveEnd wdLine, 1
                .MoveEnd wdCharacter, -1
                Set MinLineSearch = .Range
                minLineNum = i
            End If
        End With
    Next i
End Function
Function SpaceExpansion(minLineRng As Range)
    
    Dim minLineCount As Integer
    Dim i As Single
    Dim char As Range
    
    If minLineRng Is Nothing Then Exit Function
    
    minLineCount = minLineRng.Words.Count
    
    For i = 0.5 To 50 Step 0.5
        
        For Each char In minLineRng.Characters
            
            With char
                If .text = " " Then .Font.Spacing = i
            End With
        
        Next char
            
        minLineRng.Select
        With Selection
            .Collapse direction:=wdCollapseStart
            .MoveEnd wdLine, 1
            .MoveEnd wdCharacter, -1
            If minLineCount > .Words.Count Then Exit For
        End With
    Next i
End Function
Sub remove(ByVal paraCollection As Collection)
    
    Dim selectionRange As Range
    Dim paraRange As Range
    Dim char As Range
    Dim i As Integer
    
    On Error GoTo Ending
    
    Application.UndoRecord.StartCustomRecord "הסרת איזון שורה אחרונה"
    Application.ScreenUpdating = False
    
    Set selectionRange = Selection.Range
    selectionRange.Expand Unit:=wdParagraph
    
    For i = 1 To paraCollection.Count
        FormRemoveRunning.LblBalanceLastLine.Caption = "מסיר איזון שורה אחרונה " & i & " מתוך " & paraCollection.Count
        If i Mod -Int(-paraCollection.Count / 40) = 0 Then DoEvents
        If stopCode Then GoTo Ending
        
        Set paraRange = paraCollection(i)
        
        For Each char In paraRange.Characters
            With char
                If .text = " " Then .Font.Spacing = ActiveDocument.Range(.Start - 1, .Start).Font.Spacing
            End With
        Next char
    Next i
    
Ending:
    Application.UndoRecord.EndCustomRecord
    Application.ScreenUpdating = True
    selectionRange.Select
End Sub
-------------------------------------------------------------------------------
