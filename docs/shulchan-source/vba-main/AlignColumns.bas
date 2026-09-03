VBA MACRO AlignColumns.bas 
in file: word/vbaProject.bin - OLE stream: 'VBA/AlignColumns'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit
Option Private Module

Sub Repair()
    
    Dim rng As Range
    Dim firstPage As Integer
    Dim lastPage As Integer
    Dim pageRange As Range
    Dim secCollection As New Collection
    Dim secRange As Range
    Dim columnRange(1 To 2) As Range
    Dim columnPosition(1 To 2) As Double
    Dim nextSecRange As Range
    Dim nextSecPosition As Double
    Dim columnSpacing As Double
    Dim shortColumnRange As Range
    Dim shortColumnPosition As Double
    Dim para As Paragraph
    Dim allParagraphsSpace As Single
    Dim paragraphsWithSpace As Collection
    Dim pointValue As Double
    Dim i, x, counter As Integer
    
    On Error GoTo Ending
    Application.UndoRecord.StartCustomRecord "יישור טורים"
    Application.ScreenUpdating = False
    Set rng = Selection.Range
    
    rng.Characters.First.Select
    firstPage = Selection.Information(wdActiveEndPageNumber)
    rng.Characters.Last.Select
    lastPage = Selection.Information(wdActiveEndPageNumber)
    FormRunning.Show vbModeless
    
    For x = firstPage To lastPage
        counter = counter + 1
        FormRunning.Label1.Caption = "מיישר טורים עמוד " & counter & " מתוך " & lastPage - firstPage + 1: DoEvents
        If stopCode Then GoTo Ending
        
        Selection.GoTo what:=wdGoToPage, Count:=x
        Set pageRange = Selection.Bookmarks("\Page").Range
        
        Set secCollection = GetSectionsRng(pageRange, rng)
        
        For Each secRange In secCollection
        
            ' קבלת טווח הטורים
            Set columnRange(1) = Helpers.GetColumnRange(secRange.Duplicate)
            Set columnRange(2) = secRange.Duplicate
            columnRange(2).Start = columnRange(1).End
            If columnRange(2).Characters.Last = Chr(12) Then columnRange(2).End = columnRange(2).End - 1
            
            Set nextSecRange = secRange.Duplicate
            nextSecRange.Start = secRange.End
            nextSecRange.End = secRange.End + 1
            nextSecPosition = nextSecRange.Information(wdVerticalPositionRelativeToPage)
            
            ' קביעת הטור הקצר
            For i = 1 To 2
                columnRange(i).End = columnRange(i).End - 1
                Set para = columnRange(i).Paragraphs.Last
                columnRange(i).Characters.Last.Select
                columnPosition(i) = GetLineHeight2(columnRange(i)) + Selection.Information(wdVerticalPositionRelativeToPage)
            Next i
               
            If columnPosition(1) < columnPosition(2) Then
                Set shortColumnRange = columnRange(1)
                shortColumnPosition = columnPosition(1)
'               If columnRange(1).Characters.Last = Chr(13) Then columnRange(1).Paragraphs.Last.SpaceAfter = columnRange(2).Paragraphs.Last.SpaceAfter
            Else
                Set shortColumnRange = columnRange(2)
                shortColumnPosition = columnPosition(2)
'               If columnRange(2).Characters.Last = Chr(13) Then columnRange(2).Paragraphs.Last.SpaceAfter = columnRange(1).Paragraphs.Last.SpaceAfter
            End If
            
            columnSpacing = Abs(columnPosition(1) - columnPosition(2))
            
            ' קבלת המרווח בין הפיסקאות
            Set paragraphsWithSpace = New Collection
            GetAllParagraphsSpace shortColumnRange, paragraphsWithSpace, allParagraphsSpace
                        
            ' תיקון המרווח בין הפיסקאות וקבלת השארית
            If allParagraphsSpace > 0 Then
                pointValue = columnSpacing / allParagraphsSpace + 1
            Else
                pointValue = 1
            End If
            columnSpacing = columnSpacing - RepairParagraphsSpace(shortColumnRange, pointValue)
            
            ' תיקון השארית
            pointValue = 0.1
            For Each para In paragraphsWithSpace
                If columnSpacing < pointValue Then Exit For
                With para
                    .SpaceAfter = .SpaceAfter + pointValue
                End With
                columnSpacing = columnSpacing - pointValue
            Next para
            
            ' בדיקה שהשורה האחרונה לא ברחה קדימה
            Call TestLastLine(shortColumnRange, shortColumnPosition)
            
            ' בדיקה שהפיסקה הבאה לא ברחה קדימה
            Call TestNextPara(nextSecRange, nextSecPosition, shortColumnRange)
        
            ' תיקון אחרון
            Call LastRepair(shortColumnRange, shortColumnPosition)
        
        Next secRange
    Next x
    
Ending:
    Application.UndoRecord.EndCustomRecord
    Application.ScreenUpdating = True
    Unload FormRunning
End Sub
Function GetSectionsRng(pageRange As Range, rng) As Collection
    
    Dim pageNum As Integer
    Dim sec As section
    Dim secRange As Range
    Dim tempoCollection As New Collection
    
    pageNum = pageRange.Information(wdActiveEndPageNumber)
    
    For Each sec In pageRange.Sections
        Set secRange = sec.Range
        If secRange.Start < pageRange.Start Then
            secRange.Start = pageRange.Start
        End If
        
        If secRange.End > pageRange.End Then
            secRange.End = pageRange.End
        End If
        
        If secRange.Start >= rng.Start And secRange.End <= rng.End _
            And sec.PageSetup.TextColumns.Count = 2 _
            Then
                If Not Helpers.SearchAndReplace("^m", "False", rng:=secRange.Duplicate, searchMatchWildcards:=False) Is Nothing Or _
                    (sec.Index = ActiveDocument.Sections.Count And pageNum = ActiveDocument.ComputeStatistics(wdStatisticPages)) Then
                        Call AddSecBreack(secRange.Duplicate)
                End If
                tempoCollection.Add secRange
        End If
    Next sec
    
    Set GetSectionsRng = tempoCollection

End Function
Function GetLineHeight(para As Paragraph) As Double

    Dim tempoRange As Range
    Dim widowControlBefore As Boolean
    
    If para.LineSpacingRule = wdLineSpaceExactly Then
        GetLineHeight = para.LineSpacing
        Exit Function
    End If
    Set tempoRange = para.Range.Duplicate
    
    With tempoRange
        .Collapse wdCollapseStart
        .MoveUntil " " & Chr(13)
        .MoveEnd wdWord, 1
    End With
    
    With para.Range
        widowControlBefore = .ParagraphFormat.WidowControl
        .ParagraphFormat.WidowControl = True
        .Collapse wdCollapseStart
        .InsertBefore " " & Chr(11)
        With .Font
            .Position = 0
            .NameBi = tempoRange.Font.NameBi
            .name = tempoRange.Font.name
            .SizeBi = tempoRange.Font.SizeBi
            .Size = tempoRange.Font.Size
        End With
        .Collapse wdCollapseEnd
        GetLineHeight = .Information(wdVerticalPositionRelativeToTextBoundary)
        .MoveStart wdCharacter, -2
        .Delete
        .ParagraphFormat.WidowControl = widowControlBefore
    End With
    
    Application.ScreenRefresh
End Function
Function GetLineHeight2(columnRange As Range) As Double

    Dim tempoRange As Range
    Dim widowControlBefore As Boolean
    Dim paraSpaceBefore As Single
    
    Set tempoRange = columnRange.Duplicate
    
    With tempoRange
        If .Paragraphs.Last.LineSpacingRule = wdLineSpaceExactly Then
            GetLineHeight2 = .Paragraphs.Last.LineSpacing
            Exit Function
        End If
        .Collapse wdCollapseEnd
        .Select
    End With
    
    Selection.MoveStart wdLine, -1
    Set tempoRange = Selection.Range
    RemoveLastChar tempoRange
    With tempoRange
'        paraSpaceBefore = .Paragraphs(1).SpaceBefore
'        .Paragraphs(1).SpaceBefore = 0
        widowControlBefore = .ParagraphFormat.WidowControl
        .ParagraphFormat.WidowControl = True
        .InsertBefore Chr(13)
        .InsertAfter Chr(11) & " "
        Application.ScreenRefresh
        GetLineHeight2 = .Characters.Last.Information(wdVerticalPositionRelativeToTextBoundary)
        GetLineHeight2 = GetLineHeight2 - .Paragraphs(1).SpaceBefore
        .Characters.Last.Delete
        .Characters.Last.Delete
        .Characters.First.Delete
        .ParagraphFormat.WidowControl = widowControlBefore
'        .Paragraphs(1).SpaceBefore = paraSpaceBefore
    End With
    
End Function
Function GetAllParagraphsSpace(shortColumnRange As Range, ByRef paragraphsWithSpace As Collection, ByRef allParagraphsSpace As Single)
    
    Dim para As Paragraph
    Dim i As Integer
    
    allParagraphsSpace = 0
    For i = 1 To shortColumnRange.Paragraphs.Count
        Set para = shortColumnRange.Paragraphs(i)
        If i > 1 Then
'            allParagraphsSpace = allParagraphsSpace + para.SpaceBefore
        End If
        If i < shortColumnRange.Paragraphs.Count Then
            allParagraphsSpace = allParagraphsSpace + para.SpaceAfter
            paragraphsWithSpace.Add para
        End If
    Next i
    
End Function
Function RepairParagraphsSpace(shortColumnRange As Range, pointValue As Double) As Double

    Dim para As Paragraph
    Dim i As Integer
    Dim spaceBeforeRepair
    i = 0
    For Each para In shortColumnRange.Paragraphs
        i = i + 1
        With para
            If i > 1 Then
'                spaceBeforeRepair = .SpaceBefore
'                .SpaceBefore = FloorDecimal(.SpaceBefore * pointValue, 1)
'                RepairParagraphsSpace = RepairParagraphsSpace + .SpaceBefore - spaceBeforeRepair
            End If
            If i < shortColumnRange.Paragraphs.Count Then
                spaceBeforeRepair = .SpaceAfter
                .SpaceAfter = FloorDecimal(.SpaceAfter * pointValue, 1)
                RepairParagraphsSpace = RepairParagraphsSpace + .SpaceAfter - spaceBeforeRepair
            End If
        End With
    Next para

End Function
Function TestLastLine(shortColumnRange As Range, shortColumnPosition As Double)
    
    Dim lastPara As Paragraph
    Dim lastParaSpace As Single
    
    Set lastPara = shortColumnRange.Paragraphs.Last
    lastParaSpace = lastPara.SpaceAfter
    shortColumnRange.Characters.Last.Select
    Do While shortColumnPosition > shortColumnPosition / 3 + Selection.Information(wdVerticalPositionRelativeToPage)
        If lastPara.SpaceAfter > 1 Then
            lastPara.SpaceAfter = lastPara.SpaceAfter - 1
        ElseIf lastPara.SpaceAfter > 0 Then
            lastPara.SpaceAfter = 0
        Else
            lastPara.SpaceAfter = lastParaSpace
            Exit Do
        End If
    Loop
        
End Function
Function TestNextPara(nextSecRange, nextSecPosition, shortColumnRange)
    
    Dim lastPara As Paragraph
    Dim lastParaSpace As Single
    
    Set lastPara = shortColumnRange.Paragraphs.Last
    lastParaSpace = lastPara.SpaceAfter
    
    Do Until nextSecPosition + 0.1 > nextSecRange.Information(wdVerticalPositionRelativeToPage) And nextSecPosition - 0.1 < nextSecRange.Information(wdVerticalPositionRelativeToPage)
        If lastPara.SpaceAfter > 0.1 Then
            lastPara.SpaceAfter = lastPara.SpaceAfter - 0.1
        Else
            lastPara.SpaceAfter = lastParaSpace
            Exit Do
        End If
    Loop
    
End Function
Function LastRepair(shortColumnRange As Range, shortColumnPosition As Double)
    
    Dim para As Paragraph
    Dim i, x As Integer
    
    shortColumnRange.Characters.Last.Select
    Do While shortColumnPosition > shortColumnPosition / 3 + Selection.Information(wdVerticalPositionRelativeToPage)
        x = x + 1
        For i = 1 To shortColumnRange.Paragraphs.Count - 1
            Set para = shortColumnRange.Paragraphs(i)
            If para.SpaceAfter > 0.1 Then para.SpaceAfter = para.SpaceAfter - 0.1
        Next i
        If x = 20 Then Exit Do
    Loop
        
End Function
Function FloorDecimal(ByVal num As Double, ByVal decimalPlaces As Integer) As Double
    Dim factor As Double
    factor = 10 ^ decimalPlaces
    FloorDecimal = Int(num * factor) / factor
End Function

Function RangeHasSectionBreak() '(sec As Section) As Boolean
    
    Dim rng As Range
    Dim brk As Break
    
    Set rng = Selection.Range
    RangeHasSectionBreak = False
    For Each brk In ActiveDocument.Breaks
        If brk.Range.Start >= rng.Start And brk.Range.End <= rng.End Then
            If brk = wdSectionBreakNextPage _
               Or brk = wdSectionBreakContinuous _
               Or brk = wdSectionBreakEvenPage _
               Or brk = wdSectionBreakOddPage Then
                RangeHasSectionBreak = True
                Exit Function
            End If
        End If
    Next brk
End Function

Sub Delay(Seconds As Double)
    Dim EndTime As Double
    EndTime = Timer + Seconds
    Do While Timer < EndTime
        DoEvents
    Loop
End Sub
Function RemoveLastChar(ByRef tempoRange As Range)
    
    Dim lastChar As String

    Do
        lastChar = Right(tempoRange.text, 1)
        If lastChar <> Chr(9) And lastChar <> Chr(11) And lastChar <> Chr(12) And lastChar <> Chr(13) Then Exit Do
        tempoRange.End = tempoRange.End - 1
    Loop
        
End Function
Sub g()
Debug.Print Selection.Range.Information(wdVerticalPositionRelativeToPage)
End Sub
Sub AddSecBreack(secRange As Range)
    With secRange
        .MoveEnd wdCharacter, 1
        If Helpers.SearchAndReplace("^b", "False", rng:=.Duplicate, searchMatchWildcards:=False) Is Nothing Then
            .MoveEnd wdCharacter, -1
            .Collapse wdCollapseEnd
            .InsertBreak wdSectionBreakContinuous
        End If
    End With
End Sub
-------------------------------------------------------------------------------
