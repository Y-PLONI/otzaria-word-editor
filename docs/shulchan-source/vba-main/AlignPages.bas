VBA MACRO AlignPages.bas 
in file: word/vbaProject.bin - OLE stream: 'VBA/AlignPages'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit
Option Private Module

Sub Repair()
    
    Dim selectionRange As Range
    Dim i, x, y As Integer
    Dim allPagesRange As New Collection
    Dim pageRange As Range
    
    On Error GoTo Ending
    Application.UndoRecord.StartCustomRecord "יישור עמודים"
    Application.ScreenUpdating = False
    ActiveDocument.Repaginate
    
    Set selectionRange = Selection.Range
    
    FormRunning.Show vbModeless
    FormRunning.Label1.Caption = "אוסף מידע עבור יישור עמודים": DoEvents
    If stopCode Then GoTo Ending
    
    Set allPagesRange = GetPages(selectionRange)
    
    For Each pageRange In allPagesRange
        
        y = y + 1
        FormRunning.Label1.Caption = "מיישר עמודים " & y & " מתוך " & allPagesRange.Count: DoEvents
        If stopCode Then GoTo Ending

        ' פיסקה אחרונה מוגדרת 0
        If Helpers.ChekLastChars(pageRange.Duplicate, 3, Chr(13)) Then
            pageRange.Paragraphs.Last.SpaceAfter = 0
        End If
        
        Call IncreaseSpacingBetweenParagraphs(pageRange, 1.1)
        Call IncreaseSpacingBetweenParagraphs(pageRange, 1.05)
        Call IncreaseSpacingBetweenParagraphs(pageRange, 1.005)
        
    Next pageRange

Ending:
    Application.UndoRecord.EndCustomRecord
    Application.ScreenUpdating = True
    selectionRange.Select
    Unload FormRunning
End Sub
Function GetPages(rng As Range) As Collection

    Dim pages As New Collection
    Dim firstPage As Integer
    Dim lastPage As Integer
    Dim docPagesCount As Integer
    Dim pageRange As Range
    Dim i As Integer
    
    ' קבלת העמוד הראשון והאחרון בטווח
    With rng.Characters
        firstPage = .First.Information(wdActiveEndPageNumber)
        lastPage = .Last.Information(wdActiveEndPageNumber)
    End With
    
    ' קבלת מספר העמודים במסמך
    docPagesCount = ActiveDocument.ComputeStatistics(wdStatisticPages)
    
    ' מעבר על העמודים ובדיקה אם מכילים מעבר עמוד
    For i = firstPage To lastPage
        
        Selection.GoTo what:=wdGoToPage, Which:=wdGoToAbsolute, Count:=i
        Set pageRange = Selection.Bookmarks("\Page").Range
        
        If (Helpers.SearchAndReplace("^b", "False", rng:=pageRange.Duplicate, searchMatchWildcards:=False) Is Nothing And _
            i < docPagesCount) Or _
            lastPage - firstPage = 0 _
            Then
                ' אם הטווח מכיל 2 טורים ואין מעבר מקטע הטווח מצטמצם לטור השני
                If Helpers.SearchAndReplace("^m", "False", rng:=pageRange.Duplicate, searchMatchWildcards:=False) Is Nothing And _
                    pageRange.Sections(1).PageSetup.TextColumns.Count = 2 _
                    Then
                        pageRange.Start = Helpers.GetColumnRange(pageRange.Duplicate).End
                End If
                pages.Add pageRange
        End If
    Next i
    Set GetPages = pages
End Function
Function IncreaseSpacingBetweenParagraphs(pageRange As Range, num As Double)

    Dim pageNum As Integer
    Dim parasCount As Integer
    Dim orginalParaSpaceAfter() As Single
    Dim orginalParaSpaceBefore() As Single
    Dim footnoteRange As Range
    Dim footnotePosition As Double
    Dim para As Paragraph
    Dim i, x As Integer
    
    parasCount = pageRange.Paragraphs.Count
    
    ReDim orginalParaSpaceAfter(parasCount)
    ReDim orginalParaSpaceBefore(parasCount)
    
    pageNum = pageRange.Characters.Last.Information(wdActiveEndPageNumber)
    
    ' בדיקת מיקום הערות שוליים
    If pageRange.Footnotes.Count > 0 Then
        Set footnoteRange = pageRange.Footnotes(1).Range
        footnotePosition = footnoteRange.Information(wdVerticalPositionRelativeToPage)
    End If
        
    For x = 1 To 20
        
        For i = 1 To parasCount
            
            Set para = pageRange.Paragraphs(i)
            
            With para
                orginalParaSpaceAfter(i) = .SpaceAfter
                orginalParaSpaceBefore(i) = .SpaceBefore
                If i < parasCount Then .SpaceAfter = .SpaceAfter * num
                If i > 1 Then .SpaceBefore = .SpaceBefore * num
            End With
        
        Next i
    
        If pageRange.Characters.Last.Information(wdActiveEndPageNumber) > pageNum Then Exit For
        If Not footnoteRange Is Nothing Then If footnotePosition <> footnoteRange.Information(wdVerticalPositionRelativeToPage) Then Exit For
        
    Next x
    
    For i = 1 To parasCount
        
        Set para = pageRange.Paragraphs(i)
        
        With para
            .SpaceAfter = orginalParaSpaceAfter(i)
            .SpaceBefore = orginalParaSpaceBefore(i)
        End With
    
    Next i

End Function
'Function getFootnotesRange(pageRange As Range) As Range
'
'    Dim pageNum As Integer
'    Dim originalEnd As Integer
'    Dim RangePosition As Double
'
'    pageNum = pageRange.Information(wdActiveEndPageNumber)
'
'    If pageRange.Footnotes.Count > 0 Then
'        Set getFootnotesRange = pageRange.Footnotes(1).Range
'        With getFootnotesRange
'            .Collapse wdCollapseStart
'            Do
'                originalEnd = .End
'                RangePosition = .Characters.Last.Information(wdVerticalPositionRelativeToPage)
'                .MoveEnd wdWord, 1
'                If RangePosition > .Characters.Last.Information(wdVerticalPositionRelativeToPage) Then Exit Do
'                If originalEnd = .End Then Exit Do
'                If .Characters.Last.Information(wdActiveEndAdjustedPageNumber) > pageNum Then Exit Do
'            Loop
'            .End = originalEnd
'        End With
'    End If
'End Function
Sub ad()
-------------------------------------------------------------------------------
