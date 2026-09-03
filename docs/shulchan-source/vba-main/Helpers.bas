VBA MACRO Helpers.bas 
in file: word/vbaProject.bin - OLE stream: 'VBA/Helpers'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit
Option Private Module
Function GetPageHeight(rng As Range) As Double

    With rng.Sections(1).PageSetup
        GetPageHeight = .PageHeight - .BottomMargin - .TopMargin
    End With
    
End Function
Function GetPageWidth(rng As Range) As Double

    With rng.Sections(1).PageSetup
        GetPageWidth = .PageWidth - .LeftMargin - .RightMargin
    End With
    
End Function
Function SearchAndReplace(searchText As String, Optional replaceText As String = "False", Optional rng As Range = Nothing, Optional searchForward As Boolean = True, Optional searchMatchWildcards = True) As Range
    
    If rng Is Nothing Then Set rng = ActiveDocument.Range
'    On Error Resume Next
'    If IsMissing(replaceText) Then replaceText = "False"
'    On Error GoTo 0
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .text = searchText
        .Replacement.text = replaceText
        .Forward = searchForward
        .Wrap = wdFindStop
        .Format = True
        .MatchKashida = True
        .MatchWildcards = searchMatchWildcards
        If replaceText = "False" Then
            If .Execute Then Set SearchAndReplace = rng.Duplicate
        Else
            .Execute Replace:=wdReplaceAll
        End If
    End With

End Function
Function ChekLastChars(rng As Range, cest As Integer, char As String) As Boolean
    
    Dim i As Integer
    
    With rng.Characters
        
        If .Count < cest Then cest = .Count
    
        For i = 0 To cest - 1
            If rng.Characters(.Count - i).text = char Then
                ChekLastChars = True
                Exit Function
            End If
        Next i
    End With
End Function
Function GetColumnRange(pageRange As Range) As Range
    
    Dim pos As Double
    Dim startPos As Long
    Dim endPos As Long
    
    Application.ScreenUpdating = False
    
    With pageRange
        startPos = .Start
        .Collapse wdCollapseStart
        .Select
    End With
    
    With Selection
        .EndKey wdLine, Extend:=False
        Do
            pos = .Information(wdVerticalPositionRelativeToPage)
            endPos = .End
            .Move wdLine, 1
            If endPos = .End Or pos > .Information(wdVerticalPositionRelativeToPage) Then Exit Do
        Loop
    End With
    
    Set GetColumnRange = ActiveDocument.Range(Start:=startPos, End:=endPos)
    Application.ScreenUpdating = True
End Function

Function CheckDocsSaved()

    Dim doc As Document
    Dim dontSaved As Boolean
    Dim check As Boolean
    
    check = SettingsHelper.GetSavedSetting(appName, "DocsSaved", "dontShowAgain", False)
    
    If check Then
        
        Exit Function
    
    End If

    For Each doc In Application.Documents
        
        If Not doc.Saved Then
            
            dontSaved = True
            Exit For
            
        End If
    
    Next doc
    
    If dontSaved Then FormSaver.Show
    
End Function
Function TextToNumIsInRange(text As String, min As Double, max As Double, Optional ByRef numRef As Double, Optional msg As Boolean = False, Optional charsToRemove As String = "") As Boolean
    
    ' הסרת תוים מיותרים
    If charsToRemove <> "" Then text = ReplaceInString(text, charsToRemove)
    
    ' בדיקה אם מספר
    If Not IsNumeric(text) Then
        If msg Then MsgBox "אנא הכנס מספר חוקי!", vbOKOnly, "שגיאה"
        DoEvents
        Exit Function
    End If
    
    ' בדיקה אם בתוך הטווח
    If CDbl(text) < min Or CDbl(text) > max Then
        If msg Then MsgBox "נא הכנס מספר בין " & min & " ל- " & max, vbOKOnly, "שגיאה"
        DoEvents
        Exit Function
    End If
       
    ' החזרת הערכים
    numRef = CDbl(text)
    TextToNumIsInRange = True

End Function
Function ReplaceInString(text As String, charsToRemove As String) As String
    
    Dim i As Integer
    
    ReplaceInString = text
    
    For i = 1 To Len(charsToRemove)
        ReplaceInString = Replace(ReplaceInString, Right(Left(charsToRemove, i), 1), "")
    Next i

End Function

Function ShowPicDialog(dialogTitle As String) As String

    Dim dlg As FileDialog
    Dim pic As StdPicture

    Set dlg = Application.FileDialog(msoFileDialogFilePicker)
    
    With dlg
        
        .Title = dialogTitle
        .Filters.Clear
        .Filters.Add "תמונות", "*.jpg; *.jpeg; *.png; *.bmp; *.gif"
        If .Show <> -1 Then Exit Function
        ShowPicDialog = .SelectedItems(1)
    
    End With


End Function
Function RemoveLastCharInRng(tempoRange As Range) As Range
    
    Dim lastChar As String
    Do
        lastChar = Right(tempoRange.text, 1)
        If lastChar <> " " And lastChar <> Chr(9) And lastChar <> Chr(11) And lastChar <> Chr(12) And lastChar <> Chr(13) Then Exit Do
        tempoRange.End = tempoRange.End - 1
'        tempoRange.Select
    Loop
    Set RemoveLastCharInRng = tempoRange
End Function
Function AligmentTub(rng As Range, Optional stl As Style = Nothing) As Double
    With rng.PageSetup
        AligmentTub = .PageWidth - .LeftMargin - .RightMargin
    End With
    If Not stl Is Nothing Then
        AligmentTub = AligmentTub - stl.ParagraphFormat.leftIndent
    End If
End Function
Function GetLineHeight(rng As Range) As Double
    Dim prevSpaceBefore As Double
    Dim prevWidowControl As Boolean
    
    With rng
        If .Paragraphs(1).LineSpacingRule = wdLineSpaceExactly Then
            GetLineHeight = .Paragraphs(1).LineSpacing
            Exit Function
        End If
        .Expand wdParagraph
        .Collapse wdCollapseStart
        prevSpaceBefore = .Paragraphs(1).SpaceBefore
        prevWidowControl = .ParagraphFormat.WidowControl
        .Paragraphs(1).SpaceBefore = 0
        .ParagraphFormat.WidowControl = True
        .InsertAfter "א" & Chr(11)
        .Collapse wdCollapseEnd
        GetLineHeight = .Information(wdVerticalPositionRelativeToTextBoundary)
        .MoveStart wdCharacter, -2
        .Delete
        .Paragraphs(1).SpaceBefore = prevSpaceBefore
        .ParagraphFormat.WidowControl = prevWidowControl
    End With
End Function
Sub CopyFontFormat(srcRange As Range, destRange As Range)
    With destRange.Font
        ' מאפייני רגיל
        .name = srcRange.Font.name
        .Size = srcRange.Font.Size
        .Bold = srcRange.Font.Bold
        .Italic = srcRange.Font.Italic
        .Underline = srcRange.Font.Underline
        .StrikeThrough = srcRange.Font.StrikeThrough
        .Subscript = srcRange.Font.Subscript
        .Superscript = srcRange.Font.Superscript
        .Color = srcRange.Font.Color
        .Shadow = srcRange.Font.Shadow
        .Outline = srcRange.Font.Outline
        .Emboss = srcRange.Font.Emboss
        .Hidden = srcRange.Font.Hidden
        .SmallCaps = srcRange.Font.SmallCaps
        .AllCaps = srcRange.Font.AllCaps
        .Engrave = srcRange.Font.Engrave
        .Animation = srcRange.Font.Animation
        .Kerning = srcRange.Font.Kerning
        .Scaling = srcRange.Font.Scaling
        .Spacing = srcRange.Font.Spacing
        
        ' מאפייני BI לטקסט עברי/דו-כיווני
        .NameBi = srcRange.Font.NameBi
        .SizeBi = srcRange.Font.SizeBi
        .BoldBi = srcRange.Font.BoldBi
        .ItalicBi = srcRange.Font.ItalicBi
    End With
End Sub

Sub RemoveItemFromCollection(ByRef col As Collection, ByVal itemValue As Variant)
    Dim itemInCol As Variant
    Dim i As Integer
    
    For Each itemInCol In col
        i = i + 1
        If itemInCol = itemValue Then
            col.Remove i
        End If
    Next itemInCol
    
End Sub

-------------------------------------------------------------------------------
