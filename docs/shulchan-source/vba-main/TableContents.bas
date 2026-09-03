VBA MACRO TableContents.bas 
in file: word/vbaProject.bin - OLE stream: 'VBA/TableContents'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit
Option Private Module

Dim userTest As Boolean
Dim styleTest(1 To 5) As Boolean
Dim styleSelectName(1 To 5) As String
Dim boundaryBefore(1 To 5) As String
Dim tub(1 To 5) As Integer
Dim continuous(1 To 5) As Boolean
Dim numbering(1 To 5) As Integer
Dim boundary(1 To 5) As String

Sub ReadingValues()
    
    Dim i As Integer
    
    userTest = SettingsHelper.GetSavedSetting(appName, "tableContents", "userTest", False)
    
    For i = 1 To 5
        styleTest(i) = SettingsHelper.GetSavedSetting(appName, "tableContents", "styleTest" & i, False)
        styleSelectName(i) = SettingsHelper.GetSavedSetting(appName, "tableContents", "styleSelectName" & i, "")
        boundaryBefore(i) = SettingsHelper.GetSavedSetting(appName, "tableContents", "boundaryBefore" & i, "")
        tub(i) = SettingsHelper.GetSavedSetting(appName, "tableContents", "tub" & i, 0)
        continuous(i) = SettingsHelper.GetSavedSetting(appName, "tableContents", "continuous" & i, False)
        numbering(i) = SettingsHelper.GetSavedSetting(appName, "tableContents", "numbering" & i, 0)
        boundary(i) = SettingsHelper.GetSavedSetting(appName, "tableContents", "boundary" & i, "")
    Next i

End Sub

Function GetBmkName() As String
    GetBmkName = "_TableContents"
End Function
Function GetStlName() As String
    GetStlName = "תוכן עניינים"
End Function

Sub Creating()
    
    Dim doc As Document
    Dim styleName As New Collection
    Dim rng As Range
    Dim tocStartPos As Integer
    Dim tocSecStartPos As Integer
    Dim para As Paragraph
    Dim paraRange As Range
    Dim titleStyle() As Integer
    Dim firstTitleCode As Integer
    Dim x, i As Integer
    Dim fld As Field
    Dim bb As BuildingBlock
    Dim bbRng As Range
    
    On Error GoTo Ending
    Application.UndoRecord.StartCustomRecord "יצירת תוכן עניינים"
    Application.ScreenUpdating = False
    
    Selection.Collapse wdCollapseStart
    Set rng = Selection.Range
    
    Call ReadingValues
    If Not userTest Then GoTo Ending
    
    FormRunning.Show vbModeless
    Set doc = ActiveDocument
    firstTitleCode = SearchLastTitleCode
        
    ' יצירת סגנונות
    Set styleName = CreateStyles
    x = firstTitleCode
    
    ' מעבר על הפיסקאות והכנסת הערכים למערך
    For Each para In ActiveDocument.Paragraphs
        FormRunning.Label1.Caption = "סורק את המסמך": DoEvents
        Set paraRange = para.Range
        paraRange.End = paraRange.End - 1
        If ((para.Style.NameLocal = styleSelectName(1) And styleTest(1)) Or _
           (para.Style.NameLocal = styleSelectName(2) And styleTest(2)) Or _
           (para.Style.NameLocal = styleSelectName(3) And styleTest(3)) Or _
           (para.Style.NameLocal = styleSelectName(4) And styleTest(4)) Or _
           (para.Style.NameLocal = styleSelectName(5) And styleTest(5))) And _
           Len(para.Range.text) > 3 _
           Then
            
            x = x + 1
            
            ReDim Preserve titleStyle(firstTitleCode + 1 To x)
            If para.Style.NameLocal = styleSelectName(1) Then
                titleStyle(x) = 1
            ElseIf para.Style.NameLocal = styleSelectName(2) Then
                titleStyle(x) = 2
            ElseIf para.Style.NameLocal = styleSelectName(3) Then
                titleStyle(x) = 3
            ElseIf para.Style.NameLocal = styleSelectName(4) Then
                titleStyle(x) = 4
            ElseIf para.Style.NameLocal = styleSelectName(5) Then
                titleStyle(x) = 5
            End If
            Application.ScreenRefresh
            doc.Bookmarks.Add GetBmkName & x, paraRange
        End If
    Next para
    
    If x = 0 Then GoTo Ending
    
    With rng
        tocStartPos = .Start ' + 1
        .InsertAfter " "
        .Collapse wdCollapseStart
        .Select
        Selection.ClearFormatting
        
        ' מעבר על המערך והכנסת הנתונים למסמך
        For i = firstTitleCode + 1 To x
            FormRunning.Label1.Caption = "יוצר תוכן עניינים " & i - firstTitleCode & " מתוך " & x - firstTitleCode: DoEvents
            
            ' בדיקה אם יש שורה חדשה לפני
            If continuous(titleStyle(i)) And i > firstTitleCode + 1 And Not .Characters.Last = Chr(13) Then
                .InsertAfter Chr(13)
            End If
            
            ' שמירת תחילת הטווח הנוכחי
            .Collapse wdCollapseEnd
            tocSecStartPos = .Start
            
            ' הוספת שדה של טקסט כותרת
            .Collapse wdCollapseEnd
            Set fld = .Fields.Add(Range:=rng, Type:=wdFieldRef, text:=GetBmkName & i, PreserveFormatting:=True)
            .End = fld.Result.End + 1
            .Collapse wdCollapseEnd
            
            ' הוספת טאבים ותוחמים
            If tub(titleStyle(i)) > 0 Then
                .InsertAfter Chr(9) ' wdAlignTabLeft, tub(titleStyle(i))   ' - 1
            Else
                .InsertAfter " "
            End If
            If boundaryBefore(titleStyle(i)) <> "" Then .InsertAfter boundaryBefore(titleStyle(i)) & " "
            
            ' הוספת שדה של מספר עמוד
            .Collapse wdCollapseEnd
            Select Case numbering(titleStyle(i))
                Case 1
                    Set fld = .Fields.Add(Range:=rng, Type:=wdFieldPageRef, text:=GetBmkName & i, PreserveFormatting:=False)
                Case 2
                    Set fld = .Fields.Add(Range:=rng, Type:=wdFieldPageRef, text:=GetBmkName & i & "\* hebrew1", PreserveFormatting:=False)
                Case 3
                    Set bb = ThisDocument.AttachedTemplate.BuildingBlockEntries("HebrewNumbers")
                Case 4
                    Set bb = ThisDocument.AttachedTemplate.BuildingBlockEntries("HebrewNumbersClean")
                Case 5
                    Set bb = ThisDocument.AttachedTemplate.BuildingBlockEntries("HebrewNumbersShortK")
            End Select
            
            If numbering(titleStyle(i)) > 2 Then
                Set bbRng = bb.Insert(where:=rng)
                bbRng.Select
                Call ConvertToBookmarkRef(bbRng, GetBmkName & i)
                .End = bbRng.End
            ElseIf numbering(titleStyle(i)) > 0 Then
                .End = fld.Result.End + 1
            End If
            
            ' הוספת תוחם ושורה חדשה
            If i < x Then If titleStyle(i) = titleStyle(i + 1) Then .InsertAfter " " & boundary(titleStyle(i)) & " "
            
            ' עיצוב הטקסט והטאבים
            .Start = tocSecStartPos
            .Select
            Selection.ClearFormatting
            .Style = styleName(titleStyle(i))
            .ParagraphStyle = styleName(titleStyle(i))
            If continuous(titleStyle(i)) Then .InsertAfter Chr(13) Else .InsertAfter " "
        Next i
        
        .Start = tocStartPos
        Helpers.SearchAndReplace " {2,}", " ", .Duplicate
        Helpers.SearchAndReplace "( )([" & Chr(13) & Chr(9) & "])", "\2", .Duplicate
        Helpers.SearchAndReplace "([" & Chr(13) & Chr(9) & "])( )", "\1", .Duplicate
        Set rng = Helpers.RemoveLastCharInRng(rng)
        doc.Bookmarks.Add GetBmkName & firstTitleCode + 1 & "_" & x, rng
    End With
    
Ending:
    Call Update
    Application.UndoRecord.EndCustomRecord
    Application.ScreenUpdating = True
    rng.Select
    Unload FormRunning
End Sub
Sub Update()

    Dim doc As Document
    Dim bmk As Bookmark
    Dim bmkName As String
    Dim bmkCode As String
    Dim bmkRng As Range
    Dim fld As Field
    Dim i As Integer
    
    On Error GoTo Ending
    Application.UndoRecord.StartCustomRecord "עדכון תוכן עניינים"
    
    Set doc = ActiveDocument
    
    ' הרחבת הטווח של הסימניות עד סוף הפיסקה
    For i = doc.Bookmarks.Count To 1 Step -1
        
        Set bmk = doc.Bookmarks(i)
        
        With bmk
            bmkCode = Replace(.name, GetBmkName, "")
            If .name Like GetBmkName & "*" And IsNumeric(bmkCode) Then
                bmkName = .name
                Set bmkRng = .Range
                bmkRng.End = bmkRng.Paragraphs(1).Range.End
                Set bmkRng = RemoveLastCharInRng(bmkRng.Duplicate)
                .Delete
                doc.Bookmarks.Add bmkName, bmkRng
            End If
        End With
    Next i
    
    ' עדכון תוכן העניינים
    For Each bmk In doc.Bookmarks
        
        With bmk
            bmkCode = Replace(.name, GetBmkName, "")
            If .name Like GetBmkName & "*" And Not IsNumeric(bmkCode) Then
                Set bmkRng = .Range
                bmkRng.Select
                Selection.Fields.Update
            End If
        End With

    Next bmk

Ending:
    Application.UndoRecord.EndCustomRecord
End Sub
Sub Remove()
    Dim doc As Document
    Dim bmk As Bookmark
    Dim bmkName As String
    Dim bmkCode As String
    Dim bmkRng As Range
    Dim fld As Field
    Dim i As Integer
    Dim firstCode As Integer
    Dim lastCode As Integer
    
    On Error GoTo Ending
    Application.UndoRecord.StartCustomRecord "מחיקת תוכן עניינים"
    
    Set doc = ActiveDocument
    
    ' הרחבת הטווח של הסימניות עד סוף הפיסקה
    For i = doc.Bookmarks.Count To 1 Step -1
        
        Set bmk = doc.Bookmarks(i)
        
        With bmk
            bmkCode = Replace(.name, GetBmkName, "")
            If .name Like GetBmkName & "*" And Not _
               IsNumeric(bmkCode) And _
               Selection.Start >= .Range.Start And .Range.End <= .End Then
                
                firstCode = Int(Split(bmkCode, "_")(0))
                lastCode = Int(Split(bmkCode, "_")(1))
                .Range.Delete
                
            End If
        End With
    Next i
    
    ' עדכון תוכן העניינים
    For Each bmk In doc.Bookmarks
        
        With bmk
            bmkCode = Replace(.name, GetBmkName, "")
            If .name Like GetBmkName & "*" And IsNumeric(bmkCode) Then
                If bmkCode >= firstCode And bmkCode <= lastCode Then .Delete
            End If
        End With

    Next bmk

Ending:
    Application.UndoRecord.EndCustomRecord
    
End Sub
Function CreateStyles() As Collection

    Dim doc As Document
    Dim rng As Range
    Dim stl As Style
    Dim splitName() As String
    Dim highNum As Integer
    Dim stlCollection As New Collection
    Dim i As Integer
    
    Set doc = ActiveDocument
    Set rng = Selection.Range
    
    rng.Select
    Selection.ClearFormatting
    Selection.Style = "רגיל"
    
    ' מציאת המספר האחרון של תוכן העניינים
    For Each stl In ActiveDocument.Styles
        If stl.NameLocal Like GetStlName & "*" Then
            splitName = Split(stl.NameLocal, ">")
            If CInt(splitName(1)) > highNum Then
                highNum = CInt(splitName(1))
            End If
        End If
    Next stl
    
    highNum = highNum + 1
    
    ' יצירת סגנונות
    For i = 1 To 5
        Set stl = doc.Styles.Add(name:=GetStlName & ">" & highNum & ">" & "סגנון" & i, Type:=wdStyleTypeLinked)
        If tub(i) Then
            With stl.ParagraphFormat
'                .RightIndent = CentimetersToPoints(entry(i)) / 10
'                .LeftIndent = CentimetersToPoints(entry(i)) / 10
                .TabStops.Add Helpers.AligmentTub(rng, stl:=stl), wdAlignTabRight, tub(i) - 1
            End With
        End If
        stlCollection.Add stl
    Next i
    
    Set CreateStyles = stlCollection

End Function
Function SearchLastTitleCode() As Integer
    
    Dim doc As Document
    Dim bmk As Bookmark
    Dim bmkCode As String
    
    Set doc = ActiveDocument
    
    For Each bmk In doc.Bookmarks
        With bmk
            bmkCode = Replace(.name, GetBmkName, "")
            If .name Like GetBmkName & "*" And IsNumeric(bmkCode) Then
                If bmkCode > SearchLastTitleCode Then
                    SearchLastTitleCode = bmkCode
                End If
            End If
        End With
    Next bmk

End Function
Sub InsertSpecialAlignmentTab()

    ' קובע את הטווח שבו יווסף הטאב.
    ' לדוגמה, במקום הסמן הנוכחי.
    Dim myRange As Range
    Set myRange = Selection.Range

'    ' מגדיר את סוג היישור
'    Dim alignType As WdAlignmentTabAlignment
'    alignType = wdAlignTabCenter ' יישור למרכז
'    ' אפשרויות נוספות:
'    ' wdAlignTabLeft ' יישור לשמאל
'    ' wdAlignTabRight ' יישור לימין
'
'    ' מגדיר למה ליישר יחסית (שוליים או עמוד)
'    Dim relativeToType As WdAlignmentTabRelative
'    relativeToType = wdAlignTabRelativeToPage ' יחסית לעמוד
'    ' אפשרויות נוספות:
'    ' wdAlignTabRelativeToMargin ' יחסית לשוליים
'
'    ' מגדיר את התווים המובילים
'    Dim leaderType As WdTabLeader
'    leaderType = wdTabLeaderDots ' נקודות
'    ' אפשרויות נוספות:
'    ' wdTabLeaderSpaces ' רווחים (ברירת מחדל)
'    ' wdTabLeaderDashes ' קווים מפרידים
'    ' wdTabLeaderLines ' קו תחתון

    ' מוסיף את הטאב היישור המיוחד
    myRange.InsertAlignmentTab wdAlignTabLeft


End Sub

Function ConvertToBookmarkRef(rng As Range, bmkName As String)
    Dim fld As Field
    For Each fld In rng.Fields
        With fld.Code.Find
            .text = "page"
            .Replacement.text = "PAGEREF " & bmkName
            .MatchWholeWord = True
            .Execute Replace:=wdReplaceAll
        End With
        fld.Update
    Next fld
End Function
-------------------------------------------------------------------------------
