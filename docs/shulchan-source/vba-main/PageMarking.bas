VBA MACRO PageMarking.bas 
in file: word/vbaProject.bin - OLE stream: 'VBA/PageMarking'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit
Option Private Module

Sub Repair()
    Dim doc As Document
    Dim rng As Range
    Dim pageCount As Integer
    Dim i As Integer
    Dim originalPosition As Long
    
    On Error GoTo ErrorHandler
    
    Application.UndoRecord.StartCustomRecord "עיצוב 'לפני שנתחיל'"
    FormRunning1.Show vbModeless
    
    ' קבלת המסמך הפעיל
    Set doc = ActiveDocument
    pageCount = doc.ComputeStatistics(wdStatisticPages)
    
    ' מעבר על כל עמוד במסמך
    For i = 1 To pageCount
        FormRunning1.Label1 = "מסמן מסמך עמוד " & i & " מתוך " & pageCount: DoEvents
        ' הגדרת טווח העמוד
        Set rng = doc.GoTo(what:=wdGoToPage, Which:=wdGoToAbsolute, Count:=i)
        With rng
            .MoveEndUntil " "
            .Font.Color = RGB(1, 255, 1)
            .Collapse wdCollapseStart
            .End = doc.GoTo(what:=wdGoToPage, Which:=wdGoToAbsolute, Count:=i + 1).Start - 1
            .Collapse wdCollapseEnd
            Do
                originalPosition = .Start
                .MoveStart wdCharacter, -1
                If (.Characters.First = " " And Len(.text) > 2) Or .Start = originalPosition Then Exit Do
            Loop
            .Start = originalPosition
            .Font.Color = RGB(255, 1, 1)
        End With
    Next i
    
    Unload FormRunning1
    Application.UndoRecord.EndCustomRecord
    Exit Sub
    
ErrorHandler:
    Unload FormRunning1
    Application.UndoRecord.EndCustomRecord
    MsgBox "אירעה שגיאה: " & Err.Description, vbCritical, "שגיאה"
End Sub
Sub BugSearch()

    Static currentPage As Integer
    Static errorsFound As Boolean
    Dim doc As Document
    Dim rng As Range
    Dim pageCount As Integer
    Dim i As Integer
    Dim firstWord As Range
    Dim startOfPageColor As Long
 
    ' צבע לבדיקה
    startOfPageColor = RGB(1, 255, 1) ' ירוק בהיר
 
    ' אתחול משתנים
    Set doc = ActiveDocument
    pageCount = doc.ComputeStatistics(wdStatisticPages)
 
    ' התחלת בדיקה מהעמוד הראשון אם זה ההפעלה הראשונה
    If currentPage = 0 Then
        currentPage = 1
        errorsFound = False ' איפוס מצב שגיאות
    End If
 
    ' מעבר על עמודים מהעמוד הנוכחי עד סוף המסמך
    For i = currentPage To pageCount
        ' הגדרת טווח עמוד
        Set rng = doc.GoTo(what:=wdGoToPage, Which:=wdGoToAbsolute, Count:=i)
        rng.End = doc.GoTo(what:=wdGoToPage, Which:=wdGoToAbsolute, Count:=i + 1).Start - 1
 
        ' בדיקה אם יש מילים בעמוד
        If rng.Words.Count > 0 Then
            ' קבלת המילה הראשונה בדיוק מתחילת העמוד
            Set firstWord = rng.Words(1)
            If firstWord.Information(wdActiveEndPageNumber) = i Then
                ' בדיקת צבע המילה הראשונה
                If firstWord.Font.Color <> startOfPageColor Then
                    firstWord.Select
                    errorsFound = True ' נמצאה שגיאה
                    currentPage = i + 1 ' שמירת המיקום להמשך החיפוש
                    Exit Sub
                End If
            End If
        End If
    Next i
 
    ' אם הגענו לסוף המסמך
    If errorsFound Then
        MsgBox "החיפוש הסתיים, לא נמצאו עמודים נוספים שהשתנו", vbInformation
    Else
        MsgBox "החיפוש הסתיים ולא נמצאו עמודים שהשתנו", vbInformation
    End If
 
    ' איפוס המיקום והסטטוס לבדיקות חדשות
    currentPage = 0
    errorsFound = False
End Sub
Sub Remove()
    On Error GoTo ErrorHandler ' הפעלת טיפול בשגיאות
    Dim doc As Document
    Dim rng As Range
    Dim i As Integer
    Dim word As Range
    
    ' קבלת המסמך הפעיל
    Set doc = ActiveDocument
    Application.UndoRecord.StartCustomRecord "הסרת עיצוב 'לפני שנתחיל'"
    ' מעבר על כל המילים במסמך
    For Each rng In doc.StoryRanges
        Do
            For Each word In rng.Words
                ' אם הצבע אדום בהיר או ירוק בהיר, נסיר אותו
                If word.Font.Color = RGB(1, 255, 1) Or word.Font.Color = RGB(255, 1, 1) Then
                    word.Font.Color = wdColorAutomatic
                End If
            Next word
            Set rng = rng.NextStoryRange
        Loop While Not rng Is Nothing
    Next rng
    
    Application.UndoRecord.EndCustomRecord
    MsgBox "הצבעים הוסרו בהצלחה!", vbInformation
    Exit Sub
    
ErrorHandler:
    Application.UndoRecord.EndCustomRecord
    MsgBox "אירעה שגיאה: " & Err.Description, vbCritical, "שגיאה"
End Sub
-------------------------------------------------------------------------------
