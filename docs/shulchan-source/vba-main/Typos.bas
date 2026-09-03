VBA MACRO Typos.bas 
in file: word/vbaProject.bin - OLE stream: 'VBA/Typos'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit
Option Private Module

Sub Repair(selectedOptions() As Boolean)
    Dim searchText(1 To 14) As String
    Dim replaceText(1 To 14) As String
    Dim msg(1 To 14) As String
    Dim i As Integer
    
    ' רווחים עודפים ופיסקאות ריקות
    If selectedOptions(0) Then
        msg(1) = "מחיקת רווחים עודפים"
        searchText(1) = " {2,}": replaceText(1) = " "
    End If
    If selectedOptions(1) Then
        msg(2) = "מוחק פיסקאות ריקות"
        searchText(2) = "^13{2,}": replaceText(2) = "^13"
    End If
    
    ' רווח לפני תוי פיסוק
    If selectedOptions(2) Then
        msg(3) = "מוחק רווח לפני תוי פיסוק"
        searchText(3) = "( )([.,:\?\!])([! ])": replaceText(3) = "\2\1\3"
        searchText(4) = "( )([.,:\?\!])([ ])": replaceText(4) = "\2\3"
    End If
    
    ' סימני פיסוק כפולים
    If selectedOptions(3) Then
        msg(5) = "מוחק סימני פיסוק כפולים"
        searchText(5) = "([,:\!\?]){2,}": replaceText(5) = "\1"
        searchText(6) = "([!.])(..)([!.])": replaceText(6) = "\1.\3"
    End If

    ' מעל 3 נקודות
    If selectedOptions(4) Then
        msg(7) = "מוחק מעל 3 נקודות"
        searchText(7) = ".{4,}": replaceText(7) = "..."
    End If

    If selectedOptions(5) Then
        msg(8) = "מוחק רווח לפני ואחרי סוגריים"
        ' רווח לפני סוגר סוגריים
        searchText(8) = "( )([\)\]])([! ])": replaceText(8) = "\2\1\3"
        searchText(9) = "( )([\)\]])( )": replaceText(9) = "\2\3"

        ' רווח אחרי פותח סוגריים
        searchText(10) = "([! ])([\(\[])( )": replaceText(10) = "\1\3\2"
        searchText(11) = "( )([\(\[])( )": replaceText(11) = "\1\2"
    End If
    
    ' רווח לפני ואחרי פיסקה
    If selectedOptions(6) Then
        msg(12) = "מוחק רווח לפני ואחרי פיסקה"
        searchText(12) = "^13 ": replaceText(12) = "^13"
        searchText(13) = " ^13": replaceText(13) = "^13"
    End If
    
    ' זוג גרשיים בודדים לגרשיים אחד
    If selectedOptions(7) Then
        msg(14) = "מחליף זוג גרשיים בודדים בגרשיים אחד"
        searchText(14) = "''": replaceText(14) = """"
    End If
    
    On Error GoTo Ending
    Application.UndoRecord.StartCustomRecord "תיקון שגיאות מצויות"
    FormRunning.Show vbModeless
    For i = LBound(searchText) To UBound(searchText)
        If searchText(i) <> "" Then
            If msg(i) <> "" Then FormRunning.Label1 = msg(i): DoEvents
            If stopCode Then GoTo Ending
            Call Helpers.SearchAndReplace(searchText(i), replaceText(i))
        End If
    Next i
    
    ' אות אנגלית אחרי מרכאות
    If selectedOptions(8) Then
        FormRunning.Label1 = "מתקן אות אנגלית אחרי מרכאות": DoEvents
        If stopCode Then GoTo Ending
        Call FixQuotesShiftErrors
    End If

Ending:
    Unload FormRunning
    Application.UndoRecord.EndCustomRecord
End Sub
Sub FixQuotesShiftErrors()
    Dim rng As Range
    Dim badChar As String, goodChar As String
    Dim map As Object
    
    ' טבלת מיפוי אנגלית ? עברית
    Set map = CreateObject("Scripting.Dictionary")
    map.Add "T", "א"
    map.Add "C", "ב"
    map.Add "D", "ג"
    map.Add "S", "ד"
    map.Add "V", "ה"
    map.Add "U", "ו"
    map.Add "Z", "ז"
    map.Add "J", "ח"
    map.Add "Y", "ט"
    map.Add "H", "י"
    map.Add "F", "כ"
    map.Add "K", "ל"
    map.Add "N", "מ"
    map.Add "B", "נ"
    map.Add "X", "ס"
    map.Add "G", "ע"
    map.Add "P", "פ"
    map.Add "M", "צ"
    map.Add "E", "ק"
    map.Add "R", "ר"
    map.Add "A", "ש"
    map.Add ">", "ת"
    map.Add "O", "ם"
    map.Add "I", "ן"
    map.Add "<", "ץ"
    map.Add "L", "ך"
    
    Set rng = ActiveDocument.Content
    With rng.Find
        .ClearFormatting
        .text = "[""]([A-Z<>])"
        .MatchWildcards = True
        Do While .Execute
            DoEvents: If stopCode Then GoTo Ending
            badChar = rng.Characters(2).text
            If map.Exists(badChar) Then
                goodChar = map(badChar)
                rng.Characters(2).text = goodChar
            End If
            rng.Collapse wdCollapseEnd
        Loop
    End With
Ending:
End Sub
Sub FixHebrewPunctuation()
    Dim searchText(1 To 2) As String
    Dim replaceText(1 To 2) As String
    Dim rng As Range
    Dim i As Integer
    
    Set rng = Selection.Range
    searchText(1) = "[ '"",.:\(\)\[\]\?\!\{\}]": replaceText(1) = "^&"
    searchText(2) = "([!^l])(^s)": replaceText(2) = "\1 "

    Application.UndoRecord.StartCustomRecord "תיקון שגיאות העתקה"
    For i = LBound(searchText) To UBound(searchText)
        Call Helpers.SearchAndReplace(searchText(i), replaceText(i), rng:=rng.Duplicate)
    Next i
    Application.UndoRecord.EndCustomRecord
End Sub
-------------------------------------------------------------------------------
