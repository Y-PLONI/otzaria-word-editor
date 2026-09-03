VBA MACRO FormTableContents.frm 
in file: word/vbaProject.bin - OLE stream: 'VBA/FormTableContents'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit

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
Sub WritingValues()
    
    Dim i As Integer
    
    SettingsHelper.Save appName, "tableContents", "userTest", userTest
    
    For i = 1 To 5
        SettingsHelper.Save appName, "tableContents", "styleTest" & i, styleTest(i)
        SettingsHelper.Save appName, "tableContents", "styleSelectName" & i, styleSelectName(i)
        SettingsHelper.Save appName, "tableContents", "boundaryBefore" & i, boundaryBefore(i)
        SettingsHelper.Save appName, "tableContents", "tub" & i, tub(i)
        SettingsHelper.Save appName, "tableContents", "continuous" & i, continuous(i)
        SettingsHelper.Save appName, "tableContents", "numbering" & i, numbering(i)
        SettingsHelper.Save appName, "tableContents", "boundary" & i, boundary(i)
    Next i

End Sub

Private Sub UserForm_Initialize()
    
    Dim stl As Style
    
    Call ReadingValues
    userTest = False
    
    ' טעינת תיבת בחירה
    StyleCheck1.value = styleTest(1)
    StyleCheck2.value = styleTest(2)
    StyleCheck3.value = styleTest(3)
    StyleCheck4.value = styleTest(4)
    StyleCheck5.value = styleTest(5)
    
    ' טעינת סגנונות
    StyleBox1.AddItem "בחר סגנון"
    StyleBox2.AddItem "בחר סגנון"
    StyleBox3.AddItem "בחר סגנון"
    StyleBox4.AddItem "בחר סגנון"
    StyleBox5.AddItem "בחר סגנון"
    StyleBox1.ListIndex = 0
    StyleBox2.ListIndex = 0
    StyleBox3.ListIndex = 0
    StyleBox4.ListIndex = 0
    StyleBox5.ListIndex = 0
    
    If styleSelectName(1) <> "" Then StyleBox1.value = styleSelectName(1)
    If styleSelectName(2) <> "" Then StyleBox2.value = styleSelectName(2)
    If styleSelectName(3) <> "" Then StyleBox3.value = styleSelectName(3)
    If styleSelectName(4) <> "" Then StyleBox4.value = styleSelectName(4)
    If styleSelectName(5) <> "" Then StyleBox5.value = styleSelectName(5)
    
    For Each stl In ActiveDocument.Styles
        If stl.Type = wdStyleTypeParagraph Then
            StyleBox1.AddItem stl.NameLocal
            StyleBox2.AddItem stl.NameLocal
            StyleBox3.AddItem stl.NameLocal
            StyleBox4.AddItem stl.NameLocal
            StyleBox5.AddItem stl.NameLocal
        End If
    Next stl
    
    ' טעינת סוג
    ContinuousBox1.AddItem "רציף"
    ContinuousBox1.AddItem "שורה חדשה"
    ContinuousBox2.AddItem "רציף"
    ContinuousBox2.AddItem "שורה חדשה"
    ContinuousBox3.AddItem "רציף"
    ContinuousBox3.AddItem "שורה חדשה"
    ContinuousBox4.AddItem "רציף"
    ContinuousBox4.AddItem "שורה חדשה"
    ContinuousBox5.AddItem "רציף"
    ContinuousBox5.AddItem "שורה חדשה"
    ContinuousBox1.ListIndex = 0
    ContinuousBox2.ListIndex = 0
    ContinuousBox3.ListIndex = 0
    ContinuousBox4.ListIndex = 0
    ContinuousBox5.ListIndex = 0
    
    If continuous(1) Then ContinuousBox1.ListIndex = 1
    If continuous(2) Then ContinuousBox2.ListIndex = 1
    If continuous(3) Then ContinuousBox3.ListIndex = 1
    If continuous(4) Then ContinuousBox4.ListIndex = 1
    If continuous(5) Then ContinuousBox5.ListIndex = 1
    
    ' טעינת תיבות טאב
    TubBox1.AddItem "ללא"
    TubBox1.AddItem "רגיל"
    TubBox1.AddItem "....."
    TubBox2.AddItem "ללא"
    TubBox2.AddItem "רגיל"
    TubBox2.AddItem "....."
    TubBox3.AddItem "ללא"
    TubBox3.AddItem "רגיל"
    TubBox3.AddItem "....."
    TubBox4.AddItem "ללא"
    TubBox4.AddItem "רגיל"
    TubBox4.AddItem "....."
    TubBox5.AddItem "ללא"
    TubBox5.AddItem "רגיל"
    TubBox5.AddItem "....."
    TubBox1.ListIndex = 0
    TubBox2.ListIndex = 0
    TubBox3.ListIndex = 0
    TubBox4.ListIndex = 0
    TubBox5.ListIndex = 0
    If tub(1) > 0 Then TubBox1.ListIndex = tub(1)
    If tub(2) > 0 Then TubBox2.ListIndex = tub(2)
    If tub(3) > 0 Then TubBox3.ListIndex = tub(3)
    If tub(4) > 0 Then TubBox4.ListIndex = tub(4)
    If tub(5) > 0 Then TubBox5.ListIndex = tub(5)
    
    ' טעינת תיבות מספור
    NumberingBox1.AddItem "ללא מספור"
    NumberingBox1.AddItem "מספור רגיל"
'    NumberingBox1.AddItem "עברי מובנה"
'    NumberingBox1.AddItem "עברי מעל שצג"
'    NumberingBox1.AddItem "עברי לשון נקיה"
'    NumberingBox1.AddItem "עברי צמצום אלפים"
    NumberingBox2.AddItem "ללא מספור"
    NumberingBox2.AddItem "מספור רגיל"
'    NumberingBox2.AddItem "עברי מובנה"
'    NumberingBox2.AddItem "עברי מעל שצג"
'    NumberingBox2.AddItem "עברי לשון נקיה"
'    NumberingBox2.AddItem "עברי צמצום אלפים"
    NumberingBox3.AddItem "ללא מספור"
    NumberingBox3.AddItem "מספור רגיל"
'    NumberingBox3.AddItem "עברי מובנה"
'    NumberingBox3.AddItem "עברי מעל שצג"
'    NumberingBox3.AddItem "עברי לשון נקיה"
'    NumberingBox3.AddItem "עברי צמצום אלפים"
    NumberingBox4.AddItem "ללא מספור"
    NumberingBox4.AddItem "מספור רגיל"
'    NumberingBox4.AddItem "עברי מובנה"
'    NumberingBox4.AddItem "עברי מעל שצג"
'    NumberingBox4.AddItem "עברי לשון נקיה"
'    NumberingBox4.AddItem "עברי צמצום אלפים"
    NumberingBox5.AddItem "ללא מספור"
    NumberingBox5.AddItem "מספור רגיל"
'    NumberingBox5.AddItem "עברי מובנה"
'    NumberingBox5.AddItem "עברי מעל שצג"
'    NumberingBox5.AddItem "עברי לשון נקיה"
'    NumberingBox5.AddItem "עברי צמצום אלפים"
    NumberingBox1.ListIndex = 0
    NumberingBox2.ListIndex = 0
    NumberingBox3.ListIndex = 0
    NumberingBox4.ListIndex = 0
    NumberingBox5.ListIndex = 0
    
    NumberingBox1.ListIndex = numbering(1)
    NumberingBox2.ListIndex = numbering(2)
    NumberingBox3.ListIndex = numbering(3)
    NumberingBox4.ListIndex = numbering(4)
    NumberingBox5.ListIndex = numbering(5)
    
    ' טעינת תוחם ראשון
    BoundaryBeforeBox1.value = boundaryBefore(1)
    BoundaryBeforeBox2.value = boundaryBefore(2)
    BoundaryBeforeBox3.value = boundaryBefore(3)
    BoundaryBeforeBox4.value = boundaryBefore(4)
    BoundaryBeforeBox5.value = boundaryBefore(5)
    
    ' טעינת תוחם שני
    BoundaryBox1.value = boundary(1)
    BoundaryBox2.value = boundary(2)
    BoundaryBox3.value = boundary(3)
    BoundaryBox4.value = boundary(4)
    BoundaryBox5.value = boundary(5)
    
End Sub
Private Sub CB_Cancele_Click()
    Call WritingValues
    Unload Me
End Sub
Private Sub CB_OK_Click()

    If StyleCheck1.value = False And StyleCheck2.value = False And StyleCheck3.value = False And StyleCheck4.value = False And StyleCheck5.value = False Then MsgBox "יש לבחור לפחות כותרת 1": Exit Sub
    If StyleCheck1.value = True And (StyleBox1.value = "בחר סגנון" Or StyleBox1.value = "") Then MsgBox "יש לבחור סגנון": Exit Sub
    If StyleCheck2.value = True And (StyleBox2.value = "בחר סגנון" Or StyleBox2.value = "") Then MsgBox "יש לבחור סגנון": Exit Sub
    If StyleCheck3.value = True And (StyleBox3.value = "בחר סגנון" Or StyleBox3.value = "") Then MsgBox "יש לבחור סגנון": Exit Sub
    If StyleCheck4.value = True And (StyleBox4.value = "בחר סגנון" Or StyleBox4.value = "") Then MsgBox "יש לבחור סגנון": Exit Sub
    If StyleCheck5.value = True And (StyleBox5.value = "בחר סגנון" Or StyleBox5.value = "") Then MsgBox "יש לבחור סגנון": Exit Sub

    styleTest(1) = StyleCheck1.value
    styleTest(2) = StyleCheck2.value
    styleTest(3) = StyleCheck3.value
    styleTest(4) = StyleCheck4.value
    styleTest(5) = StyleCheck5.value
    
    If StyleCheck1.value = True Then
        styleSelectName(1) = StyleBox1.value
        continuous(1) = ContinuousBox1.ListIndex
        tub(1) = TubBox1.ListIndex
        boundaryBefore(1) = BoundaryBeforeBox1.value
        numbering(1) = NumberingBox1.ListIndex
        boundary(1) = BoundaryBox1.value
    End If
    
    If StyleCheck2.value = True Then
        styleSelectName(2) = StyleBox2.value
        continuous(2) = ContinuousBox2.ListIndex
        tub(2) = TubBox2.ListIndex
        boundaryBefore(2) = BoundaryBeforeBox2.value
        numbering(2) = NumberingBox2.ListIndex
        boundary(2) = BoundaryBox2.value
    End If
    
    If StyleCheck3.value = True Then
        styleSelectName(3) = StyleBox3.value
        continuous(3) = ContinuousBox3.ListIndex
        tub(3) = TubBox3.ListIndex
        boundaryBefore(3) = BoundaryBeforeBox3.value
        numbering(3) = NumberingBox3.ListIndex
        boundary(3) = BoundaryBox3.value
    End If
    
    If StyleCheck4.value = True Then
        styleSelectName(4) = StyleBox4.value
        continuous(4) = ContinuousBox4.ListIndex
        tub(4) = TubBox4.ListIndex
        boundaryBefore(4) = BoundaryBeforeBox4.value
        numbering(4) = NumberingBox4.ListIndex
        boundary(4) = BoundaryBox4.value
    End If
    
    If StyleCheck5.value = True Then
        styleSelectName(5) = StyleBox5.value
        continuous(5) = ContinuousBox5.ListIndex
        tub(5) = TubBox5.ListIndex
        boundaryBefore(5) = BoundaryBeforeBox5.value
        numbering(5) = NumberingBox5.ListIndex
        boundary(5) = BoundaryBox5.value
    End If
    
    userTest = True
    Call WritingValues
    Unload Me
End Sub

Private Sub CB_Design_Click()
    MsgBox "הקוד יוצר סגנונות חדשים בשם:" & vbNewLine & "תוכן עניינים כותרת 1" & vbNewLine & "תוכן עניינים כותרת 2" & vbNewLine & "ניתן לשנות את העיצוב משם"
End Sub
-------------------------------------------------------------------------------
