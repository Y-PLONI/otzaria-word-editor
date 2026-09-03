VBA MACRO EditingErrors.bas 
in file: word/vbaProject.bin - OLE stream: 'VBA/EditingErrors'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit
Option Private Module
Public userTest As Boolean
Public userSelection As Integer
Public itemsToList() As String
Public paraStyles As Collection
Public charStyles As Collection
Public linkedStyles As Collection

Sub PagesSize()

    Dim doc As Document
    Dim sec As section
    Dim i As Integer
    
    Dim secPageHeight() As Single
    Dim secPageWidth() As Single
    Dim secTopMargin() As Single
    Dim secBottomMargin() As Single
    Dim secRightMargin() As Single
    Dim secLeftMargin() As Single
    
    Set doc = ActiveDocument
    
    ReDim Preserve secPageHeight(i)
    ReDim Preserve secPageWidth(i)
    ReDim Preserve secTopMargin(i)
    ReDim Preserve secBottomMargin(i)
    ReDim Preserve secRightMargin(i)
    ReDim Preserve secLeftMargin(i)

    On Error GoTo Ending
    Application.UndoRecord.StartCustomRecord "תיקון שגיאות גודל ושוליים"
    
    For Each sec In doc.Sections
        
        With sec.PageSetup
            
            If NotNumberInArray(secPageHeight, .PageHeight) Or _
               NotNumberInArray(secPageWidth, .PageWidth) Or _
               NotNumberInArray(secTopMargin, .TopMargin) Or _
               NotNumberInArray(secBottomMargin, .BottomMargin) Or _
               NotNumberInArray(secRightMargin, .RightMargin) Or _
               NotNumberInArray(secLeftMargin, .LeftMargin) Then
                
                ReDim Preserve secPageHeight(i)
                ReDim Preserve secPageWidth(i)
                ReDim Preserve secTopMargin(i)
                ReDim Preserve secBottomMargin(i)
                ReDim Preserve secRightMargin(i)
                ReDim Preserve secLeftMargin(i)
                
                secPageHeight(i) = .PageHeight
                secPageWidth(i) = .PageWidth
                secTopMargin(i) = .TopMargin
                secBottomMargin(i) = .BottomMargin
                secRightMargin(i) = .RightMargin
                secLeftMargin(i) = .LeftMargin
                i = i + 1
            
            End If
        
        End With
                         
    Next sec
        
    If i = 1 Then MsgBox "לא נמצאו שגיאות": GoTo Ending
    
    ReDim itemsToList(UBound(secPageHeight))
    
    For i = LBound(secPageHeight) To UBound(secPageHeight)
        
        itemsToList(i) = "אורך: " & FormatNumber(PointsToCentimeters(secPageHeight(i)), 2) & " , " & _
                         "רוחב: " & FormatNumber(PointsToCentimeters(secPageWidth(i)), 2) & " , " & _
                         "עליונים: " & FormatNumber(PointsToCentimeters(secTopMargin(i)), 2) & " , " & _
                         "תחתונים: " & FormatNumber(PointsToCentimeters(secBottomMargin(i)), 2) & " , " & _
                         "ימניים: " & FormatNumber(PointsToCentimeters(secRightMargin(i)), 2) & " , " & _
                         "שמאליים: " & FormatNumber(PointsToCentimeters(secRightMargin(i)), 2)
    Next i
    
    FormEditingErrors.Show
    If Not userTest Then GoTo Ending
    
    For Each sec In doc.Sections
        
        With sec.PageSetup
            .PageHeight = secPageHeight(userSelection)
            .PageWidth = secPageWidth(userSelection)
            .TopMargin = secTopMargin(userSelection)
            .BottomMargin = secBottomMargin(userSelection)
            .RightMargin = secRightMargin(userSelection)
            .LeftMargin = secLeftMargin(userSelection)
        End With
    
    Next sec

Ending:
    Application.UndoRecord.EndCustomRecord
End Sub
Sub ColumnWidth()

    Dim doc As Document
    Dim sec As section
    Dim i As Integer
    Dim ms As VbMsgBoxResult
    Dim userEvenlySpaced As Integer
    
    Dim firstColumnsWidth() As Single
    Dim secondColumnsWidth() As Single
    Dim spacingWidth() As Single

    Set doc = ActiveDocument
    userEvenlySpaced = 2
    
    ReDim Preserve firstColumnsWidth(i)
    ReDim Preserve secondColumnsWidth(i)
    ReDim Preserve spacingWidth(i)
    
    On Error GoTo Ending
    Application.UndoRecord.StartCustomRecord "תיקון שגיאות רוחב טורים"
    
    For Each sec In doc.Sections
        
        With sec.PageSetup.TextColumns
        
            If .Count <> 2 Then
                GoTo NextSec
            ElseIf Not .EvenlySpaced Then
                If userEvenlySpaced = 2 Then
                    ms = MsgBox("נמצאו מקטעים בהם רוחב טור אחד גדול מהשני" & vbNewLine & "האם לתקן?", vbYesNo + vbQuestion, "שאלה")
                    If ms = vbYes Then
                        userEvenlySpaced = 1
                        .EvenlySpaced = True
                    Else
                        userEvenlySpaced = 0
                        GoTo NextSec
                    End If
                ElseIf userEvenlySpaced = 1 Then
                    .EvenlySpaced = True
                End If
            End If
            
            If NotNumberInArray(firstColumnsWidth, .Item(1).Width) Or _
               NotNumberInArray(secondColumnsWidth, .Item(2).Width) Or _
               NotNumberInArray(spacingWidth, .Spacing) Then
                
                ReDim Preserve firstColumnsWidth(i)
                ReDim Preserve secondColumnsWidth(i)
                ReDim Preserve spacingWidth(i)
                
                firstColumnsWidth(i) = .Item(1).Width
                secondColumnsWidth(i) = .Item(2).Width
                spacingWidth(i) = .Item(1).SpaceAfter
                i = i + 1
            
            End If
        
        End With
NextSec:
    Next sec
        
    If i <= 1 Then MsgBox "לא נמצאו שגיאות": GoTo Ending
    
    ReDim itemsToList(UBound(firstColumnsWidth))
    
    For i = LBound(firstColumnsWidth) To UBound(firstColumnsWidth)
        
        itemsToList(i) = "טור ראשון: " & FormatNumber(PointsToCentimeters(firstColumnsWidth(i)), 2) & " , " & _
                         "טור שני: " & FormatNumber(PointsToCentimeters(secondColumnsWidth(i)), 2) & " , " & _
                         "מרווח בין טורים: " & FormatNumber(PointsToCentimeters(spacingWidth(i)), 2)
    Next i
    
    FormEditingErrors.Show
    If Not userTest Then GoTo Ending
    
    For Each sec In doc.Sections
        With sec.PageSetup.TextColumns
            If .Count <> 2 Then GoTo NextSec2
            .EvenlySpaced = False
            .Item(1).Width = 37
            .Item(1).SpaceAfter = 5
            .Item(2).Width = 37
            .Item(1).Width = firstColumnsWidth(userSelection)
            .Item(1).SpaceAfter = spacingWidth(userSelection)
            .Item(2).Width = secondColumnsWidth(userSelection)
        End With
NextSec2:
    Next sec

Ending:
    Application.UndoRecord.EndCustomRecord
End Sub
Sub SwapDocumentStyles()

    Dim stl As Style
    Set paraStyles = New Collection
    Set charStyles = New Collection
    Set linkedStyles = New Collection
    On Error GoTo Ending
    'לולאה על כל הסגנונות במסמך
    Selection.Collapse
    For Each stl In ActiveDocument.Styles
        'בדיקה האם הסגנון בשימוש במסמך
        If stl.InUse Then
            If Not FindAndReplaceStylse(stl.NameLocal) Is Nothing Or Selection.Range.Style = stl Then
                Select Case stl.Type
                    Case wdStyleTypeParagraph
                        paraStyles.Add stl.NameLocal
                    Case wdStyleTypeCharacter
                        charStyles.Add stl.NameLocal
                    Case wdStyleTypeLinked
                        linkedStyles.Add stl.NameLocal
                End Select
            End If
        End If
    Next stl
    Application.UndoRecord.StartCustomRecord "החלפת סגנונות"
    FormReplaceStylse.Show
Ending:
    Application.UndoRecord.EndCustomRecord
    Unload FormReplaceStylse
End Sub
Sub DeleteUnusedStyles()
    Dim stl As Style
    Dim i As Long
    
    On Error GoTo Ending
    Application.UndoRecord.StartCustomRecord "מחיקת סגנונות שאינם בשימוש"
    FormRunning1.Show vbModeless
    FormRunning1.Label1.Caption = "מוחק סגנונות שאינם בשימוש": DoEvents
    Selection.Collapse
    For i = ActiveDocument.Styles.Count To 1 Step -1
        Set stl = ActiveDocument.Styles(i)
        If FindAndReplaceStylse(stl.NameLocal) Is Nothing And Not Selection.Range.Style = stl Then
            On Error Resume Next
            stl.Delete
            On Error GoTo Ending
        End If
    Next i

Ending:
    Unload FormRunning1
    Application.UndoRecord.EndCustomRecord
End Sub

Function NotNumberInArray(arr() As Single, num As Single) As Boolean

    Dim i As Integer
    
    NotNumberInArray = True
    
    For i = LBound(arr) To UBound(arr)
        If num = arr(i) Then NotNumberInArray = False
    Next i
    
End Function
Function NotNumberInArray2(num() As Variant, ParamArray arr() As Variant) As Boolean

    Dim i, x As Integer
    
    NotNumberInArray2 = True
    
    For i = LBound(arr) To UBound(arr)
        For x = LBound(arr(i)) To UBound(arr(i))
            If num(x) = arr(i)(x) Then NotNumberInArray2 = False
        Next x
    Next i
End Function
Function FindAndReplaceStylse(ByVal findStyle As String, Optional ByVal replaceStyle As String = "False", Optional rng As Range = Nothing) As Range
    
    If rng Is Nothing Then Set rng = ActiveDocument.Range
    On Error Resume Next
    If IsEmpty(replaceStyle) Then replaceStyle = "False"
    On Error GoTo 0
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Style = findStyle
        .Wrap = wdFindStop
        If replaceStyle = "False" Then
            If .Execute Then Set FindAndReplaceStylse = rng.Duplicate
        Else
            .Replacement.Style = replaceStyle
            .Execute Replace:=wdReplaceAll
        End If
    End With

End Function
-------------------------------------------------------------------------------
