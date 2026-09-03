VBA MACRO DecorationsTitles.bas 
in file: word/vbaProject.bin - OLE stream: 'VBA/DecorationsTitles'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit
Option Private Module

Public userTest As Boolean
Public SelectedStyleName As String
Public ImgPathBefore As String
Public ImgPathAfter As String
Public ImgSize As Integer
Public ImgType As Integer
Public BoundaryText As String
Public PositionValue As Integer

Public ImgPathAbove As String
Public ImgPathUnder As String

Public styleToRemove As String
Public boundaryToRemove As Integer
Public typeToRemove As Integer

Sub AddBrfore()
    
    Dim SavedRange As Range
    Dim selectedStyle As Style
    Dim listTemplate As listTemplate
    Dim para As Paragraph
    Dim pic As InlineShape
    Dim styleName As String
    
    On Error GoTo Ending
    Set SavedRange = Selection.Range
    Application.UndoRecord.StartCustomRecord
    Application.ScreenUpdating = False
    
    ' הצגת הטופס
    FormAddImgBefore.Show
    If Not userTest Then GoTo Ending
    
    ' מעבר על כל הפסקאות במסמך
    For Each para In ActiveDocument.Paragraphs
        If para.Style = SelectedStyleName Then
            With para.Range
                .Collapse direction:=wdCollapseStart
                If ImgPathBefore <> "" Then
                    Set pic = .InlineShapes.AddPicture(FileName:=ImgPathBefore, LinkToFile:=False, SaveWithDocument:=True)
                    With pic
                        .LockAspectRatio = True
                        .Height = .Height * ImgSize / 100
                    End With
                    .MoveEnd 1
                    .Font.Position = PositionValue
                    .Collapse direction:=wdCollapseEnd
                    .InsertAfter BoundaryText
                End If
                .Expand wdParagraph
                .MoveEnd wdCharacter, -1
                .Collapse direction:=wdCollapseEnd
                If ImgPathAfter <> "" Then
                    Set pic = .InlineShapes.AddPicture(FileName:=ImgPathAfter, LinkToFile:=False, SaveWithDocument:=True)
                    With pic
                        .LockAspectRatio = True
                        .Height = .Height * ImgSize / 100
                    End With
                    .MoveEnd 1
                    .Font.Position = PositionValue
                    .Collapse direction:=wdCollapseStart
                    .InsertBefore BoundaryText
                End If
            End With
            
        End If
    Next para

Ending:
    SavedRange.Select
    Application.UndoRecord.EndCustomRecord
    Application.ScreenUpdating = True
End Sub

Function AddBullet()
    
    Dim selectedStyle As Style
    Dim listTemplate As listTemplate
    Dim leftIndent As Single
    Dim firstLineIndent As Single
    
    Set selectedStyle = ActiveDocument.Styles(SelectedStyleName)
    
    ' שלב 1: שמור את ערכי הכניסה הקיימים
    With selectedStyle.ParagraphFormat
        leftIndent = .leftIndent
        firstLineIndent = .firstLineIndent
    End With
    
    ' שלב 2: צור תבליט מסוג תמונה
    Set listTemplate = ListGalleries(wdBulletGallery).ListTemplates(1)
    With listTemplate.ListLevels(1)
        .ApplyPictureBullet FileName:=ImgPathBefore
        .NumberPosition = 0
        .TextPosition = 0
        .TabPosition = wdUndefined
    End With
    
    ' שלב 3: קישור הסגנון לתבליט
    selectedStyle.LinkToListTemplate listTemplate:=listTemplate, ListLevelNumber:=1
    
    ' שלב 4: החזרת ערכי הכניסה
    With selectedStyle.ParagraphFormat
        .leftIndent = leftIndent
        .firstLineIndent = firstLineIndent
    End With

End Function

Sub AddAbove()
    
    Dim SavedRange As Range
    Dim selectedStyle As Style
    Dim para As Paragraph
    Dim pic As InlineShape
    Dim styleName As String
    
    On Error GoTo Ending
    Set SavedRange = Selection.Range
    Application.UndoRecord.StartCustomRecord
    Application.ScreenUpdating = False
    
    ' הצגת הטופס
    FormAddImgAbove.Show
    
    ' מעבר על כל הפסקאות במסמך
    For Each para In ActiveDocument.Paragraphs
        If para.Style = SelectedStyleName Then
            With para.Range
                .Collapse direction:=wdCollapseStart
                If ImgPathAbove <> "" Then
                    Set pic = .InlineShapes.AddPicture(FileName:=ImgPathAbove, LinkToFile:=False, SaveWithDocument:=True)
                    With pic
                        .LockAspectRatio = True
                        .Height = .Height * ImgSize / 100
                    End With
                    .MoveEnd 1
                    .Font.Position = PositionValue
                    .Collapse direction:=wdCollapseEnd
                    .InsertAfter Chr(11)
                End If
                .Expand wdParagraph
                .MoveEnd wdCharacter, -1
                If ImgPathUnder <> "" Then
                    .InsertAfter Chr(11)
                    .Collapse direction:=wdCollapseEnd
                    Set pic = .InlineShapes.AddPicture(FileName:=ImgPathUnder, LinkToFile:=False, SaveWithDocument:=True)
                    With pic
                        .LockAspectRatio = True
                        .Height = .Height * ImgSize / 100
                    End With
                    .MoveEnd 1
                    .Font.Position = PositionValue
                    .Collapse direction:=wdCollapseStart
                End If
            End With
            
        End If
    Next para

Ending:
    SavedRange.Select
    Application.UndoRecord.EndCustomRecord
    Application.ScreenUpdating = True
End Sub

Sub Remove()
    
    Dim selectionRange As Range
    Dim img As InlineShape
    Dim rng As Range
    Dim paraRange As Range
    Dim i As Long
    
    FormRemoveImg.Show
    If userTest = False Then Exit Sub
    On Error GoTo Ending
    Application.UndoRecord.StartCustomRecord "הסרת עיטורים מכותרות"
    Application.ScreenUpdating = False
    Set selectionRange = Selection.Range
    
    For i = ActiveDocument.InlineShapes.Count To 1 Step -1
        Set img = ActiveDocument.InlineShapes(i)
        Set rng = img.Range

        If rng.Paragraphs(1).Style = styleToRemove Then
            Set paraRange = rng.Paragraphs(1).Range
            
            If (typeToRemove = 0 And Not paraRange.text Like "*" & Chr(11) & "*") Or _
               (typeToRemove = 1 And paraRange.text Like "*" & Chr(11) & "*") Then

                If rng.Start = paraRange.Start Then
                    rng.End = rng.End + boundaryToRemove
                ElseIf rng.End = paraRange.End - 1 Then
                    rng.Start = rng.Start - boundaryToRemove
                End If
                rng.Delete
                
            End If
        End If
    Next i
    
Ending:
    Application.UndoRecord.EndCustomRecord
    Application.ScreenUpdating = True
    selectionRange.Select

End Sub




-------------------------------------------------------------------------------
