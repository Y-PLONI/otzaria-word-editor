VBA MACRO FormDecorationsTitles1.frm 
in file: word/vbaProject.bin - OLE stream: 'VBA/FormDecorationsTitles1'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit

Dim shp(1 To 2) As word.Shape
Dim ToggleImg(1 To 2) As MSForms.ToggleButton
Dim paraRange As Range
Dim isUpdateOnly As Boolean

Private Sub UserForm_Initialize()
    
    Dim prevShp As Shape
    
    On Error GoTo Ending
    Application.UndoRecord.StartCustomRecord "עיצוב כותרות"
    Application.ScreenUpdating = False
    
    Set paraRange = Selection.Paragraphs(1).Range
    
    Set ToggleImg(1) = Toggle_Img1
    Set ToggleImg(2) = Toggle_Img2
    
    TB_MM.value = "5"
    
    Combo_AddImg1.value = "בחר תמונה"
    Combo_AddImg1.AddItem "הוסף תמונה מהמחשב"
    Combo_AddImg1.AddItem "הוסף תמונה מהמסמך"
    Combo_AddImg1.AddItem "הסר מבלי למחוק מהמסמך"
    Combo_AddImg1.AddItem "הסר ומחק מהמסמך"
    
    Combo_AddImg2.value = "בחר תמונה"
    Combo_AddImg2.AddItem "הוסף תמונה מהמחשב"
    Combo_AddImg2.AddItem "הוסף תמונה מהמסמך"
    Combo_AddImg2.AddItem "הסר מבלי למחוק מהמסמך"
    Combo_AddImg2.AddItem "הסר ומחק מהמסמך"

    MultiPage1.value = 0
    
    ' קריאת ערכים קודמים
    Call RedingValues
    
    ' למקרה של עידכון אובייקטים קיימים
    If isUpdateOnly Then
        For Each prevShp In paraRange.ShapeRange
            If prevShp.name Like GetImgName & "*" Then
                If shp(1) Is Nothing Then
                    Set shp(1) = prevShp
                Else
                    Set shp(2) = prevShp
                    Exit For
                End If
            End If
        Next prevShp
        If shp(1) Is Nothing And shp(2) Is Nothing Then
            Unload Me
            MsgBox "לא נמצאו עיטורים בפיסקה הנוכחית", vbOKOnly, "שגיאה"
        End If
        MultiPage1.value = 1
    End If
Ending:
    Application.ScreenUpdating = True
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    Call WritingValues
    Application.UndoRecord.EndCustomRecord
    
End Sub
Private Sub CB_OK1_Click()
    
    Dim prevShp As Shape
    Dim imgNum, highNum As Integer
    Dim i As Integer
    
    ' בדיקה שנבחרו עיטורים
    If shp(1) Is Nothing And shp(2) Is Nothing Then
        MsgBox "לא נבחרו עיטורים"
        Exit Sub
    Else
        MultiPage1.value = 1
    End If

    ' מציאת מספר האובייקט האחרון בסגנון הנוכחי
    For Each prevShp In ActiveDocument.Shapes
        If prevShp.name Like GetImgName & "*" Then
            imgNum = Split(prevShp.name, ">")(2)
            If imgNum > highNum Then
                highNum = imgNum
            End If
        End If
    Next prevShp
    
    ' קביעת שם חדש לאובייקטים
    For i = LBound(ToggleImg) To UBound(ToggleImg)
        If Not shp(i) Is Nothing Then
            shp(i).WrapFormat.Type = wdWrapBehind
            shp(i).name = GetImgName & highNum + i
        End If
    Next i
        
End Sub

Private Sub CB_OK2_Click()
    MultiPage1.value = 2
End Sub

Private Sub CB_OK3_Click()
    
    Dim rng As Range
    Dim paraCollection As New Collection
    Dim para As Paragraph
    Dim i As Integer
    Dim newShp As Shape
    Dim originalName As String
    Dim paraShp As Shape
    
    ' קביעת טווח העבודה
    Set rng = ActiveDocument.Range
    If Check_StartInThisPara Then
        rng.Start = paraRange.Start
    End If
    
    ' קבלת הפיסקאות המתאימות
    Set paraCollection = GetParaToAddImg(rng)
    
    For i = LBound(ToggleImg) To UBound(ToggleImg)
        ' העתקת האובייקט
        If Not shp(i) Is Nothing Then
            With shp(i)
                originalName = .name
                .name = "newShp"
                .Select
                Selection.Copy
                .name = originalName
            End With
            For Each para In paraCollection
                ' הדבקת האובייקט והכנסתו למשתנה
                para.Range.Paste
                For Each paraShp In para.Range.ShapeRange
                    If paraShp.name = "newShp" Then
                        Set newShp = paraShp
                        Exit For
                    End If
                Next paraShp
    
                ' הגדרת נתוני המיקום והעיגון המקוריים
                With newShp
                    .Left = shp(i).Left
                    .Top = shp(i).Top
                    .RelativeHorizontalPosition = shp(i).RelativeHorizontalPosition
                    .RelativeVerticalPosition = shp(i).RelativeVerticalPosition
                    .name = originalName
                End With
            Next para
        End If
    Next i
        
    ' סגירת הטופס
    Unload Me
    
End Sub

Private Sub Combo_AddImg1_Change()
    
    On Error Resume Next
    
    CB_OKAddShp1.Visible = False
    
    With Combo_AddImg1
        
        If .ListIndex = 0 Then
            ' הוספת התמונה
            Set shp(1) = ActiveDocument.Shapes.AddPicture( _
                FileName:=Helpers.ShowPicDialog("בחר תמונה עבור עיטור 1"), _
                LinkToFile:=False, _
                SaveWithDocument:=True, _
                Anchor:=paraRange)
            If shp(1) Is Nothing Then .value = "בחר תמונה"
        ElseIf .ListIndex = 1 Then
            CB_OKAddShp1.Visible = True
            
        ElseIf .ListIndex = 2 Then
            Set shp(1) = Nothing
            .value = "בחר תמונה"
        
        ElseIf .ListIndex = 3 Then
            If Not shp(1) Is Nothing Then shp(1).Delete
            .value = "בחר תמונה"
        End If
    
    End With

    ' מיקום יחסי לפסקה הנוכחית
    Call AnchorToPara(shp(1))

End Sub
Private Sub Combo_AddImg2_Change()
    
    On Error Resume Next
    
    CB_OKAddShp2.Visible = False
    
    With Combo_AddImg2
        If .ListIndex = 0 Then
            ' הוסף את התמונה
            Set shp(2) = ActiveDocument.Shapes.AddPicture( _
                FileName:=Helpers.ShowPicDialog("בחר תמונה עבור עיטור 2"), _
                LinkToFile:=False, _
                SaveWithDocument:=True, _
                Anchor:=paraRange)
            If shp(2) Is Nothing Then .value = "בחר תמונה"
            
        ElseIf .ListIndex = 1 Then
            CB_OKAddShp2.Visible = True
            
        ElseIf .ListIndex = 2 Then
            Set shp(2) = Nothing
            .value = "בחר תמונה"
        
        ElseIf .ListIndex = 3 Then
            If Not shp(2) Is Nothing Then shp(2).Delete
            .value = "בחר תמונה"
        End If
    End With
            
    ' מיקום יחסי לפסקה הנוכחית
    Call AnchorToPara(shp(2))
    
End Sub
Private Sub CB_OKAddShp1_Click()

    Set shp(1) = Nothing
    
    On Error Resume Next
    Set shp(1) = Selection.ShapeRange(1)
    
    If shp(1) Is Nothing Then
        MsgBox "לא נבחרה תמונה", vbOKOnly, "שגיאה"
        Combo_AddImg1.value = "בחר תמונה"
    End If
    
    CB_OKAddShp1.Visible = False
    
End Sub
Private Sub CB_OKAddShp2_Click()

    Set shp(2) = Nothing
    
    On Error Resume Next
    Set shp(2) = Selection.ShapeRange(1)
    
    If shp(2) Is Nothing Then
        MsgBox "לא נבחרה תמונה", vbOKOnly, "שגיאה"
        Combo_AddImg2.value = "בחר תמונה"
    End If
    
    CB_OKAddShp2.Visible = False
    
End Sub

Private Sub SP_MM_Change()
    
    TB_MM.value = SP_MM.value / 10

End Sub
Private Sub TB_MM_Change()
    
    If IsNumeric(TB_MM.value) Then
        If TB_MM.value >= 0.1 And TB_MM.value <= 10 Then
            SP_MM.value = TB_MM.value * 10
        Else
            TB_MM.value = SP_MM.value / 10
        End If
    Else
        TB_MM.value = SP_MM.value / 10
    End If
    
End Sub

Private Sub CB_Expand_Click()
    
    Dim i As Integer
    
    On Error GoTo Ending
    Application.ScreenUpdating = False
    
    For i = LBound(ToggleImg) To UBound(ToggleImg)
        If ToggleImg(i).value = True And Not shp(i) Is Nothing Then
            SplitImage shp(i), MM_Value
        End If
    Next i

Ending:
    Application.ScreenUpdating = True

End Sub
Private Sub CB_Shrink_Click()

    Dim i As Integer
    
    On Error GoTo Ending
    Application.ScreenUpdating = False
    
    For i = LBound(ToggleImg) To UBound(ToggleImg)
        If ToggleImg(i).value = True And Not shp(i) Is Nothing Then
            SplitImage shp(i), -MM_Value
        End If
    Next i

Ending:
    Application.ScreenUpdating = True

End Sub

Private Sub CB_AlignCenter_Click()
    
    Dim i As Integer
    
    For i = LBound(ToggleImg) To UBound(ToggleImg)
        If ToggleImg(i).value = True And Not shp(i) Is Nothing Then
            shp(i).Left = (Helpers.GetPageWidth(paraRange) - shp(i).Width) / 2
        End If
    Next i
    
End Sub

Private Sub CB_ZoomIn_Click()
    
    Dim i As Integer
    
    For i = LBound(ToggleImg) To UBound(ToggleImg)
        If ToggleImg(i).value = True And Not shp(i) Is Nothing Then
            shp(i).LockAspectRatio = True
            shp(i).Width = shp(i).Width + MM_Value
        End If
    Next i

End Sub
Private Sub CB_ZoomOut_Click()

    Dim i As Integer
    
    For i = LBound(ToggleImg) To UBound(ToggleImg)
        If ToggleImg(i).value = True And Not shp(i) Is Nothing Then
            shp(i).LockAspectRatio = True
            shp(i).Width = shp(i).Width - MM_Value
        End If
    Next i
    
End Sub

Private Sub CB_LeftRotation_Click()
    
    Dim i As Integer
    
    For i = LBound(ToggleImg) To UBound(ToggleImg)
        If ToggleImg(i).value = True And Not shp(i) Is Nothing Then
            shp(i).rotation = shp(i).rotation - 90
        End If
    Next i
    
End Sub
Private Sub CB_RightRotation_Click()
    
    Dim i As Integer
    
    For i = LBound(ToggleImg) To UBound(ToggleImg)
        If ToggleImg(i).value = True And Not shp(i) Is Nothing Then
            shp(i).rotation = shp(i).rotation + 90
        End If
    Next i

End Sub

Private Sub CB_MoveUp_Click()
    
    Dim i As Integer
    
    For i = LBound(ToggleImg) To UBound(ToggleImg)
        If ToggleImg(i).value = True And Not shp(i) Is Nothing Then
            shp(i).Top = shp(i).Top - MM_Value
        End If
    Next i

End Sub
Private Sub CB_MoveDown_Click()
    
    Dim i As Integer
    
    For i = LBound(ToggleImg) To UBound(ToggleImg)
        If ToggleImg(i).value = True And Not shp(i) Is Nothing Then
            shp(i).Top = shp(i).Top + MM_Value
        End If
    Next i

End Sub
Private Sub CB_MoveLeft_Click()

    Dim i As Integer
    
    For i = LBound(ToggleImg) To UBound(ToggleImg)
        If ToggleImg(i).value = True And Not shp(i) Is Nothing Then
            shp(i).Left = shp(i).Left - MM_Value
        End If
    Next i

End Sub
Private Sub CB_MoveRight_Click()

    Dim i As Integer
    
    For i = LBound(ToggleImg) To UBound(ToggleImg)
        If ToggleImg(i).value = True And Not shp(i) Is Nothing Then
            shp(i).Left = shp(i).Left + MM_Value
        End If
    Next i

End Sub

Private Sub CB_Cancel1_Click()
    
    Unload Me
    
End Sub
Private Sub CB_Cancel2_Click()
    
    Unload Me
    
End Sub
Private Sub CB_Cancel3_Click()
    
    Unload Me
    
End Sub

Function SplitImage(shpToSplit As Shape, expansionSize As Single)
    
    Dim splitShp(1 To 2) As Shape
    Dim partWidth As Double
    Dim newShp As Shape
    Dim i As Integer
    
    ' גודל החיתוך
    partWidth = (shpToSplit.Width - expansionSize) / 2
    
    ' יצירת 2 תמונות
    For i = 1 To 2
        Set splitShp(i) = shp(1).Duplicate
        splitShp(i).name = "split" & i
    
    Next i
    
    ' חיתוך וסידור התמונות זו לצד זו
    With splitShp(1)
        .PictureFormat.CropRight = partWidth
        .LockAspectRatio = msoFalse
        .Width = shpToSplit.Width - .PictureFormat.CropLeft - .PictureFormat.CropRight
        .Left = shpToSplit.Left
    End With
    
    With splitShp(2)
        .PictureFormat.CropLeft = partWidth
        .LockAspectRatio = msoFalse
        .Width = shpToSplit.Width - .PictureFormat.CropLeft - .PictureFormat.CropRight
        .Left = splitShp(1).Left + splitShp(1).Width - 1
    End With
    
    ' קיבוץ
    Set newShp = ActiveDocument.Shapes.Range(Array( _
        splitShp(1).name, splitShp(2).name)).Group
    Set newShp = ConvertToPic(newShp)
    
    ' החלפת התמונה הראשית בתמונה החדשה
    newShp.name = shp.name
    shpToSplit.Delete
    Set shpToSplit = newShp

End Function

Function ConvertToPic(pic As Shape) As Shape

    Dim widthShp As Double, leftPos As Double, topPos As Double
    Dim anch As Range
    Dim newPic As Shape

    ' שמור מיקום ועוגן של התמונה המקורית
    widthShp = pic.Width
    leftPos = pic.Left
    topPos = pic.Top
    Set anch = pic.Anchor

    ' העתק את הצורה
    pic.Select
    Selection.Copy

    ' הדבק כתמונה חדשה (מוחק את הקיים)
    anch.PasteSpecial DataType:=wdPasteMetafilePicture
    Set newPic = anch.ShapeRange(1)

    ' שמור על מיקום
    With newPic
        .LockAspectRatio = True
        .Width = widthShp
        .Left = leftPos
        .Top = topPos
    End With
    
    ' החזר את הצורה החדשה
    Set ConvertToPic = newPic

End Function
Function AnchorToPara(ByRef shpToAnchor As Shape)
    
    Dim shp As Shape
    Dim originalName As String
    
    If Not shpToAnchor Is Nothing Then
        
        originalName = shpToAnchor.name
        shpToAnchor.name = "newShp"
        shpToAnchor.Select
        Selection.Cut
        DoEvents
        paraRange.Paste
        
        For Each shp In paraRange.ShapeRange
            If shp.name = "newShp" Then
                Set shpToAnchor = shp
                shpToAnchor.name = originalName
                Exit For
            End If
        Next shp

        With shpToAnchor
            .RelativeHorizontalPosition = wdRelativeHorizontalPositionMargin
            .Left = 0
            .RelativeVerticalPosition = wdRelativeVerticalPositionParagraph
            .Top = 0
        End With

    End If
        
End Function
Function MM_Value() As Double
    
    MM_Value = (CentimetersToPoints(CSng(TB_MM.value)) / 10)

End Function
Function GetImgName() As String

    GetImgName = paraRange.Paragraphs(1).Style & ">" & "עיטור צף>"

End Function
Function GetParaToAddImg(rng As Range) As Collection
    
    Dim doc As Document
    Dim tempoCollection As New Collection
    Dim para As Paragraph
    Dim originalParaLinesCount, paraLinesCount As Integer
    Dim styleName As String
    Dim shp As Shape
    Dim i As Integer
    
    ' קביעת שם הסגנון ומספר השורות
    styleName = paraRange.ParagraphStyle
    originalParaLinesCount = paraRange.ComputeStatistics(wdStatisticLines)
    
    For Each para In rng.Paragraphs
        ' בדיקת מספר השורות בפיסקה הנוכחית
        paraLinesCount = para.Range.ComputeStatistics(wdStatisticLines)
        
        ' הכנסת הפיסקה לאוסף במידה והתנאים מתקיימים
        If para.Style = styleName _
            And (Not Check_SkipParaOtherLines Or originalParaLinesCount = paraLinesCount) _
            And para.Range.Start <> paraRange.Start _
            And Len(para.Range.text) > 3 _
            Then
                tempoCollection.Add para
                If Check_DeleteOtherImg Then
                    For i = para.Range.ShapeRange.Count To 1 Step -1
                        Set shp = para.Range.ShapeRange(i)
                        If shp.name Like GetImgName & "*" Then
                            shp.Delete
                        End If
                    Next i
                End If
        End If
        
    Next para
    
    ' החזרת הפונקציה
    Set GetParaToAddImg = tempoCollection

End Function
Private Sub WritingValues()
    
    SettingsHelper.Save appName, "DecorationsTitles", "Toggle_Img1", Toggle_Img1.value
    SettingsHelper.Save appName, "DecorationsTitles", "Toggle_Img2", Toggle_Img2.value
    SettingsHelper.Save appName, "DecorationsTitles", "SP_MM", SP_MM.value
    SettingsHelper.Save appName, "DecorationsTitles", "Check_StartInThisPara", Check_StartInThisPara.value
    SettingsHelper.Save appName, "DecorationsTitles", "Check_SkipParaOtherLines", Check_SkipParaOtherLines.value
    SettingsHelper.Save appName, "DecorationsTitles", "Check_DeleteOtherImg", Check_DeleteOtherImg.value
    
End Sub
Private Sub RedingValues()
    
    Toggle_Img1.value = SettingsHelper.GetSavedSetting(appName, "DecorationsTitles", "Toggle_Img1", False)
    Toggle_Img1.value = SettingsHelper.GetSavedSetting(appName, "DecorationsTitles", "Toggle_Img1", False)
    Toggle_Img2.value = SettingsHelper.GetSavedSetting(appName, "DecorationsTitles", "Toggle_Img2", False)
    SP_MM.value = SettingsHelper.GetSavedSetting(appName, "DecorationsTitles", "SP_MM", 5)
    Check_StartInThisPara.value = SettingsHelper.GetSavedSetting(appName, "DecorationsTitles", "Check_StartInThisPara", False)
    Check_SkipParaOtherLines.value = SettingsHelper.GetSavedSetting(appName, "DecorationsTitles", "Check_SkipParaOtherLines", False)
    Check_DeleteOtherImg.value = SettingsHelper.GetSavedSetting(appName, "DecorationsTitles", "Check_DeleteOtherImg", False)
    isUpdateOnly = SettingsHelper.GetSavedSetting(appName, "DecorationsTitles1", "isUpdateOnly", False)
    
End Sub
-------------------------------------------------------------------------------
