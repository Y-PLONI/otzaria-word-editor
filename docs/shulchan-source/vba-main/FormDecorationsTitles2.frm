VBA MACRO FormDecorationsTitles2.frm 
in file: word/vbaProject.bin - OLE stream: 'VBA/FormDecorationsTitles2'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit

Dim inShpRng(1 To 2) As Range
Dim ToggleImg(1 To 2) As MSForms.ToggleButton
Dim paraRange As Range
Dim isUpdateOnly As Boolean

Private Sub TB_OnSides_Click()
    Dim i As Integer
    If TB_OnSides.value = True Then
        TB_TopBottom.value = False
        For i = LBound(ToggleImg) To UBound(ToggleImg)
            If inShpRng(i).InlineShapes.Count > 0 Then
                If i = 1 Then
                    inShpRng(i).MoveStart wdCharacter, 1
                Else
                    inShpRng(i).MoveEnd wdCharacter, -1
                End If
                inShpRng(i).text = " "
                If i = 1 Then
                    inShpRng(i).MoveStart wdCharacter, -1
                Else
                    inShpRng(i).MoveEnd wdCharacter, 1
                End If
            End If
        Next i
    Else
        TB_TopBottom.value = True
    End If
End Sub

Private Sub TB_TopBottom_Click()
    Dim i As Integer
    
    If TB_TopBottom.value = True Then
        TB_OnSides.value = False
        For i = LBound(ToggleImg) To UBound(ToggleImg)
            If inShpRng(i).InlineShapes.Count > 0 Then
                If i = 1 Then
                    inShpRng(i).MoveStart wdCharacter, 1
                Else
                    inShpRng(i).MoveEnd wdCharacter, -1
                End If
                inShpRng(i).text = Chr(11)
                If i = 1 Then
                    inShpRng(i).MoveStart wdCharacter, -1
                Else
                    inShpRng(i).MoveEnd wdCharacter, 1
                End If
            End If
        Next i
    Else
        TB_OnSides.value = True
    End If

End Sub

Private Sub UserForm_Initialize()
    
    Dim prevShp As Shape
    Dim i As Integer
    
    On Error GoTo Ending
    Application.UndoRecord.StartCustomRecord "עיצוב כותרות"
    Application.ScreenUpdating = False
    
    Set paraRange = Selection.Paragraphs(1).Range
    
    Set ToggleImg(1) = Toggle_Img1
    Set ToggleImg(2) = Toggle_Img2
    
    TB_MM.value = "5"
    
    Combo_AddImg1.value = "בחר תמונה"
    Combo_AddImg1.AddItem "הוסף תמונה מהמחשב"
'    Combo_AddImg1.AddItem "הוסף תמונה מהמסמך"
    Combo_AddImg1.AddItem "הסר מבלי למחוק מהמסמך"
    Combo_AddImg1.AddItem "הסר ומחק מהמסמך"
    
    Combo_AddImg2.value = "בחר תמונה"
    Combo_AddImg2.AddItem "הוסף תמונה מהמחשב"
'    Combo_AddImg2.AddItem "הוסף תמונה מהמסמך"
    Combo_AddImg2.AddItem "הסר מבלי למחוק מהמסמך"
    Combo_AddImg2.AddItem "הסר ומחק מהמסמך"

    MultiPage1.value = 0
    
    ' קריאת ערכים קודמים
    Call RedingValues
    ' למקרה של עידכון אובייקטים קיימים
    If isUpdateOnly Then
        
        Set inShpRng(1) = paraRange.Duplicate
        Set inShpRng(2) = paraRange.Duplicate
        inShpRng(1).Collapse wdCollapseStart
        inShpRng(1).MoveEnd wdCharacter, 2
        inShpRng(2).Collapse wdCollapseEnd
        inShpRng(2).MoveStart wdCharacter, -3
        inShpRng(2).MoveEnd wdCharacter, -1
        
        For i = LBound(inShpRng) To UBound(inShpRng)
            If inShpRng(i).InlineShapes.Count = 1 Then
                If Not inShpRng(i).InlineShapes(1).AlternativeText Like GetImgName & "*" Then
                    Set inShpRng(i) = Nothing
                End If
            Else
                Set inShpRng(i) = Nothing
            End If
        Next i
        
        If inShpRng(1) Is Nothing And inShpRng(2) Is Nothing Then
            Unload Me
            MsgBox "לא נמצאו עיטורים בפיסקה הנוכחית", vbOKOnly, "שגיאה"
            GoTo Ending
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
    
    Dim prevInShp As InlineShape
    Dim imgNum, highNum As Integer
    Dim i As Integer
    
    ' בדיקה שנבחרו עיטורים
    If inShpRng(1) Is Nothing And inShpRng(2) Is Nothing Then
        MsgBox "לא נבחרו עיטורים"
        Exit Sub
    Else
        MultiPage1.value = 1
    End If

    ' מציאת מספר האובייקט האחרון בסגנון הנוכחי
    For Each prevInShp In ActiveDocument.InlineShapes
        If prevInShp.AlternativeText Like GetImgName & "*" Then
            imgNum = Split(prevInShp.AlternativeText, ">")(2)
            If imgNum > highNum Then
                highNum = imgNum
            End If
        End If
    Next prevInShp
    If highNum Mod 2 <> 0 Then highNum = highNum + 1
    
    ' קביעת שם חדש לאובייקטים
    For i = LBound(ToggleImg) To UBound(ToggleImg)
        If inShpRng(i).InlineShapes.Count > 0 Then
            Call SetInlineShapeName(inShpRng(i), GetImgName & highNum + i)
            inShpRng(i).InlineShapes(1).AlternativeText = GetImgName & highNum + i
        End If
    Next i
    
    TB_OnSides.value = True
End Sub

Private Sub CB_OK2_Click()
    MultiPage1.value = 2
End Sub

Private Sub CB_OK3_Click()
    
    Dim rng As Range
    Dim paraCollection As New Collection
    Dim para As Paragraph
    Dim paraTopPasteRng As Range
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
        If inShpRng(i).InlineShapes.Count > 0 Then
            inShpRng(i).Select
            Selection.Copy
            DoEvents
            For Each para In paraCollection
                Set paraTopPasteRng = para.Range
                If i = 1 Then
                    paraTopPasteRng.Collapse wdCollapseStart
                Else
                    paraTopPasteRng.MoveEnd wdCharacter, -1
                    paraTopPasteRng.Collapse wdCollapseEnd
                End If
                paraTopPasteRng.Paste
            Next para
        End If
    Next i
        
    ' סגירת הטופס
    Unload Me
    
End Sub

Private Sub Combo_AddImg1_Change()
    
    On Error Resume Next
    
    With Combo_AddImg1
        
        If .ListIndex = 0 Then
            ' הוספת התמונה
            Set inShpRng(1) = paraRange.Duplicate
            inShpRng(1).Collapse wdCollapseStart
            ActiveDocument.InlineShapes.AddPicture _
                FileName:=Helpers.ShowPicDialog("בחר תמונה עבור עיטור 1"), _
                LinkToFile:=False, _
                SaveWithDocument:=True, _
                Range:=inShpRng(1)
            inShpRng(1).End = inShpRng(1).End + 1
            If inShpRng(1).InlineShapes.Count = 0 Then
                .value = "בחר תמונה"
            Else
                inShpRng(1).InsertAfter " "
            End If
        ElseIf .ListIndex = 1 Then
            Set inShpRng(1) = Nothing
            .value = "בחר תמונה"
        
        ElseIf .ListIndex = 2 Then
            If inShpRng(1).InlineShapes.Count > 0 Then inShpRng(1).Delete
            .value = "בחר תמונה"
        End If
    
    End With

End Sub
Private Sub Combo_AddImg2_Change()
    
    On Error Resume Next
    
    With Combo_AddImg2
        
        If .ListIndex = 0 Then
            ' הוספת התמונה
            Set inShpRng(2) = paraRange.Duplicate
            inShpRng(2).MoveEnd wdCharacter, -1
            inShpRng(2).Collapse wdCollapseEnd
            ActiveDocument.InlineShapes.AddPicture _
                FileName:=Helpers.ShowPicDialog("בחר תמונה עבור עיטור 2"), _
                LinkToFile:=False, _
                SaveWithDocument:=True, _
                Range:=inShpRng(2)
            inShpRng(2).End = inShpRng(2).End + 1
            If inShpRng(2).InlineShapes.Count = 0 Then
                .value = "בחר תמונה"
            Else
                inShpRng(2).InsertBefore " "
            End If
        
        ElseIf .ListIndex = 1 Then
            Set inShpRng(2) = Nothing
            .value = "בחר תמונה"
        
        ElseIf .ListIndex = 2 Then
            If inShpRng(2).InlineShapes.Count > 0 Then inShpRng(2).Delete
            .value = "בחר תמונה"
        End If
    
    End With

End Sub
'Private Sub CB_OKAddShp1_Click()
'
'    Dim shp As Shape
'
'    Set inShpRng(1) = Nothing
'
'    On Error Resume Next
'    Set shp = Selection.ShapeRange(1)
'
'    If shp Is Nothing Then
'        MsgBox "לא נבחרה תמונה", vbOKOnly, "שגיאה"
'        Combo_AddImg1.value = "בחר תמונה"
'    Else
'        Set inShpRng(1) = shp.ConvertToInlineShape
'
'    End If
'
'    CB_OKAddShp1.Visible = False
'
'End Sub
'Private Sub CB_OKAddShp2_Click()
'
'    Set inShpRng(2) = Nothing
'
'    On Error Resume Next
'    Set inShpRng(2) = Selection.ShapeRange(1)
'
'    If inShpRng(2) Is Nothing Then
'        MsgBox "לא נבחרה תמונה", vbOKOnly, "שגיאה"
'        Combo_AddImg2.value = "בחר תמונה"
'    End If
'
'    CB_OKAddShp2.Visible = False
'
'End Sub

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

Private Sub CB_ZoomIn_Click()
    
    Dim i As Integer
    
    For i = LBound(ToggleImg) To UBound(ToggleImg)
        If ToggleImg(i).value = True And inShpRng(i).InlineShapes.Count > 0 Then
            inShpRng(i).InlineShapes(1).LockAspectRatio = True
            inShpRng(i).InlineShapes(1).Width = inShpRng(i).InlineShapes(1).Width + MM_Value
        End If
    Next i

End Sub
Private Sub CB_ZoomOut_Click()

    Dim i As Integer
    
    For i = LBound(ToggleImg) To UBound(ToggleImg)
        If ToggleImg(i).value = True And inShpRng(i).InlineShapes.Count > 0 Then
            inShpRng(i).InlineShapes(1).LockAspectRatio = True
            inShpRng(i).InlineShapes(1).Width = inShpRng(i).InlineShapes(1).Width - MM_Value
        End If
    Next i
    
End Sub

Private Sub CB_LeftRotation_Click()
    
    Dim i As Integer
    
    For i = LBound(ToggleImg) To UBound(ToggleImg)
        If ToggleImg(i).value = True And inShpRng(i).InlineShapes.Count > 0 Then
            Set inShpRng(i) = RotationInlineShape(inShpRng(i), -90)
        End If
    Next i
    
End Sub
Private Sub CB_RightRotation_Click()
    
    Dim i As Integer
    
    For i = LBound(ToggleImg) To UBound(ToggleImg)
        If ToggleImg(i).value = True And inShpRng(i).InlineShapes.Count > 0 Then
            Set inShpRng(i) = RotationInlineShape(inShpRng(i), 90)
        End If
    Next i

End Sub

Private Sub CB_MoveUp_Click()
    
    Dim i As Integer
    
    For i = LBound(ToggleImg) To UBound(ToggleImg)
        If ToggleImg(i).value = True And inShpRng(i).InlineShapes.Count > 0 Then
            inShpRng(i).Font.Position = inShpRng(i).Font.Position + 1
        End If
    Next i

End Sub
Private Sub CB_MoveDown_Click()
    
    Dim i As Integer
    
    For i = LBound(ToggleImg) To UBound(ToggleImg)
        If ToggleImg(i).value = True And inShpRng(i).InlineShapes.Count > 0 Then
            inShpRng(i).Font.Position = inShpRng(i).Font.Position - 1
        End If
    Next i

End Sub
Private Sub CB_MoveLeft_Click()

    Dim i As Integer
    
    For i = LBound(ToggleImg) To UBound(ToggleImg)
        If ToggleImg(i).value = True And inShpRng(i).InlineShapes.Count > 0 Then
            inShpRng(i).Font.Spacing = inShpRng(i).Font.Spacing - MM_Value
        End If
    Next i

End Sub
Private Sub CB_MoveRight_Click()

    Dim i As Integer
    
    For i = LBound(ToggleImg) To UBound(ToggleImg)
        If ToggleImg(i).value = True And inShpRng(i).InlineShapes.Count > 0 Then
            inShpRng(i).Font.Spacing = inShpRng(i).Font.Spacing + MM_Value
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

Function MM_Value() As Double
    
    MM_Value = (CentimetersToPoints(CSng(TB_MM.value)) / 10)

End Function
Function GetImgName() As String

    GetImgName = paraRange.Paragraphs(1).Style & ">" & "עיטור כטקסט>"

End Function
Function GetParaToAddImg(rng As Range) As Collection
    
    Dim doc As Document
    Dim tempoCollection As New Collection
    Dim para As Paragraph
    Dim originalParaLinesCount, paraLinesCount As Integer
    Dim styleName As String
    Dim inShp As InlineShape
    Dim rngToRemove As Range
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
                    For i = para.Range.InlineShapes.Count To 1 Step -1
                        Set inShp = para.Range.InlineShapes(i)
                        If inShp.AlternativeText Like GetImgName & "*" Then
                            Set rngToRemove = inShp.Range
                            If Split(inShp.AlternativeText, ">")(2) Mod 2 <> 0 Then
                                rngToRemove.MoveEnd wdCharacter, 1
                            Else
                                rngToRemove.MoveStart wdCharacter, -1
                            End If
                            rngToRemove.Delete
                        End If
                    Next i
                End If
        End If
        
    Next para
    
    ' החזרת הפונקציה
    Set GetParaToAddImg = tempoCollection

End Function
Private Sub WritingValues()
    
    SettingsHelper.Save appName, "DecorationsTitles2", "Toggle_Img1", Toggle_Img1.value
    SettingsHelper.Save appName, "DecorationsTitles2", "Toggle_Img2", Toggle_Img2.value
    SettingsHelper.Save appName, "DecorationsTitles2", "SP_MM", SP_MM.value
    SettingsHelper.Save appName, "DecorationsTitles2", "Check_StartInThisPara", Check_StartInThisPara.value
    SettingsHelper.Save appName, "DecorationsTitles2", "Check_SkipParaOtherLines", Check_SkipParaOtherLines.value
    SettingsHelper.Save appName, "DecorationsTitles2", "Check_DeleteOtherImg", Check_DeleteOtherImg.value
    
End Sub
Private Sub RedingValues()
    
    Toggle_Img1.value = SettingsHelper.GetSavedSetting(appName, "DecorationsTitles2", "Toggle_Img1", False)
    Toggle_Img2.value = SettingsHelper.GetSavedSetting(appName, "DecorationsTitles2", "Toggle_Img2", False)
    SP_MM.value = SettingsHelper.GetSavedSetting(appName, "DecorationsTitles2", "SP_MM", 5)
    Check_StartInThisPara.value = SettingsHelper.GetSavedSetting(appName, "DecorationsTitles2", "Check_StartInThisPara", False)
    Check_SkipParaOtherLines.value = SettingsHelper.GetSavedSetting(appName, "DecorationsTitles2", "Check_SkipParaOtherLines", False)
    Check_DeleteOtherImg.value = SettingsHelper.GetSavedSetting(appName, "DecorationsTitles2", "Check_DeleteOtherImg", False)
    isUpdateOnly = SettingsHelper.GetSavedSetting(appName, "DecorationsTitles2", "isUpdateOnly", False)
    
End Sub
Function RotationInlineShape(inShpRng As Range, rotation As Integer) As Range
    
    Dim inShp As InlineShape
    Dim shp As Shape
    
    Set inShp = inShpRng.InlineShapes(1)
    
    Set shp = inShp.ConvertToShape
    shp.rotation = shp.rotation + rotation
    Set inShp = shp.ConvertToInlineShape

    Set RotationInlineShape = inShpRng
End Function

Function SetInlineShapeName(inShpRng As Range, name As String) As Range
    
    Dim inShp As InlineShape
    Dim shp As Shape
    
    Set inShp = inShpRng.InlineShapes(1)
    
    Set shp = inShp.ConvertToShape
    shp.name = name
    Set inShp = shp.ConvertToInlineShape

    Set SetInlineShapeName = inShpRng
End Function



-------------------------------------------------------------------------------
