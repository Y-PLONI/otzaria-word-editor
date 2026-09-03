olevba 0.60.2 on Python 3.14.4 - http://decalage.info/python/oletools
===============================================================================
FILE: C:/Users/User/Downloads/שולחן העורך - עיצוב פיסקה.dotm
Type: OpenXML
WARNING  For now, VBA stomping cannot be detected for files in memory
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls 
in file: word/vbaProject.bin - OLE stream: 'VBA/ThisDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO FontDinamiHelper.bas 
in file: word/vbaProject.bin - OLE stream: 'VBA/FontDinamiHelper'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Dim sizeOption As Integer
Dim pointsPercentage As String
Dim FontOption As Boolean
Dim ColorOption As Boolean
Dim strikethroughOption As Boolean
Dim BoldOption As Boolean
Dim ItalicOption As Boolean
Dim FillOption As Boolean
Dim EffectsOption As Boolean
Dim advancedOption As Boolean
Dim UnderlineOption As Boolean

'Dim pointsDi As Single
'Dim PointsBiDi As Single
'
'Dim FontDi As String               ' שם הגופן
'Dim FontMajorDi As String          ' גופן ראשי לפי ערכת עיצוב
'Dim FontLowAnsiDi As String        ' גופן לתווים באנגלית (Low ANSI)
'Dim FontHighAnsiDi As String       ' גופן לתווים מורחבים (High ANSI)
'Dim FontNameBiDi As String         ' שם גופן לטקסט דו-כיווני
'
'' === 3. צבע ===
'Dim ColorDi As Long                ' צבע גופן רגיל
'Dim ColorRGBDi As Long             ' צבע גופן בפורמט RGB
'
'' === 4. קו חוצה ===
'Dim strikethroughDi As Boolean     ' קו חוצה יחיד על הטקסט
'Dim DoubleStrikeThroughDi As Boolean ' קו חוצה כפול
'
'' === 5. מודגש ===
'Dim BoldDi As Boolean              ' הדגשת טקסט
'Dim BoldBiDi As Boolean            ' הדגשה לטקסט דו-כיווני
'
'' === 6. נטוי ===
'Dim ItalicDi As Boolean            ' טקסט נטוי
'Dim ItalicBiDi As Boolean          ' טקסט נטוי לטקסט דו-כיווני
'
'' === 7. מיתאר ומילוי טקסט ===
'Dim ShadowDi As Boolean            ' צל לטקסט
'Dim OutlineDi As Boolean           ' קווי מתאר בלבד
'Dim EmbossDi As Boolean            ' הקלה (בליטה) של הטקסט
'Dim EngraveDi As Boolean           ' חריטה של הטקסט
'
'' === 8. אפקטי טקסט ===
'Dim SuperscriptDi As Boolean       ' כתב עילי (x²)
'Dim SubscriptDi As Boolean         ' כתב תחתי (H?O)
'Dim HiddenDi As Boolean            ' הסתרת הטקסט
'Dim SmallCapsDi As Boolean         ' אותיות קטנות באותיות רישיות
'Dim AllCapsDi As Boolean           ' כל הטקסט באותיות רישיות
'
'' === 9. קו תחתי ===
'Dim UnderlineDi As WdUnderline     ' סוג הקו התחתי
'Dim UnderlineColorDi As Long       ' צבע הקו התחתי
'
'' === 10. כל השאר ===
'Dim DefaultDi As Boolean           ' הגדרות ברירת מחדל של עיצוב
'Dim TabDi As Variant               ' טאבולציות (אם רלוונטי)
'
'' === 11. מתקדם ===
'Dim SpacingDi As String            ' ריווח בין אותיות
'Dim positionDi As String           ' מיקום אנכי של הטקסט
'Dim KerningDi As Boolean           ' קרנינג (ריווח חכם בין תווים)
'Dim KerningMinDi As Single         ' גודל מינימלי להפעלת קרנינג
'Dim ScaleDi As String              ' קנה מידה אופקי של הטקסט
Function GetFirstWordDinamiOption()


    sizeOption = GetSavedSetting(RibbonControl.appName, "FirsrWordDinami", "SizeOption", 1)
    pointsPercentage = GetSavedSetting(RibbonControl.appName, "FirsrWordDinami", "Points", 30)
    FontOption = GetSavedSetting(RibbonControl.appName, "FirsrWordDinami", "Font", False)
    ColorOption = GetSavedSetting(RibbonControl.appName, "FirsrWordDinami", "Color", False)
    strikethroughOption = GetSavedSetting(RibbonControl.appName, "FirsrWordDinami", "StrikeThrough", False)
    BoldOption = GetSavedSetting(RibbonControl.appName, "FirsrWordDinami", "Bold", False)
    ItalicOption = GetSavedSetting(RibbonControl.appName, "FirsrWordDinami", "Italic", False)
    FillOption = GetSavedSetting(RibbonControl.appName, "FirsrWordDinami", "Fill", False)
    EffectsOption = GetSavedSetting(RibbonControl.appName, "FirsrWordDinami", "Effects", False)
    advancedOption = GetSavedSetting(RibbonControl.appName, "FirsrWordDinami", "advanced", False)
    UnderlineOption = GetSavedSetting(RibbonControl.appName, "FirsrWordDinami", "Underline", False)

End Function

Function tempoStyle() As style
    Dim rng As Range
    Dim tempStyle As style

    Set rng = Selection.Range

    ' בדוק אם הסגנון קיים, ואם לא – צור אותו
    On Error Resume Next
    Set tempStyle = ActiveDocument.Styles("tempStyle")
    If tempStyle Is Nothing Then
        Set tempStyle = ActiveDocument.Styles.Add(Name:="tempStyle", Type:=wdStyleTypeCharacter)
    End If
    On Error GoTo 0

    ' הגדרות גופן
    With tempStyle.Font
        .Name = "Calibri"
        .NameBi = "Calibri"
    End With

    ' החלת הסגנון על הטווח
    rng.style = tempStyle

    ActiveDocument.Styles("tempStyle").Delete
End Function

Function ApplyDesign(paraRange As Range)
        
    Dim tempoRange As Range
        
    Set tempoRange = paraRange.Duplicate
    With tempoRange
        .Collapse wdCollapseStart
        .MoveUntil " " & Chr(13)
        .Move wdCharacter, 1
        .MoveEnd wdCharacter, 1
    End With
    
    With paraRange.Font
        If sizeOption = 1 Then
            .Size = (1 + (pointsPercentage / 100)) * tempoRange.Font.Size
            .SizeBi = (1 + (pointsPercentage / 100)) * tempoRange.Font.SizeBi
        ElseIf sizeOption = 2 Then
            .Size = tempoRange.Font.Size
            .SizeBi = tempoRange.Font.SizeBi
        End If
        
        If Not FontOption Then
            .Name = tempoRange.Font.Name
            .NameAscii = tempoRange.Font.NameAscii
            .NameFarEast = tempoRange.Font.NameFarEast
            .NameOther = tempoRange.Font.NameOther
            .NameBi = tempoRange.Font.NameBi
        End If

        If Not ColorOption Then
            .Color = tempoRange.Font.Color
            .TextColor.RGB = tempoRange.Font.TextColor.RGB
        End If
        
        If Not strikethroughOption Then
            .strikethrough = tempoRange.Font.strikethrough
            .DoubleStrikeThrough = tempoRange.Font.DoubleStrikeThrough
        End If
        
        If Not BoldOption Then
            .Bold = tempoRange.Font.Bold
            .BoldBi = tempoRange.Font.BoldBi
        End If
                
        If Not ItalicOption Then
            .Italic = tempoRange.Font.Italic
            .ItalicBi = tempoRange.Font.ItalicBi
        End If
        If Not UnderlineOption Then
            .Underline = tempoRange.Font.Underline
            .UnderlineColor = tempoRange.Font.UnderlineColor
        End If
        If Not FillOption Then
            .Shadow = tempoRange.Font.Shadow
            .Outline = tempoRange.Font.Outline
            .Emboss = tempoRange.Font.Emboss
            .Engrave = tempoRange.Font.Engrave
        End If
        If Not EffectsOption Then
            .Superscript = tempoRange.Font.Superscript
            .Subscript = tempoRange.Font.Subscript
            .Hidden = tempoRange.Font.Hidden
            .SmallCaps = tempoRange.Font.SmallCaps
            .AllCaps = tempoRange.Font.AllCaps
        End If
        If Not advancedOption Then
            .Spacing = tempoRange.Font.Spacing
            .Position = tempoRange.Font.Position
            .Kerning = tempoRange.Font.Kerning
'            .KerningBy = tempoRange.Font.KerningBy
            .Scaling = tempoRange.Font.Scaling
        End If
    End With
End Function
-------------------------------------------------------------------------------
VBA MACRO FontDialogueHelper.bas 
in file: word/vbaProject.bin - OLE stream: 'VBA/FontDialogueHelper'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit
Option Private Module
Public fontFormatDialog As Dialog

Sub ChangeFontDialogSettings()
    LoadFontDialogSettings
    fontFormatDialog.Display
    SaveFontDialogSettings
End Sub


Sub SaveFontDialogSettings()
    If fontFormatDialog Is Nothing Then Set fontFormatDialog = Dialogs(wdDialogFormatFont)
    
    On Error Resume Next
    With fontFormatDialog
        SettingsHelper.Save RibbonControl.appName, "FontDialogSettings", "Points", .points
        SettingsHelper.Save RibbonControl.appName, "FontDialogSettings", "Underline", .Underline
        SettingsHelper.Save RibbonControl.appName, "FontDialogSettings", "Color", .Color
        SettingsHelper.Save RibbonControl.appName, "FontDialogSettings", "StrikeThrough", .strikethrough
        SettingsHelper.Save RibbonControl.appName, "FontDialogSettings", "Superscript", .Superscript
        SettingsHelper.Save RibbonControl.appName, "FontDialogSettings", "Subscript", .Subscript
        SettingsHelper.Save RibbonControl.appName, "FontDialogSettings", "Hidden", .Hidden
        SettingsHelper.Save RibbonControl.appName, "FontDialogSettings", "SmallCaps", .SmallCaps
        SettingsHelper.Save RibbonControl.appName, "FontDialogSettings", "AllCaps", .AllCaps
        SettingsHelper.Save RibbonControl.appName, "FontDialogSettings", "Spacing", .Spacing
        SettingsHelper.Save RibbonControl.appName, "FontDialogSettings", "Position", .Position
        SettingsHelper.Save RibbonControl.appName, "FontDialogSettings", "Kerning", .Kerning
        SettingsHelper.Save RibbonControl.appName, "FontDialogSettings", "KerningMin", .KerningMin
        SettingsHelper.Save RibbonControl.appName, "FontDialogSettings", "Default", .Default
        SettingsHelper.Save RibbonControl.appName, "FontDialogSettings", "Tab", .Tab
        SettingsHelper.Save RibbonControl.appName, "FontDialogSettings", "Font", .Font
        SettingsHelper.Save RibbonControl.appName, "FontDialogSettings", "Bold", .Bold
        SettingsHelper.Save RibbonControl.appName, "FontDialogSettings", "Italic", .Italic
        SettingsHelper.Save RibbonControl.appName, "FontDialogSettings", "DoubleStrikeThrough", .DoubleStrikeThrough
        SettingsHelper.Save RibbonControl.appName, "FontDialogSettings", "Shadow", .Shadow
        SettingsHelper.Save RibbonControl.appName, "FontDialogSettings", "Outline", .Outline
        SettingsHelper.Save RibbonControl.appName, "FontDialogSettings", "Emboss", .Emboss
        SettingsHelper.Save RibbonControl.appName, "FontDialogSettings", "Engrave", .Engrave
        SettingsHelper.Save RibbonControl.appName, "FontDialogSettings", "Scale", .Scale
        SettingsHelper.Save RibbonControl.appName, "FontDialogSettings", "Animations", .Animations
        SettingsHelper.Save RibbonControl.appName, "FontDialogSettings", "CharAccent", .CharAccent
        SettingsHelper.Save RibbonControl.appName, "FontDialogSettings", "FontMajor", .FontMajor
        SettingsHelper.Save RibbonControl.appName, "FontDialogSettings", "FontLowAnsi", .FontLowAnsi
        SettingsHelper.Save RibbonControl.appName, "FontDialogSettings", "FontHighAnsi", .FontHighAnsi
        SettingsHelper.Save RibbonControl.appName, "FontDialogSettings", "CharacterWidthGrid", .CharacterWidthGrid
        SettingsHelper.Save RibbonControl.appName, "FontDialogSettings", "ColorRGB", .ColorRGB
        SettingsHelper.Save RibbonControl.appName, "FontDialogSettings", "UnderlineColor", .UnderlineColor
        SettingsHelper.Save RibbonControl.appName, "FontDialogSettings", "PointsBi", .PointsBi
        SettingsHelper.Save RibbonControl.appName, "FontDialogSettings", "ColorBi", .ColorBi
        SettingsHelper.Save RibbonControl.appName, "FontDialogSettings", "FontNameBi", .FontNameBi
        SettingsHelper.Save RibbonControl.appName, "FontDialogSettings", "BoldBi", .BoldBi
        SettingsHelper.Save RibbonControl.appName, "FontDialogSettings", "ItalicBi", .ItalicBi
        SettingsHelper.Save RibbonControl.appName, "FontDialogSettings", "DiacColor", .DiacColor
    End With
    On Error GoTo 0
End Sub

Sub LoadFontDialogSettings()
    If fontFormatDialog Is Nothing Then Set fontFormatDialog = Dialogs(wdDialogFormatFont)
    
    On Error Resume Next
    With fontFormatDialog
        .points = GetSavedSetting(RibbonControl.appName, "FontDialogSettings", "Points", Nothing)
        .Underline = GetSavedSetting(RibbonControl.appName, "FontDialogSettings", "Underline", Nothing)
        .Color = GetSavedSetting(RibbonControl.appName, "FontDialogSettings", "Color", Nothing)
        .strikethrough = GetSavedSetting(RibbonControl.appName, "FontDialogSettings", "StrikeThrough", Nothing)
        .Superscript = GetSavedSetting(RibbonControl.appName, "FontDialogSettings", "Superscript", Nothing)
        .Subscript = GetSavedSetting(RibbonControl.appName, "FontDialogSettings", "Subscript", Nothing)
        .Hidden = GetSavedSetting(RibbonControl.appName, "FontDialogSettings", "Hidden", Nothing)
        .SmallCaps = GetSavedSetting(RibbonControl.appName, "FontDialogSettings", "SmallCaps", Nothing)
        .AllCaps = GetSavedSetting(RibbonControl.appName, "FontDialogSettings", "AllCaps", Nothing)
        .Spacing = GetSavedSetting(RibbonControl.appName, "FontDialogSettings", "Spacing", Nothing)
        .Position = GetSavedSetting(RibbonControl.appName, "FontDialogSettings", "Position", Nothing)
        .Kerning = GetSavedSetting(RibbonControl.appName, "FontDialogSettings", "Kerning", Nothing)
        .KerningMin = GetSavedSetting(RibbonControl.appName, "FontDialogSettings", "KerningMin", Nothing)
        .Default = GetSavedSetting(RibbonControl.appName, "FontDialogSettings", "Default", Nothing)
        .Tab = GetSavedSetting(RibbonControl.appName, "FontDialogSettings", "Tab", Nothing)
        .Font = GetSavedSetting(RibbonControl.appName, "FontDialogSettings", "Font", Nothing)
        .Bold = GetSavedSetting(RibbonControl.appName, "FontDialogSettings", "Bold", Nothing)
        .Italic = GetSavedSetting(RibbonControl.appName, "FontDialogSettings", "Italic", Nothing)
        .DoubleStrikeThrough = GetSavedSetting(RibbonControl.appName, "FontDialogSettings", "DoubleStrikeThrough", Nothing)
        .Shadow = GetSavedSetting(RibbonControl.appName, "FontDialogSettings", "Shadow", Nothing)
        .Outline = GetSavedSetting(RibbonControl.appName, "FontDialogSettings", "Outline", Nothing)
        .Emboss = GetSavedSetting(RibbonControl.appName, "FontDialogSettings", "Emboss", Nothing)
        .Engrave = GetSavedSetting(RibbonControl.appName, "FontDialogSettings", "Engrave", Nothing)
        .Scale = GetSavedSetting(RibbonControl.appName, "FontDialogSettings", "Scale", Nothing)
        .Animations = GetSavedSetting(RibbonControl.appName, "FontDialogSettings", "Animations", Nothing)
        .CharAccent = GetSavedSetting(RibbonControl.appName, "FontDialogSettings", "CharAccent", Nothing)
        .FontMajor = GetSavedSetting(RibbonControl.appName, "FontDialogSettings", "FontMajor", Nothing)
        .FontLowAnsi = GetSavedSetting(RibbonControl.appName, "FontDialogSettings", "FontLowAnsi", Nothing)
        .FontHighAnsi = GetSavedSetting(RibbonControl.appName, "FontDialogSettings", "FontHighAnsi", Nothing)
        .CharacterWidthGrid = GetSavedSetting(RibbonControl.appName, "FontDialogSettings", "CharacterWidthGrid", Nothing)
        .ColorRGB = GetSavedSetting(RibbonControl.appName, "FontDialogSettings", "ColorRGB", Nothing)
        .UnderlineColor = GetSavedSetting(RibbonControl.appName, "FontDialogSettings", "UnderlineColor", Nothing)
        .PointsBi = GetSavedSetting(RibbonControl.appName, "FontDialogSettings", "PointsBi", Nothing)
        .ColorBi = GetSavedSetting(RibbonControl.appName, "FontDialogSettings", "ColorBi", Nothing)
        .FontNameBi = GetSavedSetting(RibbonControl.appName, "FontDialogSettings", "FontNameBi", Nothing)
        .BoldBi = GetSavedSetting(RibbonControl.appName, "FontDialogSettings", "BoldBi", Nothing)
        .ItalicBi = GetSavedSetting(RibbonControl.appName, "FontDialogSettings", "ItalicBi", Nothing)
        .DiacColor = GetSavedSetting(RibbonControl.appName, "FontDialogSettings", "DiacColor", Nothing)
    End With
    On Error GoTo 0
End Sub

-------------------------------------------------------------------------------
VBA MACRO CenterLastLine.bas 
in file: word/vbaProject.bin - OLE stream: 'VBA/CenterLastLine'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit
Option Private Module

Sub Execute(ByVal paraCollection As Collection, skip2Lines As Boolean)

    Dim selectionRange As Range
    Dim paraRange As Range
    Dim i As Integer
    ' Begin undo recording
    On Error Resume Next
    Application.UndoRecord.StartCustomRecord "מירכוז שורה אחרונה"
    Application.ScreenUpdating = False
    On Error GoTo Ending
    
    FormRunning.LblCenterLastLine.Caption = "אוסף מידע עבור מירכוז שורה אחרונה"
    FormRunning.LblCenterLastLine.Font.Bold = True: DoEvents
    ' Set selection range
    Set selectionRange = Selection.Range
    Call remove(paraCollection)
        
    If skip2Lines Then
         Set paraCollection = Helpers.GetValidParagraphRangesForLastLine(paraCollection, 3)
    Else
        Set paraCollection = Helpers.GetValidParagraphRangesForLastLine(paraCollection, 2)
    End If

    Application.ActiveDocument.DefaultTabStop = 0
    
    ' Loop through paragraphs
    For i = 1 To paraCollection.Count
        Set paraRange = paraCollection(i)
        
        FormRunning.LblCenterLastLine.Caption = "ממרכז שורה אחרונה " & i & " מתוך " & paraCollection.Count
        If i Mod -Int(-paraCollection.Count / 40) = 0 Then DoEvents
        If stopCode Then GoTo Ending

        With paraRange
            .ParagraphFormat.Alignment = wdAlignParagraphDistribute
            .End = .End - 1
            .InsertAfter Chr(9)
        End With
        
    Next i

    FormRunning.LblCenterLastLine.Caption = "מירכוז שורה אחרונה - הושלם": DoEvents

Ending:
    ' End undo recording
    Application.UndoRecord.EndCustomRecord
    Application.ScreenUpdating = True
    selectionRange.Select
End Sub

Sub remove(Optional ByVal paraCollection As Collection = Nothing)
    Dim paraRange As Range
    Dim tempoRange As Range
    Dim i As Integer
    
    On Error GoTo Ending
    Application.UndoRecord.StartCustomRecord "הסרת מירכוז שורה אחרונה"
    Application.ScreenUpdating = False
    
    If paraCollection Is Nothing Then
        Set paraCollection = Helpers.GetAllSelectedParagraphRanges
    End If
    
    For i = 1 To paraCollection.Count
        FormRemoveRunning.LblCenterLastLine.Caption = "מסיר מירכוז שורה אחרונה " & i & " מתוך " & paraCollection.Count
        If i Mod -Int(-paraCollection.Count / 40) = 0 Then DoEvents
        If stopCode Then GoTo Ending
        
        Set paraRange = paraCollection(i)
        
        paraRange.ParagraphFormat.TabStops.ClearAll
        Set tempoRange = paraRange.Duplicate
        With tempoRange.Find
            .ParagraphFormat.Alignment = wdAlignParagraphDistribute
            .text = "^t^p"
            .Wrap = wdFindStop
            If .Execute Then
                tempoRange.Characters.First.Delete
                paraRange.ParagraphFormat.Alignment = wdAlignParagraphJustify
            End If
        End With
        Set tempoRange = paraRange.Duplicate
        With tempoRange.Find
            .text = "^l^t"
            .Replacement.text = ""
            .Wrap = wdFindStop
            If .Execute Then tempoRange.Delete
        End With
    Next i
Ending:
    Application.UndoRecord.EndCustomRecord
    Application.ScreenUpdating = True
End Sub
-------------------------------------------------------------------------------
VBA MACRO FormatFirstWord.bas 
in file: word/vbaProject.bin - OLE stream: 'VBA/FormatFirstWord'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit
Option Private Module

Sub ByHand(ByVal paraCollection As Collection, dinami As Boolean, alignmentUpwards As Boolean)

    Dim selectionRange As Range
    Dim paraRange As Range
    Dim i As Integer
    
    ' טעינת הדיאלוג
    If FontDialogueHelper.fontFormatDialog Is Nothing Then _
         FontDialogueHelper.LoadFontDialogSettings
    ' טעינת הדיאלוג הדינמי
    If dinami Then FontDinamiHelper.GetFirstWordDinamiOption
    
    ' Begin undo recording
    On Error Resume Next
    Application.UndoRecord.StartCustomRecord "עיצוב מילה ראשונה"
    Application.ScreenUpdating = False
    On Error GoTo Ending
    
    FormRunning.LblFirstWord.Caption = "אוסף מידע עבור עיצוב מילה ראשונה"
    FormRunning.LblFirstWord.Font.Bold = True: DoEvents
    
    ' Set selection range
    Set selectionRange = Selection.Range
    Set paraCollection = Helpers.GetValidParagraphRangesForFirstWord(paraCollection)
    
    ' Loop through paragraphs
    For i = 1 To paraCollection.Count
        Set paraRange = paraCollection(i).Duplicate
        
        FormRunning.LblFirstWord.Caption = "מעצב מילה ראשונה " & i & " מתוך " & paraCollection.Count
        If i Mod -Int(-paraCollection.Count / 40) = 0 Then DoEvents
        If stopCode Then GoTo Ending
        With paraRange
            .Collapse
            .MoveEndUntil Cset:=" "
            .Select
        End With
        
        FontDialogueHelper.fontFormatDialog.Execute
        
        If dinami Then Call FontDinamiHelper.ApplyDesign(paraRange)
        
        If alignmentUpwards Then paraRange.Font.Position = Helpers.newPosition(paraRange)
    Next i
    
    FormRunning.LblFirstWord.Caption = "עיצוב מילה ראשונה - הושלם": DoEvents
    ' Restore original selection
    selectionRange.Select
        
Ending:
    ' End undo recording
    Application.UndoRecord.EndCustomRecord
    Application.ScreenUpdating = True
End Sub

Sub ByStyle(ByVal paraCollection As Collection, dinami As Boolean, alignmentUpwards As Boolean)

    Dim selectionRange As Range
    Dim paraRange As Range
    Dim firstWordStyle As style
    Dim i As Integer
    
    ' Begin undo recording
    On Error Resume Next
    Application.UndoRecord.StartCustomRecord "עיצוב מילה ראשונה"
    Application.ScreenUpdating = False
    On Error GoTo Ending
    
    FormRunning.LblFirstWord.Caption = "אוסף מידע עבור עיצוב מילה ראשונה"
    FormRunning.LblFirstWord.Font.Bold = True: DoEvents
    
    ' Set selection range
    Set selectionRange = Selection.Range
    Set paraCollection = Helpers.GetValidParagraphRangesForFirstWord(paraCollection)
    
    'prepare style
    Set firstWordStyle = Helpers.GetFirstWordStyle()
    If dinami Then FontDinamiHelper.GetFirstWordDinamiOption
    
    ' Loop through paragraphs
    For i = 1 To paraCollection.Count
        Set paraRange = paraCollection(i).Duplicate
        FormRunning.LblFirstWord.Caption = "מעצב מילה ראשונה " & i & " מתוך " & paraCollection.Count
        If i Mod -Int(-paraCollection.Count / 40) = 0 Then DoEvents
        If stopCode Then GoTo Ending
        
        With paraRange
            .Collapse
            .MoveEndUntil Cset:=" "
            .Font.Reset
            .style = firstWordStyle
            If alignmentUpwards Then .Font.Position = Helpers.newPosition(paraRange)
            .Characters.First.Select
        End With
        
        If dinami Then Call FontDinamiHelper.ApplyDesign(paraRange)
        If alignmentUpwards Then paraRange.Font.Position = Helpers.newPosition(paraRange)
        
    Next i
    
    FormRunning.LblFirstWord.Caption = "עיצוב מילה ראשונה - הושלם": DoEvents
     
    ' Restore original selection
    selectionRange.Select
        
Ending:
    ' End undo recording
    Application.UndoRecord.EndCustomRecord
    Application.ScreenUpdating = True
End Sub

Sub remove(Optional ByVal paraCollection As Collection = Nothing)
    Dim selectionRange As Range
    Dim paraRange As Range
    Dim txt As String
    Dim i As Integer
    
    ' Begin undo recording
    On Error Resume Next
    Application.UndoRecord.StartCustomRecord "הסרת עיצוב מילה ראשונה"
    Application.ScreenUpdating = False
    On Error GoTo Ending
    
    FormRemoveRunning.LblFirstWord.Caption = "אוסף מידע עבור הסרת עיצוב מילה ראשונה"
    FormRemoveRunning.LblFirstWord.Font.Bold = True
    
    'set range
    Set selectionRange = Selection.Range
    If paraCollection Is Nothing Then
        Set paraCollection = Helpers.GetAllSelectedParagraphRanges
    End If
    
    ' Loop through paragraphs
    For i = 1 To paraCollection.Count
        Set paraRange = paraCollection(i).Duplicate
        
        FormRemoveRunning.LblFirstWord.Caption = "מסיר עיצוב מילה ראשונה " & i & " מתוך " & paraCollection.Count
        If i Mod -Int(-paraCollection.Count / 40) = 0 Then DoEvents
        If stopCode Then GoTo Ending
            
        With paraRange
            .Collapse
            .MoveEndUntil Cset:=" "
            .MoveEnd
             txt = .text
            .text = ""
            .Select
            .text = txt
        End With
        
    Next i
    
    FormRemoveRunning.LblFirstWord.Caption = "הסרת עיצוב מילה ראשונה - הושלם"
    
    'reset selection
    selectionRange.Select
    
Ending:
    ' End undo recording
    Application.UndoRecord.EndCustomRecord
    Application.ScreenUpdating = True
End Sub
-------------------------------------------------------------------------------
VBA MACRO HangingFirstWord.bas 
in file: word/vbaProject.bin - OLE stream: 'VBA/HangingFirstWord'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit
Option Private Module

Sub Execute(ByVal paraCollection As Collection, skip2Lines As Boolean)
    
    Dim selectionRange As Range
    Dim paraRange As Range
    Dim firstWordSpacing As Double
    Dim i As Integer
    
    ' Begin undo recording
    On Error Resume Next
    Application.UndoRecord.StartCustomRecord "עיצוב חלון"
    On Error GoTo Ending
    
    FormRunning.LblHangingFirstWord.Caption = "אוסף מידע עבור עיצוב חלון"
    FormRunning.LblHangingFirstWord.Font.Bold = True: DoEvents
    
    ' Set selection range
    Set selectionRange = Selection.Range
    Call remove(paraCollection)
    
    If skip2Lines = True Then
         Set paraCollection = Helpers.GetValidParagraphRanges(paraCollection, 3)
    Else
        Set paraCollection = Helpers.GetValidParagraphRanges(paraCollection, 2)
    End If
   
    
    If selectionRange.Information(wdInFootnote) Or selectionRange.Information(wdInEndnote) Then
        Call Helpers.PrepareFootnotes(paraCollection)
    End If
        
    ' Loop through paragraphs
    For i = 1 To paraCollection.Count
        Set paraRange = paraCollection(i).Duplicate
        
        FormRunning.LblHangingFirstWord.Caption = "מעצב חלון " & i & " מתוך " & paraCollection.Count: DoEvents
        If stopCode Then GoTo Ending
        
        With paraRange
            .Collapse
            .MoveUntil " "
            .Move
            firstWordSpacing = .Information(wdHorizontalPositionRelativeToTextBoundary)
            .Select
        End With
            
        Selection.EndKey

        With Selection.Range
            .text = vbVerticalTab & Chr(160)
            .Collapse (wdCollapseEnd)
'            Application.ScreenRefresh
            .Previous.Font.Spacing = .Information(wdHorizontalPositionRelativeToTextBoundary) - firstWordSpacing
        End With
                        
         
    Next i
    
    FormRunning.LblHangingFirstWord.Caption = "עיצוב חלון - הושלם": DoEvents
    ' Restore original selection
    selectionRange.Select
    
Ending:
    ' End undo recording
    Application.UndoRecord.EndCustomRecord
End Sub

Sub DoubleWindow(ByVal paraCollection As Collection)
    
    Dim selectionRange As Range
    Dim paraRange As Range
    Dim firstWordSpacing As Double
    Dim currentSpacing As Double
    Dim i As Integer
    
    ' Begin undo recording
    On Error Resume Next
    Application.UndoRecord.StartCustomRecord "עיצוב חלון"
    On Error GoTo Ending
    
    FormRunning.LblHangingFirstWord.Caption = "אוסף מידע עבור עיצוב חלון"
    FormRunning.LblHangingFirstWord.Font.Bold = True: DoEvents
    
    ' Set selection range
    Set selectionRange = Selection.Range
    Call remove(paraCollection)
    Set paraCollection = Helpers.GetValidParagraphRanges(paraCollection, 3)
    
    If selectionRange.Information(wdInFootnote) Or selectionRange.Information(wdInEndnote) Then
        Call Helpers.PrepareFootnotes(paraCollection)
    End If
        
    ' Loop through paragraphs
    For i = 1 To paraCollection.Count
        Set paraRange = paraCollection(i).Duplicate
        
        FormRunning.LblHangingFirstWord.Caption = "מעצב חלון " & i & " מתוך " & paraCollection.Count: DoEvents
        If stopCode Then GoTo Ending
            
        With paraRange
            .Collapse
            .MoveUntil " "
            .Move
             firstWordSpacing = .Information(wdHorizontalPositionRelativeToTextBoundary)
            .Select
        End With

        With Selection
            .EndKey
            .text = vbVerticalTab & Chr(160)
            .Collapse (wdCollapseEnd)
'            Application.ScreenRefresh
            firstWordSpacing = .Information(wdHorizontalPositionRelativeToTextBoundary) - firstWordSpacing
            .Previous.Font.Spacing = firstWordSpacing
            .EndKey
            .text = vbVerticalTab & Chr(160)
            .Collapse (wdCollapseEnd)
            .Previous.Font.Spacing = firstWordSpacing
        End With
                 
        
    Next i

    FormRunning.LblHangingFirstWord.Caption = "עיצוב חלון - הושלם": DoEvents
    ' Restore original selection
    selectionRange.Select
    
Ending:
    ' End undo recording
    Application.UndoRecord.EndCustomRecord
End Sub
Sub DoubleWindowThreeLines(ByVal paraCollection As Collection, skip2Lines As Boolean)
    
    Dim selectionRange As Range
    Dim paraRange As Range
    Dim firstWordSpacing As Double
    Dim currentSpacing As Double
    Dim linesCount As Integer
    Dim i As Integer
    
    ' Begin undo recording
    On Error Resume Next
    Application.UndoRecord.StartCustomRecord "עיצוב חלון"
    On Error GoTo Ending
    
    FormRunning.LblHangingFirstWord.Caption = "אוסף מידע עבור עיצוב חלון"
    FormRunning.LblHangingFirstWord.Font.Bold = True: DoEvents
    
    ' Set selection range
    Set selectionRange = Selection.Range
    Call remove(paraCollection)
    If skip2Lines = True Then
        Set paraCollection = Helpers.GetValidParagraphRanges(paraCollection, 3)
    Else
        Set paraCollection = Helpers.GetValidParagraphRanges(paraCollection, 2)
    End If
    
    If selectionRange.Information(wdInFootnote) Or selectionRange.Information(wdInEndnote) Then
        Call Helpers.PrepareFootnotes(paraCollection)
    End If
        
    ' Loop through paragraphs
    For i = 1 To paraCollection.Count
        Set paraRange = paraCollection(i).Duplicate
        
        FormRunning.LblHangingFirstWord.Caption = "מעצב חלון " & i & " מתוך " & paraCollection.Count: DoEvents
        If stopCode Then GoTo Ending
            
        With paraRange
            linesCount = .ComputeStatistics(wdStatisticLines)
            .Collapse
            .MoveUntil " "
            .Move
            firstWordSpacing = .Information(wdHorizontalPositionRelativeToTextBoundary)
            .Select
        End With

        With Selection
            .EndKey
            .text = vbVerticalTab & Chr(160)
            .Collapse (wdCollapseEnd)
            firstWordSpacing = .Information(wdHorizontalPositionRelativeToTextBoundary) - firstWordSpacing
            .Previous.Font.Spacing = firstWordSpacing
            If linesCount = 3 Then
                .EndKey
                .text = vbVerticalTab & Chr(160)
                .Collapse (wdCollapseEnd)
                .Previous.Font.Spacing = firstWordSpacing
                linesCount = paraRange.Paragraphs(1).Range.ComputeStatistics(wdStatisticLines)
                If linesCount > 3 Then
                    .MoveStart wdCharacter, -2
                    .Delete
                End If
            End If
            
        End With
                             
    Next i

    FormRunning.LblHangingFirstWord.Caption = "עיצוב חלון - הושלם": DoEvents
    ' Restore original selection
    selectionRange.Select
    
Ending:
    ' End undo recording
    Application.UndoRecord.EndCustomRecord
End Sub

Public Sub remove(Optional ByVal paraCollection As Collection = Nothing)
    Dim tempoRange As Range
    Dim paraRange As Range
    Dim i As Integer
    
    On Error GoTo Ending
    Application.UndoRecord.StartCustomRecord "הסרת עיצוב חלון"
    Application.ScreenUpdating = False

    If paraCollection Is Nothing Then Set paraCollection = Helpers.GetAllSelectedParagraphRanges
    
    For i = 1 To paraCollection.Count
        Set paraRange = paraCollection(i).Duplicate
        
        FormRemoveRunning.LblHangingFirstWord.Caption = "מסיר עיצוב חלון " & i & " מתוך " & paraCollection.Count
        If i Mod -Int(-paraCollection.Count / 40) = 0 Then DoEvents
        If stopCode Then GoTo Ending
        
        Set paraRange = paraCollection(i)
        Do
            Set tempoRange = paraRange.Duplicate
            With tempoRange.Find
                .text = vbVerticalTab & Chr(160)
                .Replacement.text = ""
                .Wrap = wdFindStop
                If .Execute Then tempoRange.Delete Else: Exit Do
            End With
        Loop
    Next i
Ending:
    Application.UndoRecord.EndCustomRecord
    Application.ScreenUpdating = True
End Sub
-------------------------------------------------------------------------------
VBA MACRO Helpers.bas 
in file: word/vbaProject.bin - OLE stream: 'VBA/Helpers'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit
Option Private Module

Sub PrepareFootnotes(ByVal paraCollection As Collection)
    Dim paraRange As Range
    
    For Each paraRange In paraCollection
        With paraRange.Find
            .text = "^f"
            .Replacement.text = "^&ֵ%%" & Chr(160)
            .Wrap = wdFindStop
            .Execute Replace:=wdReplaceAll
        End With
        
         With paraRange.Find
            .text = "%%" & Chr(160) & " "
            .Replacement.text = Chr(160)
            .Wrap = wdFindStop
            .Execute Replace:=wdReplaceAll
        End With
    Next paraRange
End Sub
Function GetAllSelectedParagraphRanges() As Collection
    
    Dim rng As Range
    Dim para As Paragraph
    Dim paraCollection As New Collection

    Set rng = Selection.Range
    For Each para In rng.Paragraphs
        paraCollection.Add para.Range
    Next para
    
    Set GetAllSelectedParagraphRanges = paraCollection

End Function
Function GetNonHeadingSelectedParagraphRanges() As Collection
    
    Dim rng As Range
    Dim para As Paragraph
    Dim paraCollection As New Collection
    Dim i As Integer
    
    Set rng = Selection.Range
    For i = 1 To rng.Paragraphs.Count
        FormRunning.Label7.Caption = "אוסף מידע " & i & " מתוך " & rng.Paragraphs.Count
        FormRemoveRunning.Label7.Caption = "אוסף מידע " & i & " מתוך " & rng.Paragraphs.Count
        DoEvents
        If stopCode Then GoTo Ending
        Set para = rng.Paragraphs(i)
        If para.OutlineLevel = 10 And _
            para.Alignment <> wdAlignParagraphCenter _
            Then
                paraCollection.Add para.Range
        End If
    Next i
Ending:
    Set GetNonHeadingSelectedParagraphRanges = paraCollection

End Function
Function GetSpecificSelectedParagraphRanges() As Collection
    
    Dim rng As Range
    Dim para As Paragraph
    Dim paraCollection As New Collection
    Dim stlName(0 To 14) As String
    Dim i As Integer
    
    Set rng = Selection.Range
    
    For i = LBound(stlName) To UBound(stlName)
        stlName(i) = SettingsHelper.GetSavedSetting(appName, "ParagraphFormat", "stlName" & i, "")
    Next i
    
    For i = 1 To rng.Paragraphs.Count
        FormRunning.Label7.Caption = "אוסף מידע " & i & " מתוך " & rng.Paragraphs.Count
        FormRemoveRunning.Label7.Caption = "אוסף מידע " & i & " מתוך " & rng.Paragraphs.Count
        DoEvents
        If stopCode Then GoTo Ending
        Set para = rng.Paragraphs(i)
        If IsStyleInArray(para, stlName) Then
            paraCollection.Add para.Range
        End If
    Next i
    
Ending:
    Set GetSpecificSelectedParagraphRanges = paraCollection

End Function
Function GetNonSpecificSelectedParagraphRanges() As Collection
    
    Dim rng As Range
    Dim para As Paragraph
    Dim paraCollection As New Collection
    Dim stlName(0 To 14) As String
    Dim i As Integer
    
    Set rng = Selection.Range
    
    For i = LBound(stlName) To UBound(stlName)
        stlName(i) = SettingsHelper.GetSavedSetting(appName, "ParagraphFormat", "stlName" & i, "")
    Next i
    
    For i = 1 To rng.Paragraphs.Count
        FormRunning.Label7.Caption = "אוסף מידע " & i & " מתוך " & rng.Paragraphs.Count
        FormRemoveRunning.Label7.Caption = "אוסף מידע " & i & " מתוך " & rng.Paragraphs.Count
        DoEvents
        If stopCode Then GoTo Ending
        Set para = rng.Paragraphs(i)
        If Not IsStyleInArray(para, stlName) Then
            paraCollection.Add para.Range
        End If
    Next i
    
Ending:
    Set GetNonSpecificSelectedParagraphRanges = paraCollection

End Function
Function GetValidParagraphRanges(ByVal prevParaCollection As Collection, targetLineCount As Integer) As Collection
     Dim paraRange As Range
     Dim paraCollection As New Collection
     
     For Each paraRange In prevParaCollection
        If paraRange.ComputeStatistics(wdStatisticLines) >= targetLineCount And _
            paraRange.ParagraphFormat.Alignment <> wdAlignParagraphCenter _
            Then
               paraCollection.Add paraRange
        End If
    Next paraRange
    
    Set GetValidParagraphRanges = paraCollection
End Function
Function GetValidParagraphRangesForLastLine(ByVal prevParaCollection As Collection, targetLineCount As Integer) As Collection
    
    Dim paraRange As Range
    Dim paraLineCount As Integer
    Dim paraTargetLineCount As Integer
    Dim paraCollection As New Collection
    
    For Each paraRange In prevParaCollection
        paraLineCount = paraRange.ComputeStatistics(wdStatisticLines)
        ' בדיקה אם הטווח מכיל עיצוב חלון
        paraTargetLineCount = targetLineCount
        If paraRange.text Like "*" & Chr(11) & Chr(160) & "*" & Chr(11) & Chr(160) & "*" Then
            If targetLineCount < 4 Then paraTargetLineCount = 4
        ElseIf paraRange.text Like "*" & Chr(11) & Chr(160) & "*" Then
            If targetLineCount < 3 Then paraTargetLineCount = 3
        End If
        
        If paraLineCount >= paraTargetLineCount And _
            paraRange.ParagraphFormat.Alignment <> wdAlignParagraphCenter _
            Then
                paraCollection.Add paraRange
        End If
    Next paraRange
    
    Set GetValidParagraphRangesForLastLine = paraCollection
End Function

Function GetValidParagraphRangesForFirstWord(ByVal prevParaCollection As Collection) As Collection
    
    Dim paraRange As Range
    Dim paraCollection As New Collection
    Dim isFootNotes As Boolean
     
    If prevParaCollection(1).Information(wdInFootnote) Or prevParaCollection(1).Information(wdInEndnote) Then isFootNotes = True
    
    For Each paraRange In prevParaCollection
        With paraRange
            If .text Like "* *" Then
                
                If isFootNotes = True And Asc(.Characters.First) = 2 Then ' adjustment for footnote references
                    .MoveStartUntil " "
                    .MoveStart
                End If
                
                paraCollection.Add paraRange
                
            End If
        End With
    Next paraRange
    
    Set GetValidParagraphRangesForFirstWord = paraCollection
End Function
Function GetValidParagraphRangesForLineSpacing(ByVal prevParaCollection As Collection, ByVal isRemove As Boolean) As Collection
    
    Dim paraRange As Range
    Dim paraCollection As New Collection
    Dim isFootNotes As Boolean
     
    If prevParaCollection(1).Information(wdInFootnote) Or prevParaCollection(1).Information(wdInEndnote) Then isFootNotes = True
    
    For Each paraRange In prevParaCollection
        With paraRange
            If .text Like "* *" And ( _
                .ParagraphFormat.LineSpacingRule <> wdLineSpaceExactly Or ( _
                isRemove And .ParagraphFormat.LineSpacingRule = wdLineSpaceExactly)) Then
                
                    If isFootNotes = True And Asc(.Characters.First) = 2 Then ' adjustment for footnote references
                        .MoveStartUntil " "
                        .MoveStart
                    End If
                    
                    paraCollection.Add paraRange
                
            End If
        End With
    Next paraRange
    
    Set GetValidParagraphRangesForLineSpacing = paraCollection
End Function
Public Function GetFirstWordStyle() As style
    Dim targetStyle As style
    Dim targetStyleName As String
    
    targetStyleName = GetSavedSetting(RibbonControl.appName, "ParagraphFormat", "FirstWordStyle", "מילה ראשונה")
    If targetStyleName = "" Then targetStyleName = "מילה ראשונה"
    
    For Each targetStyle In Application.ActiveDocument.Styles
        If targetStyle.NameLocal = targetStyleName Then
            Set GetFirstWordStyle = targetStyle
            Exit Function
        End If
    Next
        
    Set targetStyle = Application.ActiveDocument.Styles.Add(targetStyleName, wdStyleTypeCharacter)
    With targetStyle.Font
        .Bold = 1
        .BoldBi = 1
        .Size = .Size + 2
        .SizeBi = .SizeBi + 2
    End With
    targetStyle.QuickStyle = True
    
    Set GetFirstWordStyle = targetStyle
End Function

Function newPosition(paraRange As Range) As Double
    
    Dim fontSize As Double
    Dim sizeAfter As Double
    
    fontSize = paraRange.Font.SizeBi
    sizeAfter = ActiveDocument.Range(paraRange.End, paraRange.End + 1).Font.SizeBi
    newPosition = -(fontSize - sizeAfter) / 2

End Function
Function GetFontProperties(paraRange As Range, ByRef fontName As String, ByRef fontSize As Single)
    Dim tempoRange As Range
    Set tempoRange = paraRange.Duplicate
    With tempoRange
        .Collapse wdCollapseStart
        .MoveUntil " " & Chr(13)
        .MoveEnd wdCharacter, 1
        fontName = .Font.NameBi
        fontSize = .Font.SizeBi
    End With
End Function
Function IsStyleInArray(para As Paragraph, stlNameArray() As String) As Boolean
    Dim stl As String
    Dim i As Integer
    
    For i = LBound(stlNameArray) To UBound(stlNameArray)
        stl = stlNameArray(i)
        If para.style.NameLocal = stl Then
            IsStyleInArray = True
            Exit Function
        End If
    Next i
End Function

-------------------------------------------------------------------------------
VBA MACRO LastLineBalance.bas 
in file: word/vbaProject.bin - OLE stream: 'VBA/LastLineBalance'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit
Option Private Module

Sub Repair(ByVal paraCollection As Collection, skip2Lines As Boolean)

    Dim selectionRange As Range
    Dim paraRange As Range
    Dim minLineRng As Range
    Dim minLineNum As Integer
    Dim i, x As Integer
    
    On Error GoTo Ending
    Application.UndoRecord.StartCustomRecord "איזון שורה אחרונה"
    
    FormRunning.LblBalanceLastLine.Caption = "אוסף מידע עבור איזון שורה אחרונה"
    FormRunning.LblBalanceLastLine.Font.Bold = True: DoEvents
    
    Set selectionRange = Selection.Range
    
    If skip2Lines Then
        Set paraCollection = Helpers.GetValidParagraphRangesForLastLine(paraCollection, 3)
    Else
        Set paraCollection = Helpers.GetValidParagraphRangesForLastLine(paraCollection, 2)
    End If
    
    For i = 1 To paraCollection.Count
        Set paraRange = paraCollection(i).Duplicate
        
        FormRunning.LblBalanceLastLine.Caption = "מאזן שורה אחרונה " & i & " מתוך " & paraCollection.Count: DoEvents
        If stopCode Then GoTo Ending
       
        If OneWord(paraRange) Then
        
            If paraRange.text Like "*" & Chr(11) & Chr(160) & "*" & Chr(11) & Chr(160) & "*" Then
                minLineNum = 3
            ElseIf paraRange.text Like "*" & Chr(11) & Chr(160) & "*" Then
                minLineNum = 2
            Else
                minLineNum = 1
            End If
            
            For x = 1 To 10
                Set minLineRng = MinLineSearch(paraRange, minLineNum)
                Call SpaceExpansion(minLineRng)
                If Not OneWord(paraRange) Then Exit For
            Next x
        
        End If
    Next i
    
    FormRunning.LblBalanceLastLine.Caption = "איזון שורה אחרונה - הושלם": DoEvents

Ending:
    Application.UndoRecord.EndCustomRecord
    selectionRange.Select
End Sub

Function OneWord(paraRange As Range) As Boolean
    
    paraRange.Select
    With Selection
        .Collapse direction:=wdCollapseEnd
        Do While Mid(ActiveDocument.Range(.Start - 1, .Start).text, 1, 1) = " " Or Mid(ActiveDocument.Range(.Start - 1, .Start).text, 1, 1) = Chr(13) Or Mid(ActiveDocument.Range(.Start - 1, .Start).text, 1, 1) = Chr(9)
            .Move wdCharacter, -1
        Loop
        .MoveStart wdLine, -1
            
        If .text Like "* *" Then
            OneWord = False
        Else
            OneWord = True
        End If
        .MoveStart wdCharacter, -2
        ' דילוג במקרה שיש תו מעבר שורה
        If .text Like "*" & Chr(11) & "*" Then OneWord = False
    End With
End Function

Function MinLineSearch(paraRange As Range, ByRef minLineNum As Integer) As Range
    
    Dim i As Integer
    Dim lineCount As Integer
    Dim spacious As Double
    Dim startChar As Double
    
    paraRange.Select
    
    lineCount = paraRange.ComputeStatistics(wdStatisticLines)
    spacious = 1000
    
    For i = minLineNum To lineCount - 1
        paraRange.Select
        With Selection
            .Collapse direction:=wdCollapseStart
            .Move wdLine, i - 1
            .MoveUntil " "
            startChar = .Information(wdHorizontalPositionRelativeToPage)
            .Move Unit:=wdCharacter
            
            If spacious > startChar - .Information(wdHorizontalPositionRelativeToPage) Then
                spacious = startChar - .Information(wdHorizontalPositionRelativeToPage)
                .MoveStart wdLine, -1
                .MoveEnd wdLine, 1
                .MoveEnd wdCharacter, -1
                Set MinLineSearch = .Range
                minLineNum = i
            End If
        End With
    Next i
End Function
Function SpaceExpansion(minLineRng As Range)
    
    Dim minLineCount As Integer
    Dim i As Single
    Dim char As Range
    
    If minLineRng Is Nothing Then Exit Function
    
    minLineCount = minLineRng.Words.Count
    
    For i = 0.5 To 50 Step 0.5
        
        For Each char In minLineRng.Characters
            
            With char
                If .text = " " Then .Font.Spacing = i
            End With
        
        Next char
            
        minLineRng.Select
        With Selection
            .Collapse direction:=wdCollapseStart
            .MoveEnd wdLine, 1
            .MoveEnd wdCharacter, -1
            If minLineCount > .Words.Count Then Exit For
        End With
    Next i
End Function
Sub remove(ByVal paraCollection As Collection)
    
    Dim selectionRange As Range
    Dim paraRange As Range
    Dim char As Range
    Dim i As Integer
    
    On Error GoTo Ending
    
    Application.UndoRecord.StartCustomRecord "הסרת איזון שורה אחרונה"
    Application.ScreenUpdating = False
    
    Set selectionRange = Selection.Range
    selectionRange.Expand Unit:=wdParagraph
    
    For i = 1 To paraCollection.Count
        FormRemoveRunning.LblBalanceLastLine.Caption = "מסיר איזון שורה אחרונה " & i & " מתוך " & paraCollection.Count
        If i Mod -Int(-paraCollection.Count / 40) = 0 Then DoEvents
        If stopCode Then GoTo Ending
        
        Set paraRange = paraCollection(i)
        
        For Each char In paraRange.Characters
            With char
                If .text = " " Then .Font.Spacing = ActiveDocument.Range(.Start - 1, .Start).Font.Spacing
            End With
        Next char
    Next i
    
Ending:
    Application.UndoRecord.EndCustomRecord
    Application.ScreenUpdating = True
    selectionRange.Select
End Sub
-------------------------------------------------------------------------------
VBA MACRO lineSpacing.bas 
in file: word/vbaProject.bin - OLE stream: 'VBA/LineSpacing'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit
Option Private Module

Sub Repair(ByVal paraCollection As Collection)

    Dim selectionRange As Range
    Dim paraRange As Range
    Dim lineSpacing As Double
    Dim fontName As String
    Dim fontSize As Single
    Dim savedSpaceBefore As Double
    Dim i As Integer
    
    ' Begin undo recording
    On Error Resume Next
    Application.UndoRecord.StartCustomRecord "תיקון מרווח שורות"
    On Error GoTo Ending
    On Error GoTo 0
    
    FormRunning.LblLineSpacing.Caption = "אוסף מידע עבור תיקון מרווח שורות"
    FormRunning.LblLineSpacing.Font.Bold = True: DoEvents

    ' Set selection range
    Set selectionRange = Selection.Range
    Set paraCollection = Helpers.GetValidParagraphRangesForLineSpacing(paraCollection, False)
    
    ' Loop through paragraphs
    For i = 1 To paraCollection.Count
        Set paraRange = paraCollection(i).Duplicate
        
        FormRunning.LblLineSpacing.Caption = "מתקן מרווח שורות " & i & " מתוך " & paraCollection.Count
        If i Mod -Int(-paraCollection.Count / 40) = 0 Then DoEvents
        If stopCode Then GoTo Ending
        
        With paraRange
            If Not .ParagraphFormat.LineSpacingRule = wdLineSpaceExactly Then
                savedSpaceBefore = .Paragraphs(1).SpaceBefore
                .Paragraphs(1).SpaceBefore = 0
                GetFontProperties paraRange, fontName, fontSize
                .Collapse wdCollapseStart
                .InsertBefore Chr(11)
                With .Font
                    .NameBi = fontName
                    .Name = fontName
                    .SizeBi = fontSize
                    .Size = fontSize
                End With
                .Collapse wdCollapseEnd
'                Application.ScreenRefresh
                lineSpacing = .Information(wdVerticalPositionRelativeToTextBoundary)
                If lineSpacing > 0 Then
                    .ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
                    .ParagraphFormat.lineSpacing = lineSpacing
                End If
                .MoveStart wdCharacter, -1
                .Delete
                .Paragraphs(1).SpaceBefore = savedSpaceBefore
            End If
        End With
    
    Next i

    FormRunning.LblLineSpacing.Caption = "תיקון מרווח שורות - הושלם": DoEvents
    ' Restore original selection
    selectionRange.Select
    
Ending:
    ' End undo recording
    Application.UndoRecord.EndCustomRecord
End Sub

Sub remove(ByVal paraCollection As Collection)

    Dim selectionRange As Range
    Dim paraRange As Range
    Dim lineSpacingExactly As Double
    Dim lineSpacingSingle As Double
    Dim lineSpacing As Double
    Dim fontName As String
    Dim fontSize As Single
    Dim i As Integer
    
    ' Begin undo recording
    On Error Resume Next
    Application.UndoRecord.StartCustomRecord "ביטול תיקון מרווח שורות"
    On Error GoTo Ending
    
    FormRemoveRunning.LblLineSpacing.Font.Bold = True
    FormRemoveRunning.LblLineSpacing.Caption = "אוסף מידע עבור הסרת תיקון מרווח שורות"
    Set paraCollection = Helpers.GetValidParagraphRangesForLineSpacing(paraCollection, True)
    'set range
    Set selectionRange = Selection.Range
    
    ' Loop through paragraphs
    For i = 1 To paraCollection.Count
        FormRemoveRunning.LblLineSpacing.Caption = "מסיר תיקון מרווח שורות " & i & " מתוך " & paraCollection.Count
        If i Mod -Int(-paraCollection.Count / 40) = 0 Then DoEvents
        If stopCode Then GoTo Ending
        
        Set paraRange = paraCollection(i).Duplicate
        
        FormRemoveRunning.LblLineSpacing.Caption = "מסיר תיקון מרווח שורות " & i & " מתוך " & paraCollection.Count: DoEvents
        If stopCode Then GoTo Ending
            
        With paraRange
            If .ParagraphFormat.LineSpacingRule = wdLineSpaceExactly Then
                GetFontProperties paraRange, fontName, fontSize
                lineSpacingExactly = .ParagraphFormat.lineSpacing
                .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
                .Collapse wdCollapseStart
                .text = Chr(11)
                With .Font
                    .NameBi = fontName
                    .Name = fontName
                    .SizeBi = fontSize
                    .Size = fontSize
                End With
                .Collapse wdCollapseStart
                .MoveEnd wdCharacter, 1
                .Collapse wdCollapseEnd
                lineSpacingSingle = .Information(wdVerticalPositionRelativeToTextBoundary) - .Paragraphs(1).SpaceBefore
                .MoveStart wdCharacter, -1
                .Delete
                lineSpacing = lineSpacingExactly / lineSpacingSingle
                If lineSpacing > 0 Then
                    .ParagraphFormat.LineSpacingRule = wdLineSpaceMultiple
                    .ParagraphFormat.lineSpacing = LinesToPoints(lineSpacing)
                End If
            End If
        End With
        
    Next i
    
    FormRemoveRunning.LblLineSpacing.Caption = "הסרת תיקון מרווח שורות - הושלם"
    'reset selection
    selectionRange.Select
    
Ending:
    ' End undo recording
    Application.UndoRecord.EndCustomRecord
End Sub

Sub alignmentUpwards()
    
    Dim paraCollection As Collection
    Dim selectionRange As Range
    Dim paraRange As Range
    Dim firstWordStyle As style

    ' Begin undo recording
    On Error Resume Next
    Application.UndoRecord.StartCustomRecord "יישור כלפי מעלה"
    On Error GoTo Ending
    
    ' Set selection range
    Set selectionRange = Selection.Range
    Set paraCollection = Helpers.GetValidParagraphRangesForFirstWord(selectionRange)
    
    ' Loop through paragraphs
    For Each paraRange In paraCollection
            
        With paraRange
            .Collapse
            .MoveEndUntil Cset:=" " & Chr(13)
            .style = newPosition(paraRange)
        End With
        
    Next paraRange
    
    ' Restore original selection
    selectionRange.Select
        
Ending:
    ' End undo recording
    Application.UndoRecord.EndCustomRecord
End Sub
-------------------------------------------------------------------------------
VBA MACRO RibbonControl.bas 
in file: word/vbaProject.bin - OLE stream: 'VBA/RibbonControl'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit
Option Private Module
Public myRibbon As IRibbonUI
Public appName As String
Public stopCode As Boolean

' Individual variables for each setting
Dim FirstWordCheckBox As Boolean
Dim FirstWordOption1 As Boolean
Dim FirstWordOption2 As Boolean

Dim FirstWordDinami As Boolean
Dim FirstWordLineSpacing As Boolean
Dim FirstWordAlignmentUpwards As Boolean

Dim HangingCheckBox As Boolean
Dim HangingOption1 As Boolean
Dim HangingOption2 As Boolean
Dim HangingOption3 As Boolean
Dim HangingSkip2Lines As Boolean

Dim LastLineCheckBox As Boolean
Dim LastLineCenterOption As Boolean
'Dim LastLineOption2 As Boolean
Dim LastLineBalanceOption As Boolean
Dim LastLineSkip2Lines As Boolean

' Callback for ribbon load
Sub OnLoad(ribbon As IRibbonUI)
    Set myRibbon = ribbon
    appName = "ShulchanHaorech"
    
    ' Initialize variables with defaults or load from registry
    FirstWordCheckBox = SettingsHelper.GetSavedSetting(appName, "ParagraphFormat", "FirstWordCheckBox", False)
    FirstWordOption1 = SettingsHelper.GetSavedSetting(appName, "ParagraphFormat", "FirstWordOption1", False)
    FirstWordOption2 = SettingsHelper.GetSavedSetting(appName, "ParagraphFormat", "FirstWordOption2", True)
    
    FirstWordDinami = SettingsHelper.GetSavedSetting(appName, "ParagraphFormat", "FirstWordDinami", False)
    FirstWordLineSpacing = SettingsHelper.GetSavedSetting(appName, "ParagraphFormat", "FirstWordLineSpacing", False)
    FirstWordAlignmentUpwards = SettingsHelper.GetSavedSetting(appName, "ParagraphFormat", "FirstWordAlignmentUpwards", False)
    
    HangingCheckBox = SettingsHelper.GetSavedSetting(appName, "ParagraphFormat", "HangingCheckBox", False)
    HangingOption1 = SettingsHelper.GetSavedSetting(appName, "ParagraphFormat", "HangingOption1", False)
    HangingOption2 = SettingsHelper.GetSavedSetting(appName, "ParagraphFormat", "HangingOption2", False)
    HangingOption3 = SettingsHelper.GetSavedSetting(appName, "ParagraphFormat", "HangingOption3", False)
    HangingSkip2Lines = SettingsHelper.GetSavedSetting(appName, "ParagraphFormat", "HangingSkip2Lines", False)
    
    LastLineCheckBox = SettingsHelper.GetSavedSetting(appName, "ParagraphFormat", "LastLineCheckBox", False)
    LastLineCenterOption = SettingsHelper.GetSavedSetting(appName, "ParagraphFormat", "LastLineCenterOption", False)
    LastLineBalanceOption = SettingsHelper.GetSavedSetting(appName, "ParagraphFormat", "LastLineBalanceOption", False)
    LastLineSkip2Lines = SettingsHelper.GetSavedSetting(appName, "ParagraphFormat", "LastLineSkip2Lines", False)
End Sub

Sub ApplyDesignButton_OnAction(control As IRibbonControl)
    
    Dim paraCollection As New Collection
    On Error GoTo Ending
    stopCode = False
         
    If Selection.Range.Paragraphs.Count > 25 Then
        Call formShow(FormRunning): DoEvents
    End If
    
    If stopCode Then GoTo Ending
    
    Select Case control.ID
        Case "ApplyNonHeading"
            Set paraCollection = Helpers.GetNonHeadingSelectedParagraphRanges
        Case "ApplyAll"
            Set paraCollection = Helpers.GetAllSelectedParagraphRanges
        Case "ApplyByStyle"
            FormSelectStyleParagraphFormat.Show
            If stopCode Then GoTo Ending
            Set paraCollection = Helpers.GetSpecificSelectedParagraphRanges
        Case "ApplyByNonStyle"
            FormSelectStyleParagraphFormat.Show
            If stopCode Then GoTo Ending
            Set paraCollection = Helpers.GetNonSpecificSelectedParagraphRanges
    End Select
        
    FormRunning.Label7.Caption = "איסוף מידע - הושלם"
    
    If FirstWordCheckBox Then
        If FirstWordLineSpacing Then Call lineSpacing.Repair(paraCollection)
        If FirstWordOption1 Then FormatFirstWord.ByHand paraCollection, FirstWordDinami, FirstWordAlignmentUpwards
        If FirstWordOption2 Then FormatFirstWord.ByStyle paraCollection, FirstWordDinami, FirstWordAlignmentUpwards
    End If
    
    If stopCode Then GoTo Ending
    
    If HangingCheckBox Then
        If HangingOption2 Then
            Call HangingFirstWord.DoubleWindow(paraCollection)
        ElseIf HangingOption3 Then
            Call HangingFirstWord.DoubleWindowThreeLines(paraCollection, HangingSkip2Lines)
        Else
            Call HangingFirstWord.Execute(paraCollection, HangingSkip2Lines)
        End If
    End If

    If stopCode Then GoTo Ending
    
    If LastLineCheckBox Then
        If LastLineCenterOption Then
            Call CenterLastLine.Execute(paraCollection, LastLineSkip2Lines)
        End If
        If LastLineBalanceOption Then Call LastLineBalance.Repair(paraCollection, LastLineSkip2Lines)
    End If

Ending:
    Unload FormRunning
End Sub
Sub RemoveDesignButton_OnAction(control As IRibbonControl)

    Dim paraCollection As New Collection
    On Error GoTo Ending
    stopCode = False
    
    ' הצגת טופס פעולה במקרה של יותר מ25 פיסקאות
    If Selection.Range.Paragraphs.Count > 25 Then
        Call formShow(FormRemoveRunning): DoEvents
    End If
    
    Select Case control.ID
        Case "RemoveAll"
            Set paraCollection = Helpers.GetAllSelectedParagraphRanges
        Case "RemoveByStyle"
            FormSelectStyleParagraphFormat.Show
            If stopCode Then GoTo Ending
            Set paraCollection = Helpers.GetSpecificSelectedParagraphRanges
        Case "RemoveByNonStyle"
            FormSelectStyleParagraphFormat.Show
            If stopCode Then GoTo Ending
            Set paraCollection = Helpers.GetNonSpecificSelectedParagraphRanges
    End Select
    
    FormRemoveRunning.Label7.Caption = "איסוף מידע - הושלם"
    ' הסרת עיצוב מילה ראשונה
    FormRemoveRunning.LblFirstWord.Font.Bold = True
    FormRemoveRunning.LblFirstWord.Caption = "מסיר עיצוב מילה ראשונה": DoEvents
    If FirstWordCheckBox Then Call FormatFirstWord.remove(paraCollection)
    FormRemoveRunning.LblFirstWord.Caption = "הסרת עיצוב מילה ראשונה - הושלם"
    
    ' הסרת תיקון מרווח שורות
    FormRemoveRunning.LblLineSpacing.Font.Bold = True
    FormRemoveRunning.LblLineSpacing.Caption = "מסיר תיקון מרווח שורות": DoEvents
    If FirstWordCheckBox And FirstWordLineSpacing Then Call lineSpacing.remove(paraCollection)
    FormRemoveRunning.LblLineSpacing.Caption = "הסרת תיקון מרווח שורות - הושלם"
    
    ' הסרת עיצוב חלון
    FormRemoveRunning.LblHangingFirstWord.Font.Bold = True
    FormRemoveRunning.LblHangingFirstWord.Caption = "מסיר עיצוב חלון": DoEvents
    If HangingCheckBox Then Call HangingFirstWord.remove(paraCollection)
    FormRemoveRunning.LblHangingFirstWord.Caption = "הסרת עיצוב חלון - הושלם"
    
    ' הסרת מירכוז שורה אחרונה
    FormRemoveRunning.LblCenterLastLine.Font.Bold = True
    FormRemoveRunning.LblCenterLastLine.Caption = "מסיר מירכוז שורה אחרונה": DoEvents
    If LastLineCheckBox And LastLineCenterOption Then Call CenterLastLine.remove(paraCollection)
    FormRemoveRunning.LblCenterLastLine.Caption = "הסרת מירכוז שורה אחרונה - הושלם"
    
    ' הסרת איזון שורה אחרונה
    FormRemoveRunning.LblBalanceLastLine.Font.Bold = True
    FormRemoveRunning.LblBalanceLastLine.Caption = "מסיר איזון שורה אחרונה": DoEvents
    If LastLineCheckBox And LastLineBalanceOption Then Call LastLineBalance.remove(paraCollection)
    FormRemoveRunning.LblBalanceLastLine.Caption = "הסרת איזון שורה אחרונה - הושלם"

Ending:
    Unload FormRemoveRunning
End Sub
' Callback for CheckBox onAction
Sub CheckBox_OnAction(control As IRibbonControl, pressed As Boolean)
  ' Save the current setting
    SettingsHelper.Save appName, "ParagraphFormat", control.ID, pressed
    
    ' Update the appropriate variable
    Select Case control.ID
        Case "FirstWordCheckBox": FirstWordCheckBox = pressed
        Case "FirstWordDinami"
            FirstWordDinami = pressed
            If FirstWordDinami = True Then FormDinami.Show
        Case "FirstWordLineSpacing": FirstWordLineSpacing = pressed
        Case "FirstWordAlignmentUpwards": FirstWordAlignmentUpwards = pressed
        
        Case "HangingCheckBox": HangingCheckBox = pressed
        Case "HangingOption1": HangingOption1 = pressed
        Case "HangingOption2": HangingOption2 = pressed
        Case "HangingOption3": HangingOption3 = pressed
        Case "HangingSkip2Lines": HangingSkip2Lines = pressed
        
        Case "LastLineCheckBox": LastLineCheckBox = pressed
        Case "LastLineCenterOption": LastLineCenterOption = pressed
        Case "LastLineBalanceOption": LastLineBalanceOption = pressed
'        Case "LastLineOption2": LastLineOption2 = pressed
        Case "LastLineSkip2Lines": LastLineSkip2Lines = pressed
    End Select
    
    ' Group control logic: Uncheck other options in the same group  ' Invalidate controls to update the UI
    If pressed Then
        Select Case control.Tag
            Case "FirstWordGroup"
                If control.ID <> "FirstWordOption1" Then
                    FirstWordOption1 = False
                    myRibbon.InvalidateControl "FirstWordOption1"
                End If
                If control.ID <> "FirstWordOption2" Then
                    FirstWordOption2 = False
                    myRibbon.InvalidateControl "FirstWordOption2"
                End If
                        
            Case "FirstWordLineSpacingGroup"
                If control.ID <> "FirstWordLineSpacing" Then
                    FirstWordLineSpacing = False
                    myRibbon.InvalidateControl "FirstWordLineSpacing"
                End If
            
            Case "HangingGroup"
                If control.ID <> "HangingOption1" Then
                    HangingOption1 = False
                    myRibbon.InvalidateControl "HangingOption1"
                End If
                If control.ID <> "HangingOption2" Then
                    HangingOption2 = False
                    myRibbon.InvalidateControl "HangingOption2"
                End If
                If control.ID <> "HangingOption3" Then
                    HangingOption3 = False
                    myRibbon.InvalidateControl "HangingOption3"
                End If
            
        End Select
    End If
    
End Sub

' Callback for CheckBox getPressed
Sub CheckBox_OnGetPressed(control As IRibbonControl, ByRef returnedVal)

    Select Case control.ID
        Case "FirstWordCheckBox": returnedVal = FirstWordCheckBox
        
        Case "FirstWordDinami": returnedVal = FirstWordDinami
        Case "FirstWordLineSpacing": returnedVal = FirstWordLineSpacing
        Case "FirstWordAlignmentUpwards": returnedVal = FirstWordAlignmentUpwards
        
        Case "HangingCheckBox": returnedVal = HangingCheckBox
        Case "HangingOption1": returnedVal = HangingOption1
        Case "HangingOption2": returnedVal = HangingOption2
        Case "HangingOption3": returnedVal = HangingOption3
        Case "HangingSkip2Lines": returnedVal = HangingSkip2Lines
        
        Case "LastLineCheckBox": returnedVal = LastLineCheckBox
        Case "LastLineCenterOption": returnedVal = LastLineCenterOption
        Case "LastLineBalanceOption": returnedVal = LastLineBalanceOption
'        Case "LastLineOption2": returnedVal = LastLineOption2
        Case "LastLineSkip2Lines": returnedVal = LastLineSkip2Lines
        
        Case Else: returnedVal = False
    End Select
    
    ' Save the current setting
    SettingsHelper.Save appName, "ParagraphFormat", control.ID, returnedVal
End Sub
Function formShow(FormRunning As Object)

    Dim FormTop As Integer
    FormTop = 72
    
    With FormRunning
        .Show vbModeless
        If FirstWordCheckBox Then
            If FirstWordLineSpacing Then
                .LblLineSpacing.Top = FormTop
                FormTop = FormTop + 20
            Else
                .LblLineSpacing.Visible = False
            End If
            If FirstWordOption1 Or FirstWordOption2 Then
                .LblFirstWord.Top = FormTop
                FormTop = FormTop + 20
            Else
                .LblFirstWord.Visible = False
            End If
        Else
            .LblLineSpacing.Visible = False
            .LblFirstWord.Visible = False
        End If
    
        If HangingCheckBox Then
            .LblHangingFirstWord.Top = FormTop
            FormTop = FormTop + 20
        Else
            .LblHangingFirstWord.Visible = False
        End If
            
        If LastLineCheckBox Then
            If LastLineCenterOption Then
                .LblCenterLastLine.Top = FormTop
                FormTop = FormTop + 20
            Else
                .LblCenterLastLine.Visible = False
            End If
            
            If LastLineBalanceOption Then
                .LblBalanceLastLine.Top = FormTop
                FormTop = FormTop + 20
            Else
                .LblBalanceLastLine.Visible = False
            End If
        Else
            .LblCenterLastLine.Visible = False
            .LblBalanceLastLine.Visible = False
        End If
        FormTop = FormTop + 8
        .CBStop.Top = FormTop
        .CBSave.Top = FormTop
        .Height = FormTop + 55
    End With

End Function
Sub FirstWord_OnAction(control As IRibbonControl, pressed As Boolean)
    
    FirstWordOption1 = False
    FirstWordOption2 = False
    Select Case control.ID
        Case "FirstWordOption1": FirstWordOption1 = True
            FontDialogueHelper.ChangeFontDialogSettings
        Case "FirstWordOption2": FirstWordOption2 = True
            FormStyle.Show
    End Select
    
    SettingsHelper.Save appName, "ParagraphFormat", "FirstWordOption1", FirstWordOption1
    SettingsHelper.Save appName, "ParagraphFormat", "FirstWordOption2", FirstWordOption2

    myRibbon.InvalidateControl "FirstWordOption1"
    myRibbon.InvalidateControl "FirstWordOption2"
End Sub
Sub FirstWord_OnGetPressed(control As IRibbonControl, ByRef returnedVal)
    
    Select Case control.ID
        Case "FirstWordOption1": returnedVal = FirstWordOption1
        Case "FirstWordOption2": returnedVal = FirstWordOption2
    End Select
    
End Sub

Sub Hanging_OnAction(control As IRibbonControl, pressed As Boolean)
    
    HangingOption1 = False
    HangingOption2 = False
    HangingOption3 = False
    
    Select Case control.ID
        Case "HangingOption1": HangingOption1 = True
        Case "HangingOption2": HangingOption2 = True
        Case "HangingOption3": HangingOption3 = True
    End Select
    
    SettingsHelper.Save appName, "ParagraphFormat", "HangingOption1", HangingOption1
    SettingsHelper.Save appName, "ParagraphFormat", "HangingOption2", HangingOption2
    SettingsHelper.Save appName, "ParagraphFormat", "HangingOption3", HangingOption3

    myRibbon.InvalidateControl "HangingOption1"
    myRibbon.InvalidateControl "HangingOption2"
    myRibbon.InvalidateControl "HangingOption3"
End Sub
Sub Hanging_OnGetPressed(control As IRibbonControl, ByRef returnedVal)
    
    Select Case control.ID
        Case "HangingOption1": returnedVal = HangingOption1
        Case "HangingOption2": returnedVal = HangingOption2
        Case "HangingOption3": returnedVal = HangingOption3
    End Select
    
End Sub

-------------------------------------------------------------------------------
VBA MACRO FormRunning.frm 
in file: word/vbaProject.bin - OLE stream: 'VBA/FormRunning'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Private Sub CBSave_Click()
    
    Dim doc As Document
    
    On Error Resume Next
    
    For Each doc In Application.Documents
        
        doc.Save
    
    Next doc
    
    On Error GoTo 0
    
    CBSave.Caption = "המסמכים נשמרו בהצלחה"

End Sub

Private Sub CBStop_Click()
    
    stopCode = True
    
End Sub

Private Sub UserForm_Click()

End Sub
-------------------------------------------------------------------------------
VBA MACRO SettingsHelper.bas 
in file: word/vbaProject.bin - OLE stream: 'VBA/SettingsHelper'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit
Option Private Module

Sub Save(appName As String, section As String, key As String, value As Variant)
    Dim folderPath As String
    Dim FilePath As String
        
    If IsMac Then
        folderPath = Environ("HOME") & "/Library/Application Support/" & appName
        FilePath = folderPath & "/settings.ini"
    Else
        folderPath = Environ("USERPROFILE") & "\AppData\Roaming\" & appName
        FilePath = folderPath & "\settings.ini"
    End If
    
    ' Ensure directory exists
    If Dir(folderPath, vbDirectory) = "" Then
        MkDir folderPath
    End If


    Dim lines As Collection
    Set lines = New Collection
    
    Dim fileExists As Boolean
    fileExists = (Dir(FilePath) <> "")
    
    If fileExists Then
        ' Read the file into memory
        Dim fileNum As Integer
        fileNum = FreeFile
        Open FilePath For Input As fileNum
        
        Dim fileLine As String
        Dim currentSection As String
        Dim sectionFound As Boolean
        sectionFound = False
        currentSection = ""
        
        Do Until EOF(fileNum)
            Line Input #fileNum, fileLine
            
            If Left(fileLine, 1) = "[" And Right(fileLine, 1) = "]" Then
                currentSection = Mid(fileLine, 2, Len(fileLine) - 2)
            End If

            ' If we are in the correct section and find the key, update it
            If currentSection = section Then
                If InStr(fileLine, key & "=") = 1 Then
                    fileLine = key & "=" & value & ";" & TypeName(value)
                    sectionFound = True
                End If
            End If
            lines.Add fileLine
        Loop
        Close fileNum
        
        ' If the section was found but the key wasn't, add the new key-value pair
        If Not sectionFound Then
            lines.Add "[" & section & "]"
            lines.Add key & "=" & value & ";" & TypeName(value)
        End If
    Else
        ' If file doesn't exist, start a new one
        lines.Add "[" & section & "]"
        lines.Add key & "=" & value & ";" & TypeName(value)
    End If
    
    ' Write the lines back to the file
    fileNum = FreeFile
    Open FilePath For Output As fileNum
    
    Dim line As Variant
    For Each line In lines
        Print #fileNum, line
    Next line
    
    Close fileNum
End Sub

Function GetSavedSetting(appName As String, section As String, key As String, Optional defaultValue As Variant) As Variant
    Dim folderPath As String
    Dim FilePath As String
        
    If IsMac Then
        folderPath = Environ("HOME") & "/Library/Application Support/" & appName
        FilePath = folderPath & "/settings.ini"
    Else
        folderPath = Environ("USERPROFILE") & "\AppData\Roaming\" & appName
        FilePath = folderPath & "\settings.ini"
    End If
    
    If Dir(FilePath) = "" Then
        GetSavedSetting = defaultValue
        Exit Function
    End If

    Dim fileNum As Integer
    Dim fileLine As String
    Dim currentSection As String
    fileNum = FreeFile

    Open FilePath For Input As fileNum
    currentSection = ""

    Do Until EOF(fileNum)
        Line Input #fileNum, fileLine
        
        If Left(fileLine, 1) = "[" And Right(fileLine, 1) = "]" Then
            currentSection = Mid(fileLine, 2, Len(fileLine) - 2)
        ElseIf currentSection = section Then
            If InStr(fileLine, key & "=") = 1 Then
                Dim valueParts() As String
                valueParts = Split(Mid(fileLine, Len(key) + 2), ";") ' Split value and type

                Dim loadedValue As String
                loadedValue = valueParts(0)
                Dim valueType As String
                valueType = valueParts(1)

                Select Case valueType
                    Case "String"
                        GetSavedSetting = loadedValue
                    Case "Integer"
                        GetSavedSetting = CInt(loadedValue)
                    Case "Long"
                        GetSavedSetting = CLng(loadedValue)
                    Case "Boolean"
                        GetSavedSetting = CBool(loadedValue)
                    Case "Double"
                        GetSavedSetting = CDbl(loadedValue)
                    Case "Single"
                        GetSavedSetting = CSng(loadedValue)
                    Case Else
                        GetSavedSetting = defaultValue
                End Select

                Close fileNum
                Exit Function
            End If
        End If
    Loop
    
    Close fileNum
    GetSavedSetting = defaultValue
End Function

Function IsSameGroup(str1 As String, str2 As String) As Boolean
    ' Check if lengths differ by more than 1, return False immediately
    If Len(str1) <> Len(str2) Or str1 = str2 Then
        IsSameGroup = False
        Exit Function
    End If
    
    ' Compare the substrings excluding the last character
    If Left(str1, Len(str1) - 1) = Left(str2, Len(str2) - 1) Then
        IsSameGroup = True
    Else
        IsSameGroup = False
    End If
End Function

Function IsMac() As Boolean
    On Error Resume Next
    IsMac = (Environ("HOME") <> "")
    On Error GoTo 0
End Function
-------------------------------------------------------------------------------
VBA MACRO FormSaver.frm 
in file: word/vbaProject.bin - OLE stream: 'VBA/FormSaver'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

Private Sub CBNext_Click()
    
    Unload Me

End Sub

Private Sub CBSave_Click()
    
    Dim doc As Document
    
    On Error Resume Next
    
    For Each doc In Application.Documents
        
        doc.Save
    
    Next doc
    
    On Error GoTo 0
    
    Unload Me
    
End Sub

Private Sub CBStop_Click()
    
    stopCode = True
    Unload Me
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
        
    Dim check As Boolean
    
    check = BoxDontShowAgain
    SettingsHelper.Save appName, "DocsSaved", "dontShowAgain", check
    
End Sub
-------------------------------------------------------------------------------
VBA MACRO FormDinami.frm 
in file: word/vbaProject.bin - OLE stream: 'VBA/FormDinami'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit

Private Sub CbOk_Click()
    With Me
        SettingsHelper.Save RibbonControl.appName, "FirsrWordDinami", "SizeOption", .sizeOption.ListIndex
        SettingsHelper.Save RibbonControl.appName, "FirsrWordDinami", "Points", .sizeValueSpin.value
        SettingsHelper.Save RibbonControl.appName, "FirsrWordDinami", "Font", .Font.value
        SettingsHelper.Save RibbonControl.appName, "FirsrWordDinami", "Underline", .Underline.value
        SettingsHelper.Save RibbonControl.appName, "FirsrWordDinami", "Color", .Color.value
        SettingsHelper.Save RibbonControl.appName, "FirsrWordDinami", "StrikeThrough", .strikethrough.value
        SettingsHelper.Save RibbonControl.appName, "FirsrWordDinami", "Bold", .Bold.value
        SettingsHelper.Save RibbonControl.appName, "FirsrWordDinami", "Italic", .Italic.value
        SettingsHelper.Save RibbonControl.appName, "FirsrWordDinami", "Fill", .Fill.value
        SettingsHelper.Save RibbonControl.appName, "FirsrWordDinami", "Effects", .Effects.value
        SettingsHelper.Save RibbonControl.appName, "FirsrWordDinami", "advanced", .advanced.value
    
    End With
    Unload Me
End Sub

Private Sub CommandButton2_Click()
    MsgBox "עצב באופן דינמי - מיועד עבור טקסטים מרובי עיצוב" & vbNewLine & _
           "יש לסמן את ההגדרות שאיתם אנו רוצים לעצב את המילה הראשונה" & vbNewLine & _
           "לדוגמא באם בחרנו 'B'" & vbNewLine & _
           "העיצוב של המילה הראשונה יהיה לפי סגנון/עיצוב מילה ראשונה שבחרנו" & vbNewLine & _
           "כלומר - באם בהגדרות הסגנון מופיע מודגש, המילה תודגש" & vbNewLine & _
           "ובאם מופיע לא מודגש המילה תעוצב כלא מודגש" & vbNewLine & _
           "באם לא נסמן 'B'" & vbNewLine & _
           "המילה הראשונה תקבל את העיצוב של שאר הפיסקה - לא משנה מה מוגדר בסגנון/עיצוב מילה ראשונה שבחרנו"
End Sub

Private Sub sizeOption_Change()
    If sizeOption.value = "גודל משתנה" Then
        points.Locked = False
        sizeValueSpin.Enabled = True
        points.BackColor = &H80000005
    Else
        points.Locked = True
        sizeValueSpin.Enabled = False
        points.BackColor = &H80000000
    End If
End Sub

Private Sub points_Change()
    Dim text As String
    text = points.text
    text = Replace(text, " ", "")
    text = Replace(text, "%", "")
    If IsNumeric(text) Then
        If text > 0 And text < 1000 Then
            sizeValueSpin.value = text
        Else
            points.text = sizeValueSpin.value & "%"
        End If
    Else
        points.text = sizeValueSpin.value & "%"
    End If

End Sub

Private Sub sizeValueSpin_Change()
    points.text = sizeValueSpin.value & "%"
End Sub

Private Sub UserForm_Initialize()

    With sizeOption
        .AddItem "ללא שינוי"
        .AddItem "גודל משתנה"
        .AddItem "גודל קבוע"
    End With
    With Me
        .sizeOption.ListIndex = GetSavedSetting(RibbonControl.appName, "FirsrWordDinami", "SizeOption", 1)
        .sizeValueSpin.value = GetSavedSetting(RibbonControl.appName, "FirsrWordDinami", "Points", 30)
        .Font.value = GetSavedSetting(RibbonControl.appName, "FirsrWordDinami", "Font", False)
        .Color.value = GetSavedSetting(RibbonControl.appName, "FirsrWordDinami", "Color", False)
        .strikethrough.value = GetSavedSetting(RibbonControl.appName, "FirsrWordDinami", "StrikeThrough", False)
        .Bold.value = GetSavedSetting(RibbonControl.appName, "FirsrWordDinami", "Bold", False)
        .Italic.value = GetSavedSetting(RibbonControl.appName, "FirsrWordDinami", "Italic", False)
        .Fill.value = GetSavedSetting(RibbonControl.appName, "FirsrWordDinami", "Fill", False)
        .Effects.value = GetSavedSetting(RibbonControl.appName, "FirsrWordDinami", "Effects", False)
        .advanced.value = GetSavedSetting(RibbonControl.appName, "FirsrWordDinami", "advanced", False)
        .Underline.value = GetSavedSetting(RibbonControl.appName, "FirsrWordDinami", "Underline", False)
    End With
End Sub
-------------------------------------------------------------------------------
VBA MACRO FormStyle.frm 
in file: word/vbaProject.bin - OLE stream: 'VBA/FormStyle'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit

Dim SelectedStyleName As String
Private Sub UserForm_Initialize()
    
    Dim stl As style
    
    ' טעינת סגנונות
    For Each stl In ActiveDocument.Styles
        If stl.Type = wdStyleTypeCharacter Then StyleBox.AddItem stl.NameLocal
    Next stl
    Me.StyleBox.value = GetSavedSetting(RibbonControl.appName, "ParagraphFormat", "FirstWordStyle", "")

End Sub

Private Sub StyleBox_Change()
    
    Dim selectedStyle As style
    Dim StyleFont As Font

    SelectedStyleName = StyleBox.value
    On Error Resume Next
    Set selectedStyle = ActiveDocument.Styles(SelectedStyleName)
    On Error GoTo 0

    If Not selectedStyle Is Nothing Then
        
        Set StyleFont = selectedStyle.Font
                
        ' עיצוב תצוגה מקדימה
        With PreviewText
            With .Font
                .Name = StyleFont.NameBi
                If .Name = "+כותרות עבריות" Then .Name = "Times New Roman"
                .Size = StyleFont.SizeBi
                .Bold = StyleFont.BoldBi
                .Italic = StyleFont.ItalicBi
                .Underline = StyleFont.Underline
            End With
            .Caption = SelectedStyleName
        End With
    End If
    
End Sub


Private Sub CbOk_Click()
    SettingsHelper.Save appName, "ParagraphFormat", "FirstWordStyle", SelectedStyleName
    Unload Me
End Sub
-------------------------------------------------------------------------------
VBA MACRO FormRemoveRunning.frm 
in file: word/vbaProject.bin - OLE stream: 'VBA/FormRemoveRunning'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO FormSelectStyleParagraphFormat.frm 
in file: word/vbaProject.bin - OLE stream: 'VBA/FormSelectStyleParagraphFormat'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit

Private Sub StyleList_Change()
    Dim selectedCount As Long
    Dim i As Long
    
    selectedCount = 0
    For i = 0 To StyleList.ListCount - 1
        If StyleList.Selected(i) Then
            selectedCount = selectedCount + 1
        End If
    Next i
    
    LblSelectedCount.Caption = "נבחרו " & selectedCount & " מתוך 15 סגנונות"

End Sub

Private Sub UserForm_Initialize()
    Dim stl As style
    Dim stlName As Variant
    Dim stlSelectedCollection As New Collection
    Dim i As Integer
    
    For Each stl In ActiveDocument.Styles
        If stl.Type = wdStyleTypeParagraph Or _
            stl.Type = wdStyleTypeLinked Or _
            stl.Type = wdStyleTypeParagraphOnly Then
                StyleList.AddItem stl.NameLocal
        End If
    Next stl
    Set stlSelectedCollection = ReadingValues
    For Each stlName In stlSelectedCollection
        If stlName = "" Then Exit For
        For i = 0 To StyleList.ListCount - 1
            If StyleList.List(i) = stlName Then
                StyleList.Selected(i) = True
            End If
        Next i
    Next stlName
End Sub
Private Sub CB_Cancel_Click()
    Call SaveValues
    stopCode = True
    Unload Me
End Sub
Private Sub CB_OK_Click()
    Call SaveValues
    Unload Me
End Sub
Function SaveValues()
    
    Dim i As Integer
    Dim selectedListNum As Integer
    
    For i = 0 To StyleList.ListCount - 1
        If StyleList.Selected(i) Then
            selectedListNum = selectedListNum + 1
            Call SettingsHelper.Save(appName, "ParagraphFormat", "stlName" & selectedListNum, StyleList.List(i))
            If selectedListNum = 14 Then Exit For
        End If
    Next i
    
    For i = selectedListNum + 1 To 14
        Call SettingsHelper.Save(appName, "ParagraphFormat", "stlName" & i, "")
    Next i

End Function
Function ReadingValues() As Collection
    
    Dim i As Integer
    Dim stlCollection As New Collection
    For i = 1 To 15
        stlCollection.Add SettingsHelper.GetSavedSetting(appName, "ParagraphFormat", "stlName" & i, "")
        If stlCollection(i) = "" Then Exit For
    Next i
    Set ReadingValues = stlCollection
End Function

-------------------------------------------------------------------------------
VBA FORM STRING IN 'word/vbaProject.bin' - OLE stream: 'FormDinami/o'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
0k�Uk�
-------------------------------------------------------------------------------
VBA FORM STRING IN 'word/vbaProject.bin' - OLE stream: 'FormDinami/o'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
0k�ab�
-------------------------------------------------------------------------------
VBA FORM STRING IN 'word/vbaProject.bin' - OLE stream: 'FormDinami/o'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
K�Qlt
-------------------------------------------------------------------------------
VBA FORM STRING IN 'word/vbaProject.bin' - OLE stream: 'FormDinami/o'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
GIF89a
-------------------------------------------------------------------------------
VBA FORM STRING IN 'word/vbaProject.bin' - OLE stream: 'FormDinami/o'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
����ssr::8���FFD������QQO������������hhg�
-------------------------------------------------------------------------------
VBA FORM STRING IN 'word/vbaProject.bin' - OLE stream: 'FormDinami/o'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
4kIH� `
-------------------------------------------------------------------------------
VBA FORM STRING IN 'word/vbaProject.bin' - OLE stream: 'FormDinami/o'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
� $&��
-------------------------------------------------------------------------------
VBA FORM STRING IN 'word/vbaProject.bin' - OLE stream: 'FormDinami/o'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
���2a0f
-------------------------------------------------------------------------------
VBA FORM STRING IN 'word/vbaProject.bin' - OLE stream: 'FormDinami/o'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
M��4*
-------------------------------------------------------------------------------
VBA FORM STRING IN 'word/vbaProject.bin' - OLE stream: 'FormDinami/o'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
5��|N������
-------------------------------------------------------------------------------
VBA FORM STRING IN 'word/vbaProject.bin' - OLE stream: 'FormDinami/o'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
0b��
-------------------------------------------------------------------------------
VBA FORM STRING IN 'word/vbaProject.bin' - OLE stream: 'FormDinami/o'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
0b��
-------------------------------------------------------------------------------
VBA FORM STRING IN 'word/vbaProject.bin' - OLE stream: 'FormDinami/o'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
0b��
-------------------------------------------------------------------------------
VBA FORM STRING IN 'word/vbaProject.bin' - OLE stream: 'FormDinami/o'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
0b��
-------------------------------------------------------------------------------
VBA FORM STRING IN 'word/vbaProject.bin' - OLE stream: 'FormDinami/o'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
����a
-------------------------------------------------------------------------------
VBA FORM STRING IN 'word/vbaProject.bin' - OLE stream: 'FormRemoveRunning/o'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Tahoma�
-------------------------------------------------------------------------------
VBA FORM STRING IN 'word/vbaProject.bin' - OLE stream: 'FormRemoveRunning/o'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Tahomae
-------------------------------------------------------------------------------
VBA FORM STRING IN 'word/vbaProject.bin' - OLE stream: 'FormRemoveRunning/o'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Tahoma�
-------------------------------------------------------------------------------
VBA FORM STRING IN 'word/vbaProject.bin' - OLE stream: 'FormRemoveRunning/o'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Tahoma�
-------------------------------------------------------------------------------
VBA FORM STRING IN 'word/vbaProject.bin' - OLE stream: 'FormRemoveRunning/o'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Tahomae
-------------------------------------------------------------------------------
VBA FORM STRING IN 'word/vbaProject.bin' - OLE stream: 'FormRemoveRunning/o'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Tahomae
-------------------------------------------------------------------------------
VBA FORM STRING IN 'word/vbaProject.bin' - OLE stream: 'FormRunning/o'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Tahomaox
-------------------------------------------------------------------------------
VBA FORM STRING IN 'word/vbaProject.bin' - OLE stream: 'FormRunning/o'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Tahoma�
-------------------------------------------------------------------------------
VBA FORM STRING IN 'word/vbaProject.bin' - OLE stream: 'FormRunning/o'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Tahoma�
-------------------------------------------------------------------------------
VBA FORM STRING IN 'word/vbaProject.bin' - OLE stream: 'FormSaver/o'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Tahomaox
-------------------------------------------------------------------------------
VBA FORM STRING IN 'word/vbaProject.bin' - OLE stream: 'FormSaver/o'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Tahomaox
-------------------------------------------------------------------------------
VBA FORM STRING IN 'word/vbaProject.bin' - OLE stream: 'FormSaver/o'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Tahomaox
-------------------------------------------------------------------------------
VBA FORM STRING IN 'word/vbaProject.bin' - OLE stream: 'FormSaver/o'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(�,���
-------------------------------------------------------------------------------
VBA FORM STRING IN 'word/vbaProject.bin' - OLE stream: 'FormSaver/o'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
0aho�
-------------------------------------------------------------------------------
VBA FORM STRING IN 'word/vbaProject.bin' - OLE stream: 'FormSelectStyleParagraphFormat/o'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Tahomaox
-------------------------------------------------------------------------------
VBA FORM STRING IN 'word/vbaProject.bin' - OLE stream: 'FormStyle/o'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Tahomae
-------------------------------------------------------------------------------
VBA FORM STRING IN 'word/vbaProject.bin' - OLE stream: 'FormStyle/o'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Tahomae
-------------------------------------------------------------------------------
VBA FORM Variable "b'Label1'" IN 'word/vbaProject.bin' - OLE stream: 'FormDinami'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
None
-------------------------------------------------------------------------------
VBA FORM Variable "b'Bold'" IN 'word/vbaProject.bin' - OLE stream: 'FormDinami'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
b'0'
-------------------------------------------------------------------------------
VBA FORM Variable "b'Italic'" IN 'word/vbaProject.bin' - OLE stream: 'FormDinami'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
b'0'
-------------------------------------------------------------------------------
VBA FORM Variable "b'Underline'" IN 'word/vbaProject.bin' - OLE stream: 'FormDinami'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
b'0'
-------------------------------------------------------------------------------
VBA FORM Variable "b'strikethrough'" IN 'word/vbaProject.bin' - OLE stream: 'FormDinami'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
b'0'
-------------------------------------------------------------------------------
VBA FORM Variable "b'Color'" IN 'word/vbaProject.bin' - OLE stream: 'FormDinami'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
b'0'
-------------------------------------------------------------------------------
VBA FORM Variable "b'points'" IN 'word/vbaProject.bin' - OLE stream: 'FormDinami'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
b''
-------------------------------------------------------------------------------
VBA FORM Variable "b'Font'" IN 'word/vbaProject.bin' - OLE stream: 'FormDinami'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
b'0'
-------------------------------------------------------------------------------
VBA FORM Variable "b'CBOk'" IN 'word/vbaProject.bin' - OLE stream: 'FormDinami'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
None
-------------------------------------------------------------------------------
VBA FORM Variable "b'sizeOption'" IN 'word/vbaProject.bin' - OLE stream: 'FormDinami'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
b''
-------------------------------------------------------------------------------
VBA FORM Variable "b'advanced'" IN 'word/vbaProject.bin' - OLE stream: 'FormDinami'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
b'0'
-------------------------------------------------------------------------------
VBA FORM Variable "b'Fill'" IN 'word/vbaProject.bin' - OLE stream: 'FormDinami'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
b'0'
-------------------------------------------------------------------------------
VBA FORM Variable "b'Effects'" IN 'word/vbaProject.bin' - OLE stream: 'FormDinami'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
b'0'
-------------------------------------------------------------------------------
VBA FORM Variable "b'CommandButton2'" IN 'word/vbaProject.bin' - OLE stream: 'FormDinami'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
None
-------------------------------------------------------------------------------
VBA FORM Variable "b'sizeValueSpin'" IN 'word/vbaProject.bin' - OLE stream: 'FormDinami'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
None
-------------------------------------------------------------------------------
VBA FORM Variable "b'Label6'" IN 'word/vbaProject.bin' - OLE stream: 'FormRemoveRunning'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
None
-------------------------------------------------------------------------------
VBA FORM Variable "b'LblLineSpacing'" IN 'word/vbaProject.bin' - OLE stream: 'FormRemoveRunning'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
None
-------------------------------------------------------------------------------
VBA FORM Variable "b'LblFirstWord'" IN 'word/vbaProject.bin' - OLE stream: 'FormRemoveRunning'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
None
-------------------------------------------------------------------------------
VBA FORM Variable "b'LblHangingFirstWord'" IN 'word/vbaProject.bin' - OLE stream: 'FormRemoveRunning'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
None
-------------------------------------------------------------------------------
VBA FORM Variable "b'LblCenterLastLine'" IN 'word/vbaProject.bin' - OLE stream: 'FormRemoveRunning'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
None
-------------------------------------------------------------------------------
VBA FORM Variable "b'LblBalanceLastLine'" IN 'word/vbaProject.bin' - OLE stream: 'FormRemoveRunning'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
None
-------------------------------------------------------------------------------
VBA FORM Variable "b'CBSave'" IN 'word/vbaProject.bin' - OLE stream: 'FormRemoveRunning'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
None
-------------------------------------------------------------------------------
VBA FORM Variable "b'CBStop'" IN 'word/vbaProject.bin' - OLE stream: 'FormRemoveRunning'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
None
-------------------------------------------------------------------------------
VBA FORM Variable "b'Label7'" IN 'word/vbaProject.bin' - OLE stream: 'FormRemoveRunning'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
None
-------------------------------------------------------------------------------
VBA FORM Variable "b'LblLineSpacing'" IN 'word/vbaProject.bin' - OLE stream: 'FormRunning'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
None
-------------------------------------------------------------------------------
VBA FORM Variable "b'LblFirstWord'" IN 'word/vbaProject.bin' - OLE stream: 'FormRunning'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
None
-------------------------------------------------------------------------------
VBA FORM Variable "b'LblHangingFirstWord'" IN 'word/vbaProject.bin' - OLE stream: 'FormRunning'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
None
-------------------------------------------------------------------------------
VBA FORM Variable "b'LblCenterLastLine'" IN 'word/vbaProject.bin' - OLE stream: 'FormRunning'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
None
-------------------------------------------------------------------------------
VBA FORM Variable "b'LblBalanceLastLine'" IN 'word/vbaProject.bin' - OLE stream: 'FormRunning'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
None
-------------------------------------------------------------------------------
VBA FORM Variable "b'Label6'" IN 'word/vbaProject.bin' - OLE stream: 'FormRunning'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
None
-------------------------------------------------------------------------------
VBA FORM Variable "b'CBSave'" IN 'word/vbaProject.bin' - OLE stream: 'FormRunning'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
None
-------------------------------------------------------------------------------
VBA FORM Variable "b'CBStop'" IN 'word/vbaProject.bin' - OLE stream: 'FormRunning'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
None
-------------------------------------------------------------------------------
VBA FORM Variable "b'Label7'" IN 'word/vbaProject.bin' - OLE stream: 'FormRunning'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
None
-------------------------------------------------------------------------------
VBA FORM Variable "b'lblMessage'" IN 'word/vbaProject.bin' - OLE stream: 'FormSaver'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
None
-------------------------------------------------------------------------------
VBA FORM Variable "b'CBStop'" IN 'word/vbaProject.bin' - OLE stream: 'FormSaver'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
None
-------------------------------------------------------------------------------
VBA FORM Variable "b'CBSave'" IN 'word/vbaProject.bin' - OLE stream: 'FormSaver'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
None
-------------------------------------------------------------------------------
VBA FORM Variable "b'CBNext'" IN 'word/vbaProject.bin' - OLE stream: 'FormSaver'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
None
-------------------------------------------------------------------------------
VBA FORM Variable "b'BoxDontShowAgain'" IN 'word/vbaProject.bin' - OLE stream: 'FormSaver'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
b'0'
-------------------------------------------------------------------------------
VBA FORM Variable "b'CB_Cancel'" IN 'word/vbaProject.bin' - OLE stream: 'FormSelectStyleParagraphFormat'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
None
-------------------------------------------------------------------------------
VBA FORM Variable "b'CB_OK'" IN 'word/vbaProject.bin' - OLE stream: 'FormSelectStyleParagraphFormat'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
None
-------------------------------------------------------------------------------
VBA FORM Variable "b'Label2'" IN 'word/vbaProject.bin' - OLE stream: 'FormSelectStyleParagraphFormat'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
None
-------------------------------------------------------------------------------
VBA FORM Variable "b'LblSelectedCount'" IN 'word/vbaProject.bin' - OLE stream: 'FormSelectStyleParagraphFormat'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
None
-------------------------------------------------------------------------------
VBA FORM Variable "b'StyleList'" IN 'word/vbaProject.bin' - OLE stream: 'FormSelectStyleParagraphFormat'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
b''
-------------------------------------------------------------------------------
VBA FORM Variable "b'StyleBox'" IN 'word/vbaProject.bin' - OLE stream: 'FormStyle'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
b''
-------------------------------------------------------------------------------
VBA FORM Variable "b'CbOk'" IN 'word/vbaProject.bin' - OLE stream: 'FormStyle'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
None
-------------------------------------------------------------------------------
VBA FORM Variable "b'PreviewText'" IN 'word/vbaProject.bin' - OLE stream: 'FormStyle'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
None

