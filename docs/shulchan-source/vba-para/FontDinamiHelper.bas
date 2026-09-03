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
