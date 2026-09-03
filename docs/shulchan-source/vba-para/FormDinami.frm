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
