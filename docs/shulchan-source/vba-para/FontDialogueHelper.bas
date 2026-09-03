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
