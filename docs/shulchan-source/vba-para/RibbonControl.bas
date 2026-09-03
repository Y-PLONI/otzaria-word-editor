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
