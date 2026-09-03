VBA MACRO FormDocReduction.frm 
in file: word/vbaProject.bin - OLE stream: 'VBA/FormDocReduction'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit

Private Sub UserForm_Initialize()
    
    If DocReduction.pageCountTarget > 0 Then PageCountSpin.value = pageCountTarget
    If DocReduction.marginsTarget > 0 Then MarginsSpin.value = marginsTarget
    If DocReduction.paraSpaceTarget > 0 Then ParaSpaceSpin.value = paraSpaceTarget
    If DocReduction.linesSpaceTarget > 0 Then LinesSpaceSpin.value = linesSpaceTarget
    If DocReduction.fontSizeTarget > 0 Then FontSizeSpin.value = fontSizeTarget
    If DocReduction.fontLimitTarget > 0 Then FontLimitSpin.value = fontLimitTarget
    DocReduction.pageCountTarget = 0

End Sub
Private Sub CbCancel_Click()
    Unload Me
End Sub

Private Sub CbOk_Click()
    
    pageCountTarget = PageCountSpin.value
    
    DocReduction.marginsTarget = MarginsSpin.value
    If Not MarginsCheck Then DocReduction.marginsTarget = 0
    
    DocReduction.paraSpaceTarget = ParaSpaceSpin.value
    If Not ParaSpaceCheck Then DocReduction.paraSpaceTarget = 0
    
    DocReduction.linesSpaceTarget = LinesSpaceSpin.value
    If Not LinesSpaceCheck Then DocReduction.linesSpaceTarget = 0
    
    DocReduction.fontSizeTarget = FontSizeSpin.value
    If Not FontSizeCheck Then DocReduction.fontSizeTarget = 0
    
    DocReduction.fontLimitTarget = FontLimitSpin.value
    If Not FontLimitCheck Then DocReduction.fontLimitTarget = 5
    
    Unload Me
End Sub

Private Sub PageCountSpin_Change()
    PageCountBox.text = PageCountSpin.value
End Sub
Private Sub MarginsSpin_Change()
    MarginsBox.text = MarginsSpin.value
End Sub
Private Sub ParaSpaceSpin_Change()
    ParaSpaceBox.text = ParaSpaceSpin.value
End Sub
Private Sub LinesSpaceSpin_Change()
    LinesSpaceBox.text = LinesSpaceSpin.value
End Sub
Private Sub FontSizeSpin_Change()
    FontSizeBox.text = FontSizeSpin.value
End Sub
Private Sub FontLimitSpin_Change()
    FontLimitBox.text = FontLimitSpin.value
End Sub

Private Sub PageCountBox_Change()
    Dim text As String
    text = PageCountBox.value
    If IsNumeric(text) Then
        If text > 0 Then
            PageCountSpin.value = text
        Else
            PageCountBox.text = PageCountSpin.value
        End If
    Else
        PageCountBox.text = PageCountSpin.value
    End If
End Sub
Private Sub MarginsBox_Change()
    Dim text As String
    text = MarginsBox.value
    If IsNumeric(text) Then
        If text > 0 And text < 10000 Then
            MarginsSpin.value = text
        Else
            MarginsBox.text = MarginsSpin.value
        End If
    Else
        MarginsBox.text = MarginsSpin.value
    End If
End Sub
Private Sub ParaSpaceBox_Change()
    Dim text As String
    text = ParaSpaceBox.value
    If IsNumeric(text) Then
        If text > 0 And text < 11 Then
            ParaSpaceSpin.value = text
        Else
            ParaSpaceBox.text = ParaSpaceSpin.value
        End If
    Else
        ParaSpaceBox.text = ParaSpaceSpin.value
    End If
End Sub
Private Sub LinesSpaceBox_Change()
    Dim text As String
    text = LinesSpaceBox.value
    If IsNumeric(text) Then
        If text > 0 And text < 11 Then
            LinesSpaceSpin.value = text
        Else
            LinesSpaceBox.text = LinesSpaceSpin.value
        End If
    Else
        LinesSpaceBox.text = LinesSpaceSpin.value
    End If
End Sub
Private Sub FontSizeBox_Change()
    Dim text As String
    text = FontSizeBox.value
    If IsNumeric(text) Then
        If text > 0 And text < 11 Then
            FontSizeSpin.value = text
        Else
            FontSizeBox.text = FontSizeSpin.value
        End If
    Else
        FontSizeBox.text = FontSizeSpin.value
    End If
End Sub
Private Sub FontLimitBox_Change()
    Dim text As String
    text = FontLimitBox.value
    If IsNumeric(text) Then
        If text > 4 And text < 100 Then
            FontLimitSpin.value = text
        Else
            FontLimitBox.text = FontLimitSpin.value
        End If
    Else
        FontLimitBox.text = FontLimitSpin.value
    End If
End Sub


-------------------------------------------------------------------------------
