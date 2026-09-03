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
