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

