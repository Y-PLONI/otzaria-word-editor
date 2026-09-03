VBA MACRO UserForm1.frm 
in file: word/vbaProject.bin - OLE stream: 'VBA/UserForm1'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Private Sub UserForm_Initialize()
    Dim stl As Style
    
    For Each stl In ActiveDocument.Styles
        If stl.Type = wdStyleTypeParagraph Or _
            stl.Type = wdStyleTypeLinked Or _
            stl.Type = wdStyleTypeParagraphOnly Then
                ListBox1.AddItem stl.NameLocal
        End If
    Next stl
End Sub
-------------------------------------------------------------------------------
