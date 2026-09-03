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
