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

-------------------------------------------------------------------------------
