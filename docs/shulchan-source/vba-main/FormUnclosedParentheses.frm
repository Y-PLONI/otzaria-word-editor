VBA MACRO FormUnclosedParentheses.frm 
in file: word/vbaProject.bin - OLE stream: 'VBA/FormUnclosedParentheses'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Dim entireDocument As Boolean
Private Sub UserForm_Initialize()
    entireDocument = SettingsHelper.GetSavedSetting(appName, "UnclosedParentheses", "entireDocument", False)
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Call UnclosedParentheses.ResetRange
End Sub
Private Sub CBStart_Click()
    Call UnclosedParentheses.ResetRange
    Call UnclosedParentheses.Search(entireDocument)
End Sub

Private Sub CBNext_Click()
    If CBNext.Caption = "סיים" Then Unload Me: Exit Sub
    Call UnclosedParentheses.Search(entireDocument)
End Sub
-------------------------------------------------------------------------------
