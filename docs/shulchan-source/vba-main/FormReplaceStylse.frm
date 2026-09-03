VBA MACRO FormReplaceStylse.frm 
in file: word/vbaProject.bin - OLE stream: 'VBA/FormReplaceStylse'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit

Private Sub CB_Cancel_Click()
    Unload Me
End Sub

Private Sub CB_OK_Click()
    Dim splitName() As String
    If Findlist.ListIndex = -1 Or ReplaceList.ListIndex = -1 Then
        MsgBox "יש לבחור סגנונות לביצוע הפעולה", vbOKOnly, "שגיאה"
        Exit Sub
    ElseIf Findlist.value = ReplaceList.value Then
        MsgBox "יש לבחור סגנונות שונים לביצוע הפעולה", vbOKOnly, "שגיאה"
        Exit Sub
    End If
    splitName = Split(ReplaceList.value, " - ")
    Call EditingErrors.FindAndReplaceStylse(Split(Findlist.value, " - ")(0), splitName(0))
    
    Select Case splitName(1)
        Case "תו ופיסקה"
            Call Helpers.RemoveItemFromCollection(EditingErrors.linkedStyles, splitName(0))
        Case "תו"
            Call Helpers.RemoveItemFromCollection(EditingErrors.charStyles, splitName(0))
        Case "פיסקה"
            Call Helpers.RemoveItemFromCollection(EditingErrors.paraStyles, splitName(0))
    End Select
    Call UserForm_Initialize
End Sub

Private Sub Findlist_Click()
    Dim stlName As Variant
    
    ReplaceList.Clear
    
    If Findlist.value Like "*" & "- תו ופיסקה" Then
        For Each stlName In EditingErrors.linkedStyles
            If stlName <> "" Then ReplaceList.AddItem stlName & " - תו ופיסקה"
        Next stlName
    ElseIf Findlist.value Like "*" & "- פיסקה" Then
        For Each stlName In EditingErrors.paraStyles
            If stlName <> "" Then ReplaceList.AddItem stlName & " - פיסקה"
        Next stlName
        For Each stlName In EditingErrors.linkedStyles
            If stlName <> "" Then ReplaceList.AddItem stlName & " - תו ופיסקה"
        Next stlName
    ElseIf Findlist.value Like "*" & "- תו" Then
        For Each stlName In EditingErrors.charStyles
            If stlName <> "" Then ReplaceList.AddItem stlName & " - תו"
        Next stlName
        For Each stlName In EditingErrors.linkedStyles
            If stlName <> "" Then ReplaceList.AddItem stlName & " - תו ופיסקה"
        Next stlName
    End If
End Sub

Private Sub UserForm_Initialize()
    Dim stlName As Variant
    Findlist.Clear
    ReplaceList.Clear
    For Each stlName In EditingErrors.paraStyles
        If stlName <> "" Then Findlist.AddItem CStr(stlName) & " - פיסקה"
    Next stlName
    For Each stlName In EditingErrors.charStyles
        If stlName <> "" Then Findlist.AddItem CStr(stlName) & " - תו"
    Next stlName
    For Each stlName In EditingErrors.linkedStyles
        If stlName <> "" Then Findlist.AddItem CStr(stlName) & " - תו ופיסקה"
    Next stlName
End Sub
-------------------------------------------------------------------------------
