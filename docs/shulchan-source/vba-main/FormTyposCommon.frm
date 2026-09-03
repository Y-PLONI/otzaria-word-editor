VBA MACRO FormTyposCommon.frm 
in file: word/vbaProject.bin - OLE stream: 'VBA/FormTyposCommon'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit

Private Sub CB_Cancel_Click()
    Call SaveValues
    Unload Me
End Sub

Private Sub CB_OK_Click()
    Dim selectedOptions(0 To 8) As Boolean
    Dim i As Integer
    
    For i = 0 To OptionsList.ListCount - 1
        If OptionsList.Selected(i) Then
            selectedOptions(i) = True
        End If
    Next i
    Call SaveValues
    Unload Me
    Call Typos.Repair(selectedOptions)
    
End Sub

Private Sub UserForm_Initialize()
    
    Dim i As Integer
    
    OptionsList.AddItem "מחיקת רווחים מיותרים"
    OptionsList.AddItem "פיסקאות ריקות"
    OptionsList.AddItem "רוח לפני תוי פיסוק"
    OptionsList.AddItem "סימני פיסוק כפולים"
    OptionsList.AddItem "מעל 3 נקודות"
    OptionsList.AddItem "רווח לפני ואחרי סוגריים"
    OptionsList.AddItem "רווח לפני ואחרי פיסקה"
    OptionsList.AddItem "זוג גרשיים בודדים לגרשיים אחד"
    OptionsList.AddItem "אות אנגלית אחרי מרכאות"
    
    For i = 0 To OptionsList.ListCount - 1
        OptionsList.Selected(i) = SettingsHelper.GetSavedSetting(appName, "TyposCommon", "selectedOptions" & i, False)
    Next i

End Sub
Sub SaveValues()
    
    Dim i As Integer
    
    For i = 0 To OptionsList.ListCount - 1
        Call SettingsHelper.Save(appName, "TyposCommon", "selectedOptions" & i, OptionsList.Selected(i))
    Next i
    
End Sub
-------------------------------------------------------------------------------
