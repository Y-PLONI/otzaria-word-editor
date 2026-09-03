VBA MACRO FormTextAlternating.frm 
in file: word/vbaProject.bin - OLE stream: 'VBA/FormTextAlternating'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit
Private Sub CB_Cancel_Click()
    Call SaveValues
    Unload Me
End Sub

Private Sub CB_OK_Click()
    Call SaveValues
    Unload Me
    Call TextAlternating.Starting(EndCharBox, StartCharBox)
End Sub

Private Sub UserForm_Initialize()
    StartCharBox.value = SettingsHelper.GetSavedSetting(appName, "TextAlternating", "startChar", ":")
    EndCharBox.value = SettingsHelper.GetSavedSetting(appName, "TextAlternating", "endChar", ".")
End Sub
Sub SaveValues()
    Call SettingsHelper.Save(appName, "TextAlternating", "endChar", EndCharBox.value)
    Call SettingsHelper.Save(appName, "TextAlternating", "startChar", StartCharBox.value)
End Sub
-------------------------------------------------------------------------------
