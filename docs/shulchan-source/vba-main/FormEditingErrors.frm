VBA MACRO FormEditingErrors.frm 
in file: word/vbaProject.bin - OLE stream: 'VBA/FormEditingErrors'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

Private Sub UserForm_Initialize()

    Dim ListBox1Height As Double
    
    EditingErrors.userTest = False
    
    With ListBox1
        For i = LBound(itemsToList) To UBound(itemsToList)
            .AddItem EditingErrors.itemsToList(i)
        Next i
        ListBox1Height = .ListCount * 12 + 1
        .Height = ListBox1Height
    
        CB_Next.Top = ListBox1Height + 10 + .Top
        CB_Cancel.Top = ListBox1Height + 10 + .Top
        Me.Height = CB_Next.Top + CB_Next.Height + 35
    End With

End Sub
Private Sub CB_Next_Click()
    EditingErrors.userTest = True
    EditingErrors.userSelection = ListBox1.ListIndex
    Unload Me
End Sub
Private Sub CB_Cancel_Click()
    Unload Me
End Sub


-------------------------------------------------------------------------------
