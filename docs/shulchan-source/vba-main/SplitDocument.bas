VBA MACRO SplitDocument.bas 
in file: word/vbaProject.bin - OLE stream: 'VBA/SplitDocument'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Sub SplitDocumentToMainAndNotes(footnotesTest As Boolean, endnotesTest As Boolean)
    Dim doc As Document
    Dim ftNoteDoc As Document
    Dim edNoteDoc As Document
    Dim ftNote As Footnote
    Dim enNote As Endnote
    Dim rng As Range
    Dim i As Integer
    
    FormRunning.Show vbModeless
    Application.UndoRecord.StartCustomRecord "פירוק מסמך"
    Set doc = ActiveDocument
    
    ' בודק שיש הערות שולים במסמך
    If Not footnotesTest Or doc.Footnotes.Count = 0 Then GoTo SecondAction
    
    ' עובר על כל הערות השולים ומקבע את המספר בגוף ההערה
    For i = 1 To doc.Footnotes.Count
        FormRunning.Label1.Caption = "מקבע מספר הערת שולים " & i & " מתוך " & doc.Footnotes.Count: DoEvents
        
        Set ftNote = doc.Footnotes(i)
        Set rng = ftNote.Range
        With rng
            .Expand wdParagraph
            .Collapse wdCollapseStart
            .MoveEndUntil " " & Chr(13)
            .Delete
            .InsertCrossReference _
                ReferenceType:=wdRefTypeFootnote, _
                ReferenceKind:=wdFootnoteNumber, _
                ReferenceItem:=ftNote.Index, _
                InsertAsHyperlink:=False, _
                IncludePosition:=False, _
                SeparateNumbers:=False
            .MoveEnd
            .Fields(1).Update
            .Fields(1).Result.Font.Superscript = True
            .Fields(1).Result.InsertAfter " "
            .Fields(1).Unlink
        End With
    Next i
    Application.UndoRecord.EndCustomRecord
    
    ' מעתיק את הערות השולים למסמך נפרד
    FormRunning.Label1.Caption = "יוצר מסמך חדש עם הערות השולים": DoEvents
    doc.Footnotes(1).Range.Select
    Selection.WholeStory
    Selection.Copy
    DoEvents
    Set ftNoteDoc = Application.Documents.Add
    doc.Activate
    ftNoteDoc.Range.Paste
    
    Application.UndoRecord.StartCustomRecord "פירוק מסמך"

    ' מוחק את כל הערות השולים ומקבע את המספר בטקסט הראשי
    For i = doc.Footnotes.Count To 1 Step -1
        FormRunning.Label1.Caption = "מוחק הערות שולים " & doc.Footnotes.Count & " נותרו": DoEvents
        
        Set ftNote = doc.Footnotes(i)
        Set rng = ftNote.Reference
        With rng
            .Collapse wdCollapseEnd
            .InsertCrossReference _
                ReferenceType:=wdRefTypeFootnote, _
                ReferenceKind:=wdFootnoteNumber, _
                ReferenceItem:=ftNote.Index, _
                InsertAsHyperlink:=False, _
                IncludePosition:=False, _
                SeparateNumbers:=False
            .MoveEnd
            .Fields(1).Update
            .Fields(1).Result.Font.Superscript = True
            .Fields(1).Unlink
            ftNote.Delete
        End With
    Next i

SecondAction:
    ' בודק שיש הערות סיום במסמך
    If Not endnotesTest Or doc.Endnotes.Count = 0 Then GoTo Ending
    
    ' עובר על כל הערות הסיום ומקבע את המספר בגוף ההערה
    For i = 1 To doc.Endnotes.Count
        FormRunning.Label1.Caption = "מקבע מספר הערת סיום " & i & " מתוך " & doc.Endnotes.Count: DoEvents
        
        Set enNote = doc.Endnotes(i)
        Set rng = enNote.Range
        With rng
            .Expand wdParagraph
            .Collapse wdCollapseStart
            .MoveEndUntil " " & Chr(13)
            .Delete
            .InsertCrossReference _
                ReferenceType:=wdRefTypeEndnote, _
                ReferenceKind:=wdEndnoteNumber, _
                ReferenceItem:=enNote.Index, _
                InsertAsHyperlink:=False, _
                IncludePosition:=False, _
                SeparateNumbers:=False
            .MoveEnd
            .Fields(1).Update
            .Fields(1).Result.Font.Superscript = True
            .Fields(1).Result.InsertAfter " "
            .Fields(1).Unlink
        End With
    Next i
    Application.UndoRecord.EndCustomRecord
    
    ' מעתיק את הערות הסיום למסמך נפרד
    FormRunning.Label1.Caption = "יוצר מסמך חדש עם הערות הסיום": DoEvents
    doc.Endnotes(1).Range.Select
    Selection.WholeStory
    Selection.Copy
    DoEvents
    Set enNoteDoc = Application.Documents.Add
    doc.Activate
    enNoteDoc.Range.Paste
    
    Application.UndoRecord.StartCustomRecord "פירוק מסמך"
    
    ' מוחק את כל הערות הסיום ומקבע את המספר בטקסט הראשי
    For i = doc.Endnotes.Count To 1 Step -1
        FormRunning.Label1.Caption = "מוחק הערות סיום " & doc.Endnotes.Count & " נותרו": DoEvents
        
        Set enNote = doc.Endnotes(i)
        Set rng = enNote.Reference
        With rng
            .Collapse wdCollapseEnd
            .InsertCrossReference _
                ReferenceType:=wdRefTypeEndnote, _
                ReferenceKind:=wdEndnoteNumber, _
                ReferenceItem:=enNote.Index, _
                InsertAsHyperlink:=False, _
                IncludePosition:=False, _
                SeparateNumbers:=False
            .MoveEnd
            .Fields(1).Update
            .Fields(1).Result.Font.Superscript = True
            .Fields(1).Unlink
            enNote.Delete
        End With
    Next i
Ending:
    Application.UndoRecord.EndCustomRecord
    Unload FormRunning
End Sub
-------------------------------------------------------------------------------
