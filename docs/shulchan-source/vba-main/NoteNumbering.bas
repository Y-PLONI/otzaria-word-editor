VBA MACRO NoteNumbering.bas 
in file: word/vbaProject.bin - OLE stream: 'VBA/NoteNumbering'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Sub SplitDocumentToMainAndNotes()
    Dim doc As Document
    Dim ftNoteDoc As Document
    Dim edNoteDoc As Document
    Dim ftNote As Footnote
    Dim enNote As Endnote
    Dim rng As Range
    Dim bbRng As Range
    Dim bb As BuildingBlock
    Dim i As Integer
    
    FormRunning.Show vbModeless
    Application.UndoRecord.StartCustomRecord "מספור עברי להערות שולים"
    Set doc = ActiveDocument
    
'    Select Case numbering(titleStyle(i))
'        Case 3
            Set bb = ThisDocument.AttachedTemplate.BuildingBlockEntries("HebrewNumbers")
'        Case 4
'            Set bb = ThisDocument.AttachedTemplate.BuildingBlockEntries("HebrewNumbersClean")
'        Case 5
'            Set bb = ThisDocument.AttachedTemplate.BuildingBlockEntries("HebrewNumbersShortK")
'    End Select
    
    ' בודק שיש הערות שולים במסמך
    If doc.Footnotes.Count = 0 Then GoTo SecondAction
    
    ' עובר על כל הערות השולים ומקבע את המספר בגוף ההערה
    For i = 1 To doc.Footnotes.Count
        FormRunning.Label1.Caption = "ממיר מספור הערות שולים " & i & " מתוך " & doc.Footnotes.Count: DoEvents
        
        Set ftNote = doc.Footnotes(i)
        Set rng = ftNote.Range
        With rng
            .Expand wdParagraph
            .Collapse wdCollapseStart
            .MoveEndUntil " " & Chr(13)
            .Delete
            Set bbRng = bb.Insert(where:=rng)
            Call ConvertToNoteRef(bbRng, "FOOTNOTEREF " & ftNote.Index)
            .End = bbRng.End
            .Font.Superscript = True
            .Fields.Update
        End With
        
        Set rng = ftNote.Reference
        With rng
            .Collapse wdCollapseEnd
            Set bbRng = bb.Insert(where:=rng)
            Call ConvertToNoteRef(bbRng, "FOOTNOTEREF " & ftNote.Index)
            .End = bbRng.End
            .Font.Superscript = True
            .Fields.Update
            ftNote.Reference.Font.Hidden = False
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
Function ConvertToNoteRef(rng As Range, noteIndxs As String)
    Dim fld As Field
    For Each fld In rng.Fields
        With fld.Code.Find
            .text = "page "
            .Replacement.text = noteIndxs
            .MatchWholeWord = True
            .Execute Replace:=wdReplaceAll
        End With
'        fld.Update
    Next fld
End Function

-------------------------------------------------------------------------------
