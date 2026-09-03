VBA MACRO SH_RibbonControl.bas 
in file: word/vbaProject.bin - OLE stream: 'VBA/SH_RibbonControl'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit
Option Private Module

Public myRibbon As IRibbonUI
Public appName As String
Public stopCode As Boolean

Sub OnLoad(ribbon As IRibbonUI)
    
    Set myRibbon = ribbon
    appName = "ShulchanHaorech"
    Call BracketsAndFootnotes.OnLoad

End Sub

Sub AlignPages_OnAction(control As IRibbonControl)
    
    Dim SavedRange As Range
    
    stopCode = False
    Call CheckDocsSaved
    If stopCode Then Exit Sub
    
    Set SavedRange = Selection.Range
    
    Select Case control.ID
        Case "AlignPagesToEnd"
            Selection.EndKey unit:=wdStory, Extend:=wdExtend
        Case "AlignPagesAll"
            Selection.HomeKey unit:=wdStory
            Selection.EndKey unit:=wdStory, Extend:=wdExtend
    End Select
    
    Call AlignPages.Repair
    
    SavedRange.Select

End Sub
Sub AlignColumns_OnAction(control As IRibbonControl)
    
    Dim SavedRange As Range
    Dim rng As Range
    
    stopCode = False
    Call CheckDocsSaved
    If stopCode Then Exit Sub
        
    Set SavedRange = Selection.Range
    
    Select Case control.ID
        Case "AlignColumnsButton"
            Set rng = Selection.Range
            Selection.GoTo what:=wdGoToPage, Count:=Selection.Characters.First.Information(wdActiveEndPageNumber)
            Selection.End = rng.End
            Selection.Bookmarks("\Page").Range.Select
        Case "AlignColumnsSec"
            Set rng = Selection.Range
            rng.Expand (wdSection)
            Selection.GoTo what:=wdGoToPage, Count:=Selection.Characters.First.Information(wdActiveEndPageNumber)
            Selection.Bookmarks("\Page").Range.Select
            If Selection.Start < rng.Start Then
                Selection.Start = rng.Start
            End If
            If Selection.End > rng.End Then
                Selection.End = rng.End
            End If
        Case "AlignColumnsAllSec"
            Selection.Expand wdSection
        Case "AlignColumnsToEnd"
            Selection.GoTo what:=wdGoToPage, Count:=Selection.Characters.First.Information(wdActiveEndPageNumber)
            Selection.EndKey unit:=wdStory, Extend:=wdExtend
        Case "AlignColumnsAll"
            Selection.HomeKey unit:=wdStory
            Selection.EndKey unit:=wdStory, Extend:=wdExtend
    End Select
    
    Call AlignColumns.Repair
    
    SavedRange.Select
End Sub
Sub Footnotes_OnAction(control As IRibbonControl)
        
    Select Case control.ID
        Case "FootnoteButton"
            Footnotes.Add
        Case "FootnotesReplace"
            Footnotes.ReplaceSelection
        Case "FootnotesReplaceAll"
            stopCode = False
            Call CheckDocsSaved
            If stopCode Then Exit Sub
            Footnotes.ReplaceAll
    End Select

End Sub


Sub DecorationsTitlesOption1_OnAction(control As IRibbonControl)
    
    Dim shp As Shape
    Dim shpName As String
    Dim i As Integer
    
    On Error Resume Next
    Select Case control.ID
        Case "DecorationsTitlesCreateOption1"
            SettingsHelper.Save appName, "DecorationsTitles1", "isUpdateOnly", False
            FormDecorationsTitles1.Show vbModeless
        
        Case "DecorationsTitlesUpdateOption1"
            SettingsHelper.Save appName, "DecorationsTitles1", "isUpdateOnly", True
            FormDecorationsTitles1.Show vbModeless
        
        Case "DecorationsTitlesRemoveOption1"
            Application.UndoRecord.StartCustomRecord "הסרת עיטורים לכותרות"
            
            If Selection.Range.ShapeRange.Count = 1 Then
                shpName = Selection.Range.ShapeRange(1).name
            Else
                shpName = Selection.Range.Paragraphs(1).Style & ">" & "עיטור צף>" & "*"
            End If
            
            For i = ActiveDocument.Shapes.Count To 1 Step -1
                Set shp = ActiveDocument.Shapes(i)
                If shp.name Like shpName Then
                    shp.Delete
                End If
            Next i
            Application.UndoRecord.EndCustomRecord
    End Select
    On Error GoTo 0
End Sub
Sub DecorationsTitlesOption2_OnAction(control As IRibbonControl)
    
    Dim inShp As InlineShape
    Dim shpName As String
    Dim rngToRemove As Range
    Dim i As Integer
    
    On Error Resume Next
    Select Case control.ID
        Case "DecorationsTitlesCreateOption2"
            SettingsHelper.Save appName, "DecorationsTitles2", "isUpdateOnly", False
            FormDecorationsTitles2.Show vbModeless
        
        Case "DecorationsTitlesUpdateOption2"
            SettingsHelper.Save appName, "DecorationsTitles2", "isUpdateOnly", True
            FormDecorationsTitles2.Show vbModeless
            
        Case "DecorationsTitlesRemoveOption2"
            Application.UndoRecord.StartCustomRecord "הסרת עיטורים לכותרות"
            
            If Selection.Range.InlineShapes.Count = 1 Then
                shpName = Selection.Range.InlineShapes(1).AlternativeText
            Else
                shpName = Selection.Range.Paragraphs(1).Style & ">" & "עיטור כטקסט>" & "*"
            End If
            
            For i = ActiveDocument.InlineShapes.Count To 1 Step -1
                Set inShp = ActiveDocument.InlineShapes(i)
                If inShp.AlternativeText Like shpName Then
                    Set rngToRemove = inShp.Range
                    If Split(inShp.AlternativeText, ">")(2) Mod 2 <> 0 Then
                        rngToRemove.MoveEnd wdCharacter, 1
                    Else
                        rngToRemove.MoveStart wdCharacter, -1
                    End If
                    rngToRemove.Delete
                End If
            Next i
            Application.UndoRecord.EndCustomRecord
    End Select
    On Error GoTo 0
End Sub
Sub DocReduction_OnAction(control As IRibbonControl)
    
    stopCode = False
    Call CheckDocsSaved
    If stopCode Then Exit Sub
    
    Call DocReduction.Repair

End Sub
Sub PageMarking_OnAction(control As IRibbonControl)
    
    Select Case control.ID
        Case "PageMarkingButton"
            PageMarking.Repair
        Case "PageMarkingSearch"
            PageMarking.BugSearch
        Case "PageMarkingRemove"
            PageMarking.Remove
    End Select

End Sub

Sub BracketsAndFootnotes_OnAction(control As IRibbonControl)
    
    Select Case control.ID
        Case "ConvertToFootnots"
            Call BracketsAndFootnotes.ConvertToFootnots
        Case "ConvertToBrackets"
            Call BracketsAndFootnotes.ConvertToBrackets
    End Select
    
End Sub

Sub CropMarks_OnAction(control As IRibbonControl)

    Select Case control.ID
        Case "CropMarksAdd"
            CropMarks.Add
        Case "CropMarksRemove"
            CropMarks.Remove
    End Select

End Sub
Sub ContinuousFootnotes_OnAction(control As IRibbonControl)
    Select Case control.ID
        Case "ContinuousFootnotes"
                stopCode = False
                Call CheckDocsSaved
                If stopCode Then Exit Sub
            Call ContinuousFootnotes.Repair
        Case "ContinuousFootnotesRemove"
            Call ContinuousFootnotes.Remove
        Case "ContinuousFootnotesAddSeparator"
            Call ContinuousFootnotes.AddSeparator
        Case "ContinuousFootnotesUpdateSeparator"
            Call ContinuousFootnotes.UpdateSeparator
    End Select
End Sub
Sub TableContents_OnAction(control As IRibbonControl)
    Select Case control.ID
        Case "TableContentsButton"
            FormTableContents.Show
            Call TableContents.Creating
            
        Case "TableContentsUpdate"
            Call TableContents.Update
            
        Case "TableContentsRemove"
            Call TableContents.Remove
            
    End Select
End Sub
Sub Typos_OnAction(control As IRibbonControl)
    Select Case control.ID
        Case "TyposCommon"
            stopCode = False
            Call CheckDocsSaved
            If stopCode Then Exit Sub
            FormTyposCommon.Show
        Case "UnclosedParentheses"
            FormUnclosedParentheses.Show vbModeless
            SettingsHelper.Save appName, "UnclosedParentheses", "entireDocument", True
        Case "UnclosedParenthesesEachPara"
            FormUnclosedParentheses.Show vbModeless
            SettingsHelper.Save appName, "UnclosedParentheses", "entireDocument", False
        Case "FixHebrewPunctuation"
            Call Typos.FixHebrewPunctuation
    End Select
End Sub
Sub EditingErrors_OnAction(control As IRibbonControl)
    Select Case control.ID
        Case "EditingErrorPageSize"
            Call EditingErrors.PagesSize
        Case "EditingErrorColumnWidth"
            Call EditingErrors.ColumnWidth
        Case "EditingErrorSwapDocumentStyles"
            Call EditingErrors.SwapDocumentStyles
        Case "EditingErrorDeleteUnusedStyles"
            Call EditingErrors.DeleteUnusedStyles
    End Select

End Sub
Sub SplitDocument_OnAction(control As IRibbonControl)
    Select Case control.ID
        Case "SplitDocument"
            Call SplitDocument.SplitDocumentToMainAndNotes(True, True)
        Case "SplitDocumentFootnots"
            Call SplitDocument.SplitDocumentToMainAndNotes(True, False)
        Case "SplitDocumentEndnots"
            Call SplitDocument.SplitDocumentToMainAndNotes(False, True)
    End Select

End Sub
Sub LineCommentBox_OnAction(control As IRibbonControl)
    Select Case control.ID
        Case "LineCommentBox"
            Call LineCommentBox.Add
        Case "LineCommentBoxShowLine"
            Call LineCommentBox.ShowBoxLine(True)
        Case "LineCommentBoxUnShowLine"
            Call LineCommentBox.ShowBoxLine(False)
        Case "LineCommentBoxRemove"
            Call LineCommentBox.Remove
    End Select
End Sub
Sub TextAlternating_OnAction(control As IRibbonControl)
    Select Case control.ID
        Case "TextAlternating": Call FormTextAlternating.Show
    End Select
End Sub
Sub About_OnAction(control As IRibbonControl)
    
    Select Case control.ID
        Case "UnInstall": Call About.OpenStartupFolder
        Case "About": Call About.OpenAddinLink
    End Select

End Sub

-------------------------------------------------------------------------------
