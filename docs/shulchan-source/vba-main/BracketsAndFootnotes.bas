VBA MACRO BracketsAndFootnotes.bas 
in file: word/vbaProject.bin - OLE stream: 'VBA/BracketsAndFootnotes'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit
Option Private Module

Dim BracketsAndFootnotesAllDoc As Boolean
Dim BracketsAndFootnotesEachPara As Boolean

Dim RoundBrackets As Boolean
Dim SquareBrackets As Boolean
Dim CurlyBrackets As Boolean
Dim AngleBrackets As Boolean

Sub OnLoad()
    
    BracketsAndFootnotesAllDoc = SettingsHelper.GetSavedSetting(appName, "BracketsAndFootnotes", "BracketsAndFootnotesAllDoc", True)
    BracketsAndFootnotesEachPara = SettingsHelper.GetSavedSetting(appName, "BracketsAndFootnotes", "BracketsAndFootnotesEachPara", False)
    
    RoundBrackets = SettingsHelper.GetSavedSetting(appName, "BracketsAndFootnotes", "RoundBrackets", True)
    SquareBrackets = SettingsHelper.GetSavedSetting(appName, "BracketsAndFootnotes", "SquareBrackets", False)
    CurlyBrackets = SettingsHelper.GetSavedSetting(appName, "BracketsAndFootnotes", "CurlyBrackets", False)
    AngleBrackets = SettingsHelper.GetSavedSetting(appName, "BracketsAndFootnotes", "AngleBrackets", False)

End Sub
Sub SH_RibbonControl_RngOnAction(control As IRibbonControl, pressed As Boolean)
    
    BracketsAndFootnotesAllDoc = False
    BracketsAndFootnotesEachPara = False
    
    Select Case control.ID
        Case "BracketsAndFootnotesAllDoc": BracketsAndFootnotesAllDoc = True
        Case "BracketsAndFootnotesEachPara": BracketsAndFootnotesEachPara = True
    End Select
    
    SettingsHelper.Save appName, "BracketsAndFootnotes", "BracketsAndFootnotesAllDoc", BracketsAndFootnotesAllDoc
    SettingsHelper.Save appName, "BracketsAndFootnotes", "BracketsAndFootnotesEachPara", BracketsAndFootnotesEachPara

    myRibbon.InvalidateControl "BracketsAndFootnotesAllDoc"
    myRibbon.InvalidateControl "BracketsAndFootnotesEachPara"

End Sub
Sub SH_RibbonControl_OnAction(control As IRibbonControl, pressed As Boolean)
    
    RoundBrackets = False
    SquareBrackets = False
    CurlyBrackets = False
    AngleBrackets = False
    
    Select Case control.ID
        Case "RoundBrackets": RoundBrackets = True
        Case "SquareBrackets": SquareBrackets = True
        Case "CurlyBrackets": CurlyBrackets = True
        Case "AngleBrackets": AngleBrackets = True
    End Select
    
    SettingsHelper.Save appName, "BracketsAndFootnotes", "RoundBrackets", RoundBrackets
    SettingsHelper.Save appName, "BracketsAndFootnotes", "SquareBrackets", SquareBrackets
    SettingsHelper.Save appName, "BracketsAndFootnotes", "CurlyBrackets", CurlyBrackets
    SettingsHelper.Save appName, "BracketsAndFootnotes", "AngleBrackets", AngleBrackets
    
    myRibbon.InvalidateControl "RoundBrackets"
    myRibbon.InvalidateControl "SquareBrackets"
    myRibbon.InvalidateControl "CurlyBrackets"
    myRibbon.InvalidateControl "AngleBrackets"

End Sub
Sub OnGetPressed(control As IRibbonControl, ByRef returnedVal)
    
    Select Case control.ID
        
        Case "BracketsAndFootnotesAllDoc": returnedVal = BracketsAndFootnotesAllDoc
        Case "BracketsAndFootnotesEachPara": returnedVal = BracketsAndFootnotesEachPara
    
        Case "RoundBrackets": returnedVal = RoundBrackets
        Case "SquareBrackets": returnedVal = SquareBrackets
        Case "CurlyBrackets": returnedVal = CurlyBrackets
        Case "AngleBrackets": returnedVal = AngleBrackets

    End Select
    
End Sub
Function GetBracketsType(openBrackets As Boolean) As String
    
    If openBrackets Then
        If RoundBrackets Then GetBracketsType = "("
        If SquareBrackets Then GetBracketsType = "["
        If CurlyBrackets Then GetBracketsType = "{"
        If AngleBrackets Then GetBracketsType = "<"
    Else
        If RoundBrackets Then GetBracketsType = ")"
        If SquareBrackets Then GetBracketsType = "]"
        If CurlyBrackets Then GetBracketsType = "}"
        If AngleBrackets Then GetBracketsType = ">"
    End If
    
End Function
Sub ConvertToFootnots()
    
    Dim rng As Range
    Dim parasCount As Integer
    Dim paraRange As Range
    Dim bracketsRng As Range
    Dim fnt As Footnote
    Dim text As String
    Dim i As Integer
    
'    On Error GoTo Ending
    Application.UndoRecord.StartCustomRecord "המרת סוגריים להערות"
    Application.ScreenUpdating = False
    
    Set rng = Selection.Range
    parasCount = rng.Paragraphs.Count
    
    FormRunning1.Show vbModeless
    
    For i = 1 To parasCount
    
        Set paraRange = rng.Paragraphs(i).Range
        FormRunning1.Label1.Caption = "ממיר סוגריים להערות פיסקה " & i & " מתוך " & parasCount
        DoEvents
        
        Do
            Set bracketsRng = SearchBrackets(paraRange.Duplicate)
            If bracketsRng Is Nothing Then Exit Do
            text = bracketsRng.text
            text = Right(text, Len(text) - 1)
            text = Left(text, Len(text) - 1)
            
            Set fnt = bracketsRng.Footnotes.Add(bracketsRng, text:=text)
            
            bracketsRng.Delete
        Loop
    Next i

Ending:
    Application.UndoRecord.EndCustomRecord
    Application.ScreenUpdating = True
    Unload FormRunning1
End Sub
Function ConvertToBrackets()
    
    Dim rng As Range
    Dim ftn As Footnote
    Dim ftnRng As Range
    Dim text As String
    Dim i As Integer
    
    On Error GoTo Ending
    Application.UndoRecord.StartCustomRecord "המרת הערות לסוגריים"
    Application.ScreenUpdating = False
    
    Set rng = Selection.Range
    
    For i = rng.Footnotes.Count To 1 Step -1
    
        Set ftn = rng.Footnotes(i)
        
        Set ftnRng = ftn.Reference
        text = " " & GetBracketsType(True) & ftn.Range.text & GetBracketsType(False)
        
        ftnRng.Delete
        ftnRng.text = text
        
    Next i

Ending:

    Application.UndoRecord.EndCustomRecord
    Application.ScreenUpdating = True
    
End Function

Function SearchBrackets(rng As Range) As Range
    
    Dim opens As Integer
    Dim pos As Long
    
    With rng
            
            .Collapse wdCollapseStart
            .MoveUntil GetBracketsType(True) & GetBracketsType(False) & Chr(13)
                    
        Do
            .MoveEndUntil GetBracketsType(True) & GetBracketsType(False) & Chr(13)
            .MoveEnd wdCharacter, 1
            
            If .text = Chr(13) Then
                Exit Do
            ElseIf .Characters.Last = GetBracketsType(True) Then
                opens = opens + 1
            ElseIf .Characters.Last = GetBracketsType(False) And opens > 0 Then
                opens = opens - 1
                If opens = 0 Then Set SearchBrackets = rng: Exit Do
            Else
                Exit Do
            End If
        Loop
    
    End With
    
End Function
-------------------------------------------------------------------------------
