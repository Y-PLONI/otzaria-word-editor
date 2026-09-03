VBA MACRO DocReduction.bas 
in file: word/vbaProject.bin - OLE stream: 'VBA/DocReduction'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit
Option Private Module

Public pageCountTarget As Integer
Public marginsTarget As Integer
Public paraSpaceTarget As Integer
Public linesSpaceTarget As Integer
Public fontSizeTarget As Integer
Public fontLimitTarget As Integer

Sub Repair()
    
    Dim SavedRange As Range
    Dim rng As Range
    Dim PageCountTest As Integer
    
    Dim MarginsTest As Integer
    Dim ParaSpaceTest As Integer
    Dim LinesSpaceTest As Integer
    Dim FontSizeTest As Integer
    Dim FontLimitTest As Integer
    
    Dim word As Range
    Dim para As Paragraph
    Dim sec As section
    Dim col As Column
    
    Dim i As Integer
    Dim msg(1 To 4) As Integer
    
    PageCountTest = ActiveDocument.ComputeStatistics(wdStatisticPages)
'    If PageCountTest > 100 Then
'        MsgBox "...לצמצם כזה מסמך גדול?! נו באמת" & vbNewLine & "לא ניתן לצמצם מסמך יותר מ100 עמודים מחשש לתקיעת המחשב", vbOKOnly, "...אופס"
'        Exit Sub
'    End If
    FormDocReduction.Show
'    If pageCountTarget = 0 Then Exit Sub
'    If pageCountTarget < PageCountTest / 4 Then
'        MsgBox "נראה שאתה מנסה לצמצם את המסמך באופן דרסטי" & vbNewLine & "%לא ניתן (וגם לא מומלץ) לצמצם מסמך ביותר מ75" & vbNewLine & vbNewLine & "אם זאת ניתן לבצע את הצימצום ב2 פעימות", vbOKOnly, "...אופס"
'        Exit Sub
'    End If
    
    On Error GoTo Ending
    Application.ScreenUpdating = False
    Application.UndoRecord.StartCustomRecord
    Set SavedRange = Selection.Range
    
    Set rng = ActiveDocument.Range
    FormRunning.Show vbModeless
    
    For i = 1 To 20
    
        MarginsTest = MarginsTest + 1
        FontSizeTest = FontSizeTest + 1
        ParaSpaceTest = ParaSpaceTest + 1
        LinesSpaceTest = LinesSpaceTest + 1
    
        If PageCountTest <= pageCountTarget Then Exit For
        
        FormRunning.Label1 = "כמות העמודים הנוכחית היא: " & PageCountTest
        DoEvents
        If stopCode Then GoTo Ending
        
        ' שוליים
        If MarginsTest = marginsTarget Then
            For Each sec In rng.Sections
                DoEvents: If stopCode Then GoTo Ending
                With sec.PageSetup
                    
                    If .TopMargin > 28.35 Then .TopMargin = .TopMargin * 0.9
                    If .TopMargin < 28.35 Then .TopMargin = 28.35
                    If .BottomMargin > 28.35 Then .BottomMargin = .BottomMargin * 0.9
                    If .BottomMargin < 28.35 Then .BottomMargin = 28.35
                    If .LeftMargin > 28.35 Then .LeftMargin = .LeftMargin * 0.9
                    If .LeftMargin < 28.35 Then .LeftMargin = 28.35
                    If .RightMargin > 28.35 Then .RightMargin = .RightMargin * 0.9
                    If .RightMargin < 28.35 Then .RightMargin = 28.35
                    
                    
                    With .TextColumns
                        .EvenlySpaced = True
                        If .Count > 1 And .Spacing > 28.35 Then .Spacing = .Spacing * 0.9
                        If .Spacing < 28.35 Then .Spacing = 28.35
                    End With
                End With
            Next sec
            msg(1) = msg(1) + 1
            PageCountTest = ActiveDocument.ComputeStatistics(wdStatisticPages)
            FormRunning.Label1 = "כמות העמודים הנוכחית היא: " & PageCountTest
            DoEvents
            If stopCode Then GoTo Ending
            MarginsTest = 0
        End If
        
        If PageCountTest <= pageCountTarget Then Exit For
            
        ' מרווח בין פיסקאות
        If ParaSpaceTest = paraSpaceTarget Then
            For Each para In rng.Paragraphs
                DoEvents: If stopCode Then GoTo Ending
                With para
                    If .SpaceBefore > 2 Then .SpaceBefore = .SpaceBefore * 0.9
                    If .SpaceAfter > 2 Then .SpaceAfter = .SpaceAfter * 0.9
                End With
            Next para
            msg(2) = msg(2) + 1
            PageCountTest = ActiveDocument.ComputeStatistics(wdStatisticPages)
            FormRunning.Label1 = "כמות העמודים הנוכחית היא: " & PageCountTest
            DoEvents
            If stopCode Then GoTo Ending
            ParaSpaceTest = 0
        End If
        
        If PageCountTest <= pageCountTarget Then Exit For
        
        ' מרווח בין שורות
        If LinesSpaceTest = linesSpaceTarget Then
            For Each para In rng.Paragraphs
                DoEvents: If stopCode Then GoTo Ending
                With para
                    If .LineSpacing > 13.2 Then
                        .LineSpacing = .LineSpacing * 0.9
                    ElseIf .LineSpacing > 12 Then
                        .LineSpacing = 12
                    End If
                End With
            Next para
            msg(3) = msg(3) + 1
            PageCountTest = ActiveDocument.ComputeStatistics(wdStatisticPages)
            FormRunning.Label1 = "כמות העמודים הנוכחית היא: " & PageCountTest
            DoEvents
            If stopCode Then GoTo Ending
            LinesSpaceTest = 0
        End If
        
        If PageCountTest <= pageCountTarget Then Exit For
        
        ' גודל גופן
        If FontSizeTest = fontSizeTarget Then
            For Each word In rng.Words
                DoEvents: If stopCode Then GoTo Ending
                If word.Font.SizeBi > fontLimitTarget Then
                    word.Font.SizeBi = word.Font.SizeBi - 1
                Else: fontSizeTarget = 0
                End If
                If word.Font.Size > fontLimitTarget Then
                    word.Font.Size = word.Font.Size - 1
                Else: fontSizeTarget = 0
                End If
            Next word
            If fontSizeTarget > 0 Then msg(4) = msg(4) + 1
            PageCountTest = ActiveDocument.ComputeStatistics(wdStatisticPages)
            FormRunning.Label1 = "כמות העמודים הנוכחית היא: " & PageCountTest
            DoEvents
            If stopCode Then GoTo Ending
            FontSizeTest = 0
        End If
                
        If PageCountTest <= pageCountTarget Then Exit For
    
    Next i
    MsgBox vbNewLine & "הפעולה הושלמה בהצלחה!!" & vbNewLine & vbNewLine & _
        "כמות העמודים הנוכחית היא: " & PageCountTest _
        & vbNewLine & "השולים הוקטנו ב: " & 100 - ((0.9 ^ msg(1)) * 100) & "%" _
        & vbNewLine & "המרווח בין הפיסקאות צומצם ב: " & 100 - ((0.9 ^ msg(2)) * 100) & "%" _
        & vbNewLine & "המרווח בין השורות צומצם ב: " & 100 - ((0.9 ^ msg(3)) * 100) & "%" _
        & vbNewLine & "'הפונט הוקטן ב: " & msg(4) & " נק", vbOKOnly, "?אז מה היה לנו כאן"
        
Ending:
    SavedRange.Select
    Application.UndoRecord.EndCustomRecord
    Application.ScreenUpdating = True
    Unload FormRunning
End Sub

-------------------------------------------------------------------------------
