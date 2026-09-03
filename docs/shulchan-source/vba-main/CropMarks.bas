VBA MACRO CropMarks.bas 
in file: word/vbaProject.bin - OLE stream: 'VBA/CropMarks'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit
Option Private Module

Function GetCropMarksName() As String
    
    GetCropMarksName = "CropMarks"
    
End Function

Sub Add()

    Dim sec As section
    Dim ftr As HeaderFooter
    Dim addLine As Shape
    Dim text As String
    Dim milimeters As Double
    Dim Balance As Double
    Dim Balance2 As Double
    Dim startLineY As Variant
    Dim startLineX As Variant
    Dim endLineY As Variant
    Dim endLineX As Variant
    Dim i As Integer
    
    On Error GoTo Ending
    Application.UndoRecord.StartCustomRecord "הוספת סימני חיתוך"
    Application.ScreenUpdating = False
    
    If SearchCropMarks Then
        MsgBox "קיימים כבר סימני חיתוך במסמך זה", vbOKOnly, "שגיאה"
        GoTo Ending
    End If
    text = InputBox("הזן רוחב סימני חיתוך במילימטרים", "הכנס נתונים")
    
    If TextToNumIsInRange(text, 5, 50, msg:=True) Then
        milimeters = CentimetersToPoints(text) / 10
    Else
        GoTo Ending
    End If
    
    Balance = milimeters / 5
    Balance2 = Balance / 1.4
    
    FormRunning1.Show vbModeless
    
    ' הרחבת העמוד והשולים
    For i = 1 To ActiveDocument.Sections.Count
        Set sec = ActiveDocument.Sections(i)
        FormRunning1.Label1.Caption = "מגדיל עמוד ושולים מקטע " & i & " מתוך " & ActiveDocument.Sections.Count: DoEvents
        With sec.PageSetup
            .PageHeight = .PageHeight + (milimeters * 2)
            .PageWidth = .PageWidth + (milimeters * 2)
            .TopMargin = .TopMargin + milimeters
            .BottomMargin = .BottomMargin + milimeters
            .RightMargin = .RightMargin + milimeters
            .LeftMargin = .LeftMargin + milimeters
            .HeaderDistance = .HeaderDistance + milimeters
            .FooterDistance = .FooterDistance + milimeters

        End With
    Next i
    
    With ActiveDocument.Sections(1).PageSetup
        startLineY = Array(0 + Balance, 0 + Balance, milimeters, milimeters, .PageHeight - milimeters, .PageHeight - milimeters, .PageHeight - Balance, .PageHeight - Balance)
        startLineX = Array(milimeters, .PageWidth - milimeters, 0 + Balance, .PageWidth - Balance, 0 + Balance, .PageWidth - Balance, milimeters, .PageWidth - milimeters)
        endLineY = Array(milimeters - Balance2, milimeters - Balance2, milimeters, milimeters, .PageHeight - milimeters, .PageHeight - milimeters, .PageHeight - milimeters + Balance2, .PageHeight - milimeters + Balance2)
        endLineX = Array(milimeters, .PageWidth - milimeters, milimeters - Balance2, .PageWidth - milimeters + Balance2, milimeters - Balance2, .PageWidth - milimeters + Balance2, milimeters, .PageWidth - milimeters)
    End With
    For Each ftr In ActiveDocument.Sections(1).Footers
        FormRunning1.Label1.Caption = "מוסיף סימני חיתוך": DoEvents
        For i = 0 To 7
            Set addLine = ftr.Shapes.addLine(startLineX(i), startLineY(i), endLineX(i), endLineY(i))
            With addLine
                .line.Weight = milimeters / 100
                .line.ForeColor = 0
                .name = PointsToCentimeters(milimeters) & " " & GetCropMarksName
            End With
        Next i
    Next ftr

Ending:
    Application.UndoRecord.EndCustomRecord
    Application.ScreenUpdating = True
    Unload FormRunning1
End Sub
Sub Remove()
    
    Dim sec As section
    Dim ftr As HeaderFooter
    Dim line As Shape
    Dim lineName As String
    Dim milimeters As Double
    Dim i As Integer
    
    On Error GoTo Ending
    Application.UndoRecord.StartCustomRecord "הסרת סימני חיתוך"
    Application.ScreenUpdating = False
    FormRunning1.Show vbModeless
    
    For Each ftr In ActiveDocument.Sections(1).Footers
        FormRunning1.Label1.Caption = "מסיר סימני חיתוך": DoEvents
        For i = ftr.Shapes.Count To 1 Step -1
            Set line = ftr.Shapes(i)
            With line
                If .name Like "*" & GetCropMarksName & "*" Then
                    lineName = Replace(.name, " " & GetCropMarksName, "")
                    .Delete
                End If
            End With
        Next i
    Next ftr
    
    milimeters = CInt(CentimetersToPoints(lineName))
    
    For i = 1 To ActiveDocument.Sections.Count
        Set sec = ActiveDocument.Sections(i)
        FormRunning1.Label1.Caption = "מקטין עמוד ושולים מקטע " & i & " מתוך " & ActiveDocument.Sections.Count: DoEvents
        With sec.PageSetup
            .PageHeight = .PageHeight - (milimeters * 2)
            .PageWidth = .PageWidth - (milimeters * 2)
            .TopMargin = .TopMargin - milimeters
            .BottomMargin = .BottomMargin - milimeters
            .RightMargin = .RightMargin - milimeters
            .LeftMargin = .LeftMargin - milimeters
            .HeaderDistance = .HeaderDistance - milimeters
            .FooterDistance = .FooterDistance - milimeters
        End With
    Next i
    
Ending:
    Application.UndoRecord.EndCustomRecord
    Application.ScreenUpdating = True
    Unload FormRunning1
    
End Sub
Function ReplaceInString(text As String, charsToRemove As String) As String
    
    Dim i As Integer
    ReplaceInString = text
    For i = 1 To Len(charsToRemove)
        ReplaceInString = Replace(ReplaceInString, Right(Left(charsToRemove, i), 1), "")
    Next i
End Function
Function SearchCropMarks() As Boolean

    Dim shp As Shape
    
    For Each shp In ActiveDocument.Sections(1).Footers(1).Shapes
    
        If shp.name Like "*" & GetCropMarksName & "*" Then
            
            SearchCropMarks = True
            Exit Function
        
        End If
        
    Next shp

End Function
-------------------------------------------------------------------------------
