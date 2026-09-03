VBA MACRO SettingsHelper.bas 
in file: word/vbaProject.bin - OLE stream: 'VBA/SettingsHelper'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit
Option Private Module

Sub Save(appName As String, section As String, key As String, value As Variant)
    Dim folderPath As String
    Dim FilePath As String
        
    If IsMac Then
        folderPath = Environ("HOME") & "/Library/Application Support/" & appName
        FilePath = folderPath & "/settings.ini"
    Else
        folderPath = Environ("USERPROFILE") & "\AppData\Roaming\" & appName
        FilePath = folderPath & "\settings.ini"
    End If
    
    ' Ensure directory exists
    If Dir(folderPath, vbDirectory) = "" Then
        MkDir folderPath
    End If


    Dim lines As Collection
    Set lines = New Collection
    
    Dim fileExists As Boolean
    fileExists = (Dir(FilePath) <> "")
    
    If fileExists Then
        ' Read the file into memory
        Dim fileNum As Integer
        fileNum = FreeFile
        Open FilePath For Input As fileNum
        
        Dim fileLine As String
        Dim currentSection As String
        Dim sectionFound As Boolean
        sectionFound = False
        currentSection = ""
        
        Do Until EOF(fileNum)
            Line Input #fileNum, fileLine
            
            If Left(fileLine, 1) = "[" And Right(fileLine, 1) = "]" Then
                currentSection = Mid(fileLine, 2, Len(fileLine) - 2)
            End If

            ' If we are in the correct section and find the key, update it
            If currentSection = section Then
                If InStr(fileLine, key & "=") = 1 Then
                    fileLine = key & "=" & value & ";" & TypeName(value)
                    sectionFound = True
                End If
            End If
            lines.Add fileLine
        Loop
        Close fileNum
        
        ' If the section was found but the key wasn't, add the new key-value pair
        If Not sectionFound Then
            lines.Add "[" & section & "]"
            lines.Add key & "=" & value & ";" & TypeName(value)
        End If
    Else
        ' If file doesn't exist, start a new one
        lines.Add "[" & section & "]"
        lines.Add key & "=" & value & ";" & TypeName(value)
    End If
    
    ' Write the lines back to the file
    fileNum = FreeFile
    Open FilePath For Output As fileNum
    
    Dim line As Variant
    For Each line In lines
        Print #fileNum, line
    Next line
    
    Close fileNum
End Sub

Function GetSavedSetting(appName As String, section As String, key As String, Optional defaultValue As Variant) As Variant
    Dim folderPath As String
    Dim FilePath As String
        
    If IsMac Then
        folderPath = Environ("HOME") & "/Library/Application Support/" & appName
        FilePath = folderPath & "/settings.ini"
    Else
        folderPath = Environ("USERPROFILE") & "\AppData\Roaming\" & appName
        FilePath = folderPath & "\settings.ini"
    End If
    
    If Dir(FilePath) = "" Then
        GetSavedSetting = defaultValue
        Exit Function
    End If

    Dim fileNum As Integer
    Dim fileLine As String
    Dim currentSection As String
    fileNum = FreeFile

    Open FilePath For Input As fileNum
    currentSection = ""

    Do Until EOF(fileNum)
        Line Input #fileNum, fileLine
        
        If Left(fileLine, 1) = "[" And Right(fileLine, 1) = "]" Then
            currentSection = Mid(fileLine, 2, Len(fileLine) - 2)
        ElseIf currentSection = section Then
            If InStr(fileLine, key & "=") = 1 Then
                Dim valueParts() As String
                valueParts = Split(Mid(fileLine, Len(key) + 2), ";") ' Split value and type

                Dim loadedValue As String
                loadedValue = valueParts(0)
                Dim valueType As String
                valueType = valueParts(1)

                Select Case valueType
                    Case "String"
                        GetSavedSetting = loadedValue
                    Case "Integer"
                        GetSavedSetting = CInt(loadedValue)
                    Case "Long"
                        GetSavedSetting = CLng(loadedValue)
                    Case "Boolean"
                        GetSavedSetting = CBool(loadedValue)
                    Case "Double"
                        GetSavedSetting = CDbl(loadedValue)
                    Case "Single"
                        GetSavedSetting = CSng(loadedValue)
                    Case Else
                        GetSavedSetting = defaultValue
                End Select

                Close fileNum
                Exit Function
            End If
        End If
    Loop
    
    Close fileNum
    GetSavedSetting = defaultValue
End Function

Function IsSameGroup(str1 As String, str2 As String) As Boolean
    ' Check if lengths differ by more than 1, return False immediately
    If Len(str1) <> Len(str2) Or str1 = str2 Then
        IsSameGroup = False
        Exit Function
    End If
    
    ' Compare the substrings excluding the last character
    If Left(str1, Len(str1) - 1) = Left(str2, Len(str2) - 1) Then
        IsSameGroup = True
    Else
        IsSameGroup = False
    End If
End Function

Function IsMac() As Boolean
    On Error Resume Next
    IsMac = (Environ("HOME") <> "")
    On Error GoTo 0
End Function
-------------------------------------------------------------------------------
