Option Explicit

Private Const CONFIG_DELIMITER As String = "="
Private m_configPath As String

Public Sub Add(Section As String, Key As String, Value As String)
    If KeyExists(Section, Key) Then
        Call Update(Section, Key, Value)
    Else
        Call AddWrapper(Section, Key, Value)
    End If
End Sub

Private Sub AddWrapper(Section As String, Key As String, Value As String)
    ' Add a new section and key with a value
    Dim file As Object
    Dim tempFilepath As String
    Dim tempFile As Object
    Dim sectionFound As Boolean
    Dim keyFound As Boolean
    Dim currentLine As String
    
    ' Initialize variables
    sectionFound = False
    keyFound = False
    tempFilepath = ThisWorkbook.Path & "\TempConfig.ini"
    ' Open the config file for reading
    Set file = FSO.OpenTextFile(m_configPath, 1)
    ' Open a temporary file for writing
    Set tempFile = FSO.CreateTextFile(tempFilepath, True)
    ' Process each line in the config file
    Do Until file.AtEndOfStream
        currentLine = file.ReadLine()
        ' Write the current line to the temporary file
        tempFile.WriteLine currentLine
        ' Check if the current line matches the section
        If LCase(Trim(currentLine)) = "[" & LCase(Section) & "]" Then
            sectionFound = True
        End If
        ' Check if the section is found and has ended
        If sectionFound Then
            ' Write the new key-value pair to the temporary file
            tempFile.WriteLine Key & CONFIG_DELIMITER & Value
            sectionFound = False
        End If
    Loop
    
    ' Close the files
    file.Close
    tempFile.Close
    ' Replace the original config file with the updated temporary file
    If FSO.FileExists(m_configPath) Then
        FSO.DeleteFile m_configPath
    End If
    FSO.MoveFile tempFilepath, m_configPath
End Sub

Public Sub Delete(Section As String, Key As String)
    ' Delete a specific key within a section
    Dim file As Object
    Dim tempFile As Object
    Dim tempFilepath As String
    Dim sectionFound As Boolean
    Dim keyFound As Boolean
    Dim currentLine As String
    
    ' Initialize variables
    sectionFound = False
    keyFound = False
    
    tempFilepath = ThisWorkbook.Path & "\TempConfig.ini"
    
    ' Open the config file for reading
    Set file = FSO.OpenTextFile(m_configPath, 1)
    
    ' Open a temporary file for writing
    Set tempFile = FSO.CreateTextFile(tempFilepath, True)
    
    ' Process each line in the config file
    Do Until file.AtEndOfStream
        currentLine = file.ReadLine()
        
        ' Check if the current line matches the section
        If LCase(Trim(currentLine)) = "[" & LCase(Section) & "]" Then
            sectionFound = True
        End If
        
        ' Check if the current line matches the key
        If sectionFound And LCase(Trim(currentLine)) Like LCase(Key & "*") Then
            keyFound = True
        Else
            ' Write the current line to the temporary file
            tempFile.WriteLine currentLine
        End If
    Loop
    
    ' Close the files
    file.Close
    tempFile.Close
    
    ' Replace the original config file with the updated temporary file
    If FSO.FileExists(m_configPath) Then
        FSO.DeleteFile m_configPath
    End If
    FSO.MoveFile tempFilepath, m_configPath
    
    ' Display a message if the section or key was not found
    If Not sectionFound Then
        MsgBox "[" & Section & "] section not found in config file", vbCritical + vbOKOnly, "Warning"
    ElseIf Not keyFound Then
        MsgBox "'" & Key & "' key (section : [" & Section & "]) not found in config file", vbCritical + vbOKOnly, "Warning"
    End If
End Sub

Public Sub Update(Section As String, Key As String, NewValue As String)
    ' Update the value of a specific key within a section
    Dim file As Object
    Dim tempFilepath As String
    Dim tempFile As Object
    Dim sectionFound As Boolean
    Dim keyFound As Boolean
    Dim currentLine As String
    Dim outputLine As String
    
    ' Initialize variables
    sectionFound = False
    keyFound = False
    tempFilepath = ThisWorkbook.Path & "\TempConfig.ini"
    
    ' Open the config file for reading
    Set file = FSO.OpenTextFile(m_configPath, 1)
    
    ' Open a temporary file for writing
    Set tempFile = FSO.CreateTextFile(tempFilepath, True)
    
    ' Process each line in the config file
    Do Until file.AtEndOfStream
        currentLine = file.ReadLine()
        
        ' Check if the current line matches the section
        If LCase(Trim(currentLine)) = "[" & LCase(Section) & "]" Then
            sectionFound = True
        End If
        
        ' Check if the current line matches the key
        If sectionFound And LCase(Trim(currentLine)) Like LCase(Key & "*") Then
            ' Write the updated key-value pair to the temporary file
            outputLine = Key & CONFIG_DELIMITER & NewValue
            tempFile.WriteLine outputLine
            keyFound = True
        Else
            ' Write the current line to the temporary file
            tempFile.WriteLine currentLine
        End If
    Loop
    
    ' Close the files
    file.Close
    tempFile.Close

    ' Replace the original config file with the updated temporary file
    If FSO.FileExists(m_configPath) Then
        FSO.DeleteFile m_configPath
    End If
    Call FSO.MoveFile(tempFilepath, m_configPath)
    
    ' Display a message if the section or key was not found
    If Not sectionFound Then
        MsgBox "[" & Section & "] section not found in config file", vbCritical + vbOKOnly, "Warning"
    ElseIf Not keyFound Then
        MsgBox "'" & Key & "' key (section : [" & Section & "]) not found in config file", vbCritical + vbOKOnly, "Warning"
    End If
End Sub

Public Function ReadList(Section As String, Key As String, Optional Delimiter As Variant) As Variant
    'returns array
    Dim result As Variant
    If IsMissing(Delimiter) Then
        result = Split(Read(Section, Key), ",")
    Else
        result = Split(Read(Section, Key), CStr(Delimiter))
    End If
    ReadList = result
End Function

Public Function Read(Section As String, Key As String) As String
    'Read the content
    Dim file As Variant, nextLine As String, keyMessage As Variant, result As Variant, sectionFound As Boolean, keyFound As Boolean, line As Integer
    result = ""
    sectionFound = False
    keyFound = False
    Set file = FSO.OpenTextFile(m_configPath, 1, True)
    Do Until file.AtEndOfStream
        nextLine = file.ReadLine()
        If sectionFound = True Then
            keyMessage = Split(nextLine, CONFIG_DELIMITER)
            If LCase(Trim(keyMessage(0))) = LCase(Key) Then
                If UBound(keyMessage) = 1 Then
                    keyFound = True
                    result = CStr(keyMessage(1))
                    file.Close
                    Exit Do
                End If
            End If
        End If
        If LCase(Trim(nextLine)) = "[" & LCase(Section) & "]" Then
            sectionFound = True
        End If
    Loop
    file.Close
    If Not sectionFound Then
        Call MsgBox("[" & Section & "] section not found in config file", vbCritical + vbOKOnly, "Warning")
    End If
    If Not keyFound Then
        MsgBox "'" & Key & "' key (section : [" & Section & "]) not found in config file", vbCritical + vbOKOnly, "Warning"
    End If
    Read = Trim(result)
End Function

Public Function KeyExists(Section As String, Key As String) As Boolean
    Dim file As Variant, nextLine As String, keyMessage As Variant, sectionFound As Boolean, keyFound As Boolean, line As Integer
    KeyExists = False
    sectionFound = False
    keyFound = False
    Set file = FSO.OpenTextFile(m_configPath, 1, True)
    Do Until file.AtEndOfStream
        nextLine = Trim(file.ReadLine())
        If sectionFound = True Then
            keyMessage = Split(nextLine, CONFIG_DELIMITER)
            If nextLine <> "" Then
                If LCase(Trim(keyMessage(0))) = LCase(Key) Then
                    If UBound(keyMessage) = 1 Then
                        keyFound = True
                        KeyExists = True
                        file.Close
                        Exit Function
                    End If
                End If
            End If
        End If
        If LCase(Trim(nextLine)) = "[" & LCase(Section) & "]" Then
            sectionFound = True
        End If
    Loop
    file.Close
End Function

Private Sub Class_Initialize()
    m_configPath = ThisWorkbook.Path + "\Config.ini"
    If Not FSO.FileExists(m_configPath) Then
        Call FSO.CreateTextFile(m_configPath, False)
    End If
End Sub

Private Function FSO() As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
End Function
