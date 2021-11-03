Option Explicit

'Created by Mustafa Can Ozturk on 08.02.2021
'===========================================
'The functions below will be used for file operations
'It includes the most common functions
'It needs to be used with "File" module
'Functions:
'Copy , Move, Delete, Exist, Rename, BuiltPath, ParentFolder, GetPath, GetPaths, OpenZipFile
'CreateTXT, ReadTXT, LogTo, CreateLog, TerminateLog, IsTXTOpen, IdenticalFiles
'Properties:
'Name, Extension
'===========================================


'=======Auxiliary Functions=======

Private ProcessDuration As Double



Private Function FSO() As Object
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
End Function
'=======End of Auxiliary Functions=======

Public Property Get ParentFolder(Path As String) As String
    
    ParentFolder = FSO.GetParentFolderName(Path) + DetectSlash(Path)
    
End Property

Public Property Get Extension(byval Path As String) As String
    
    If Exist(Path) = True Then
        
        Extension = FSO.GetExtensionName(Path)
    
    Else
    
        Extension = "NOT FOUND"
        
    End If
        
End Property

Public Function Exist(byval FilePath As String) As Boolean

    If FSO.fileexists(FilePath) = True Then
        
        Exist = True
    
    Else
    
        Exist = False
        
    End If
    
End Function

Public Sub Rename(Path As String, NewName As String)
    
    Dim baseFolder As String
    Dim defaultExtension As String
    
    If Exist(Path) = True Then
        
        baseFolder = ParentFolder(Path)
        defaultExtension = "." + Extension(Path)
        
        If hasExtension(NewName) = False Then
        
            Name Path As baseFolder + NewName + defaultExtension
        
        Else
        
            Name Path As baseFolder + NewName
            
        End If
    
    End If

End Sub

Private Function hasExtension(FileName As String) As Boolean
    
    If Right(FileName, 4) Like ".???" Then
    
        hasExtension = True
    
    Else
    
        hasExtension = False
        
    End If
    
End Function


Public Function Move(FromPath As String, ToPathWithDifferentName As String) As Boolean
    
    If Exist(FromPath) = False Then
    
        Move = False
        Exit Function
    
    End If
    
    FSO.Move FromPath, ToPathWithDifferentName
    
    Move = True
    
End Function

Public Function Copy(Source As String, Destination As String) As Boolean
    
    If Exist(Source) = False Then
    
        Copy = False
        Exit Function
        
    End If
    
    If Right(Destination, 1) <> "\" Or Right(Destination, 1) <> "/" Then
        
        Destination = Destination + DetectSlash(Destination)
    
    End If
    
    FSO.Copy Source, Destination
    
    Copy = True
    
End Function

Private Function DetectSlash(Path As String) As String
    
    If InStr(Path, "\") > 0 Then
    
        DetectSlash = "\"
        
    ElseIf InStr(Path, "/") > 0 Then
    
        DetectSlash = "/"
        
    End If
    
End Function

Public Sub Delete(Path As String)
    
    If Exist(Path) = True Then
    
        FSO.DeleteFile Path
    
    End If
    
End Sub

Public Property Get Name(byval Path As String) As String
    
    If Exist(Path) = True Then
    
        Name = FSO.GetFileName(Path)
        
    Else
        
        Name = "NOT FOUND"
        
    End If

End Property

Public Function BuildPath(Path As String, FileName As String) As String

    BuildPath = FSO.BuildPath(Path, FileName)

End Function

Public Function GetPath(FileDescription As String, FileExtension As String, WindowHeader As String) As String
    
    Dim fd As FileDialog
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    Dim ExcelFile As Variant
    
    With fd
        .Title = WindowHeader
        .ButtonName = "Select"
        .Filters.Clear
        .Filters.Add FileDescription, FileExtension
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            GetPath = .SelectedItems(1)
        Else
            GetPath = ""
        End If
    
    End With

End Function
                            
Public Function GetFolderPath(Title As String)

Dim fd As FileDialog

Set fd = Application.FileDialog(msoFileDialogFolderPicker)

With fd
    .AllowMultiSelect = False
    .ButtonName = "Select"
    .Title = Title
    
    If .Show = -1 Then
        GetFolderPath = .SelectedItems(1)
    Else
        GetFolderPath = ""
    End If
    
End With

End Function
                                
Public Function GetPaths(FileDescription As String, FileExtension As String, WindowHeader As String) As Variant
    
    Dim fd As FileDialog
    Dim i As Long
    Dim files() As Variant
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .Title = WindowHeader
        .ButtonName = "Select"
        .Filters.Clear
        .Filters.Add FileDescription, FileExtension
        .AllowMultiSelect = True
        
        If .Show = -1 Then
            For i = 1 To .SelectedItems.Count
            ReDim Preserve files(i)
            files(i) = .SelectedItems(i)
            Next i
        Else
            ReDim Preserve files(1)
            files(1) = ""
        End If
    
    End With
    
    GetPaths = files

End Function

Public Sub OpenZipFile(FilePath)

    Shell "Explorer.exe /e, " & FilePath, vbNormalFocus
    
End Sub

Public Function IsPath(byval Path As String) As Boolean
    
    If InStr(Path, "/") > 0 Or InStr(Path, "\") > 0 Then
    
        IsPath = True
        
    Else
    
        IsPath = False
        
    End If
    
End Function

Public Sub CreateTXT(DataArray As Variant, FullPath As String)

    Dim t As Object
    Dim i As Long, j As Long
    
    
    If IsPath(FullPath) = False Then
        Debug.Print "CreateTXT error : Enter a valid path"
        Exit Sub
    End If
    
    Set t = FSO.CreateTextFile(FullPath, True)
    
    For i = LBound(DataArray, 1) To UBound(DataArray, 1)
    
        For j = LBound(DataArray, 2) To UBound(DataArray, 2)
        
            If j = UBound(DataArray, 2) Then
                t.Write (DataArray(i, j))
            Else
                t.Write (DataArray(i, j) & "<:>")
            End If
            
        Next j
        
        t.WriteLine ("")
        
    Next i
    
End Sub

Public Sub ReadTXT(FullPath As String, ToSheet As Worksheet, Delimiter As String, Optional IncludeHeader As Boolean = True, Optional InitialRow As Variant, Optional InitialColumn As Variant)
'FSO can read only ASCII, not UTF-8
    Dim fs As Object
    Dim textLine As String
    Dim lineArr As Variant
    Dim i As Long, j As Long, k As Long, m As Long
    Dim row As Long, column As Long
    Dim ignoreFirstLine As Boolean
    
    Set fs = FSO.OpenTextFile(FullPath, 1, 0) '1 is for reading, 0 is for ASCII (ForAppending, TristateFalse)
    
    If IsMissing(InitialRow) = True Then
        row = 1
    Else
        row = InitialRow
    End If
    
    If IsMissing(InitialColumn) = True Then
        column = 1
    Else
        column = InitialColumn
    End If
    
    
    If IncludeHeader = True Then
        ignoreFirstLine = False
    Else
        ignoreFirstLine = True
    End If
    
    Do While Not fs.AtEndOfStream
                
        If ignoreFirstLine = True Then
            textLine = fs.readline
            ignoreFirstLine = False
        End If
        
        textLine = fs.readline
        lineArr = Split(textLine, Delimiter)
        j = UBound(lineArr) + 1
        m = 0
        For k = column To j + column - 1

            ToSheet.Cells(row, k) = lineArr(m)
            m = m + 1
            
        Next k
        row = row + 1
        
    Loop
    
    fs.Close
        
End Sub

Public Sub ReadTXTADO(FullPath As String, ToSheet As Worksheet, Delimiter As String, Optional IncludeHeader As Boolean = True)
'ADO can read UTF-8
    Dim objStream, strData As String

    Set objStream = CreateObject("ADODB.Stream")
    FastMode
    objStream.Charset = "utf-8"
    objStream.Open
    objStream.LoadFromFile (FullPath)
    
    strData = objStream.ReadText()
    
    objStream.Close
    'Import Text To cells
    Dim i As Long, j As Long
    Dim m As Long, n As Long
    Dim k As Long, l As Long
    Dim data As Variant, dataNew As Variant

    data = Split(strData, vbNewLine)
    
    If IncludeHeader = True Then
        i = LBound(data)
    Else
        i = LBound(data) + 1
    End If
    
    j = UBound(data)

    For k = i To j
        
        dataNew = Split(data(k), Delimiter)
        m = LBound(dataNew)
        n = UBound(dataNew)
        
        For l = m To n
            ToSheet.Cells(k + 1, l + 1) = dataNew(l)
        Next l
    
    Next k
    
    If IncludeHeader = False Then ToSheet.Range("A1").EntireRow.Delete
    ToSheet.UsedRange.Columns.AutoFit
    
    NormalMode
    
    Set objStream = Nothing
    
End Sub

Public Sub LogTo(ByVal Path As String, Text As String, Optional ErrorMessage As Boolean = False)
    
    Do
    DoEvents
    Loop Until IsTXTOpen(Path) = False
    
    Open Path For Append As #1
        If ErrorMessage = True Then
            Print #1, String(Len(CStr(Now) & "   ->   " & Text), "=")
        End If
        Print #1, CStr(Now) & "   ->   " & Text
        If ErrorMessage = True Then
            Print #1, String(Len(CStr(Now) & "   ->   " & Text), "=")
        End If
        Print #1, vbNewLine
    Close #1
    
End Sub


Public Function IsTXTOpen(ByVal FileName As String) As Boolean
    Dim iFilenum As Long
    Dim iErr As Long
     
    On Error Resume Next
    iFilenum = FreeFile()
    Open FileName For Input Lock Read As #iFilenum
    Close iFilenum
    iErr = Err
    On Error GoTo 0
     
    Select Case iErr
    Case 0:    IsTXTOpen = False
    Case 70:   IsTXTOpen = True
    Case Else: Error iErr
    End Select
     
End Function

                
Public Sub CreateLog(FilePath As String)
    
    Dim f As Object
 
    Set f = FSO.CreateTextFile(FilePath, True)
    
    StartTimer
    
    Set f = Nothing
    
    LogTo FilePath, UCase("Process has been started")
    
End Sub

Private Sub StartTimer()

    ProcessDuration = Timer
    
End Sub
                    
Public Sub TerminateLog(FilePath As String)

    file.LogTo FilePath, UCase("Process has been completed in " & ProcessTime & " seconds.")
    
End Sub

Private Function ProcessTime() As String
    
    ProcessTime = CStr(Format(Timer - ProcessDuration, "00.00"))
    
End Function
                    
Public Function IdenticalFiles(strFilename1 As String, strFilename2 As String) As Boolean

    Dim byt1() As Byte
    Dim byt2() As Byte
    Dim f1 As Integer
    Dim f2 As Integer
    Dim lngFileLen1 As Long
    Dim lngFileLen2 As Long
    Dim i As Long

    'Test to see if we have actually been passed 2 filenames
    If LenB(strFilename1) = 0 Or LenB(strFilename2) = 0 Then Exit Function
    'Test to see if the first file exists
    If LenB(Dir(strFilename1)) = 0 Then Exit Function
    'Test to see if the second file exists
    If LenB(Dir(strFilename2)) = 0 Then Exit Function

    'OK now start looking at the file contents
    f1 = FreeFile
    Open strFilename1 For Binary Access Read As #f1
    f2 = FreeFile
    Open strFilename2 For Binary Access Read As #f2
    lngFileLen1 = LOF(f1)
    lngFileLen2 = LOF(f2)
    If lngFileLen1 = lngFileLen2 Then
      'Continue - there is a possibility they are the same
      ReDim byt1(1 To lngFileLen1) As Byte
      ReDim byt2(1 To lngFileLen2) As Byte
      Get #f1, , byt1
      Get #f2, , byt2
      For i = 1 To lngFileLen1
        If byt1(i) <> byt2(i) Then GoTo IdenticalFiles_Exit 'The 2 files are not the same
      Next
      'We got this far so the 2 files must be the same
      IdenticalFiles = True
    End If

    IdenticalFiles_Exit:
    Close #f1
    Close #f2

End Function
