Option Explicit

'Created by Mustafa Can Ozturk on 08.02.2021
'===========================================
'The functions below will be used for file operations
'It includes the most common functions
'It needs to be used with "File" module
'Functions:
'Copy , Move, Delete, Exist, Rename, BuiltPath, ParentFolder, GetPath, GetPaths, OpenZipFile
'CreateTXT, ReadTXT, LogTo
'Properties:
'Name, Extension
'===========================================


'=======Auxiliary Functions=======
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

Public Sub LogTo(Path As String, Text As String)

    Open Path For Append As #1
        'Print #1, vbNewLine
        Print #1, Text
    Close #1
    
End Sub