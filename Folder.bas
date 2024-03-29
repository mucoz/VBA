Option Explicit

'Created by Mustafa Can Ozturk on 08.02.2021
'===========================================
'The functions below will be used for folder operations
'It includes the most common functions

'Functions:
'Copy, Rename, Exist, Create, Delete, ParentFolder, SubFolders, SubFilesNames, SubFilesPaths, GetAllSubFolders, CreateZipFile, UnzipFile, IsThereFile
'===========================================

'These fiels will be used for "SubFolders" function
Private arr() As String
Private Counter As Long

'=======Auxiliary Functions=======
Private Function FSO() As Object
    
    'Set FSO = New FileSystemObject
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
End Function
'=======End of Auxiliary Functions=======

Public Sub Rename(FolderPath As String, NewName As String)
    
    Dim baseFolder As String
    
    If Exist(FolderPath) = True Then
    
        baseFolder = ParentFolder(FolderPath)
        Name FolderPath As baseFolder + NewName
    
    End If
    
End Sub

Public Sub CopyFolder(Source As String, Destination As String)

    FSO.CopyFolder Source, Destination
    
End Sub

Public Property Get Name(FolderPath) As String

    Name = FSO.GetFolder(FolderPath).Name
    
End Property

Public Function SubFilesNames(FolderPath As String) As Variant
    
    Dim subFilesArr() As Variant
    Dim files As Object
    Dim file As Object
    Dim i As Long
    
    If Exist(FolderPath) = False Then
    
        SubFilesNames = "PATH NOT FOUND"
        Exit Function
    
    End If
    
    Set files = FSO.GetFolder(FolderPath).files
    
    If files.Count = 0 Then
    
        SubFilesNames = "FILE NOT FOUND"
        Exit Function

    End If
    
    i = 0
    
    For Each file In files
        
        ReDim Preserve subFilesArr(i)
        subFilesArr(i) = file.Name
        i = i + 1
        
    Next file
    
    SubFilesNames = subFilesArr
    
End Function

Public Function SubFilesPaths(FolderPath As String) As Variant

    Dim subFilesArr() As Variant
    Dim files As Object
    Dim file As Object
    Dim i As Long
    
    If Exist(FolderPath) = False Then
    
        SubFilesPaths = "PATH NOT FOUND"
        Exit Function
    
    End If
    
    Set files = FSO.GetFolder(FolderPath).files
    
    If files.Count = 0 Then
    
        SubFilesPaths = "FILE NOT FOUND"
        Exit Function

    End If
    
    i = 0
    
    For Each file In files
        
        ReDim Preserve subFilesArr(i)
        subFilesArr(i) = file.Path
        i = i + 1
        
    Next file
    
    SubFilesPaths = subFilesArr
    
End Function

Public Property Get ParentFolder(ByVal Path As String) As String
    
    ParentFolder = FSO.GetParentFolderName(Path) + DetectSlash(Path)
    
End Property

Private Function DetectSlash(Path As String) As String
    
    If Right(Path, 1) <> "/" Or Right(Path, 1) <> "\" Then
    
        If InStr(Path, "\") > 0 Then
    
            DetectSlash = "\"
        
        ElseIf InStr(Path, "/") > 0 Then
    
            DetectSlash = "/"
        
        End If
    
    End If
    
End Function

Public Function Exist(ByVal FolderPath As String) As Boolean

    If FSO.FolderExists(FolderPath) = True Then
        
        Exist = True
    
    Else
    
        Exist = False
        
    End If
    
End Function

Public Sub Move(ByVal SourcePath As String, DestinationPath As String)

    If Exist(SourcePath) = False Then
    
        Debug.Print "Folder.Move error : source path doesn't exist"
        Exit Sub
        
    ElseIf Exist(DestinationPath) = False Then
    
        Debug.Print "Folder.Move error : destination path doesn't exist"
        Exit Sub
        
    End If
    
    FSO.MoveFolder SourcePath, DestinationPath + DetectSlash(DestinationPath)

End Sub

Public Sub Delete(Path As String)
    
    If Exist(Path) = True Then
    
        FSO.DeleteFolder Path
        
    End If
    
End Sub

Public Sub Create(Path As String)

    If Exist(Path) = False Then
    
        FSO.CreateFolder Path
    
    End If

End Sub

Public Sub CreateZipFile(SourceFolderPath As Variant, DestinationFullName As Variant)

    Dim ShellApp As Object

    'Create an empty zip file
    Open DestinationFullName For Output As #1
    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Close #1
    
    'Copy the files & folders into the zip file
    Set ShellApp = CreateObject("Shell.Application")
    ShellApp.Namespace(DestinationFullName).CopyHere ShellApp.Namespace(SourceFolderPath).Items
    
    'Zipping the files may take a while, create loop to pause the macro until zipping has finished.
    On Error Resume Next
    Do Until ShellApp.Namespace(DestinationFullName).Items.Count = ShellApp.Namespace(SourceFolderPath).Items.Count
        Application.Wait (Now + TimeValue("0:00:01"))
    Loop
    On Error GoTo 0

End Sub

Public Sub UnzipFile(SourceZipPath As Variant, UnzipToPath As Variant)
    
    Dim ShellApp As Object
    
    If Exist(CStr(UnzipToPath)) = False Then
        
        Create CStr(UnzipToPath)
        
        Do While Exist(UnzipToPath) = False
            DoEvents
            Application.Wait (Now + TimeValue("0:00:01"))
        Loop
        
    End If
    
    'Copy the files & folders from the zip into a folder
    Set ShellApp = CreateObject("Shell.Application")

    ShellApp.Namespace(UnzipToPath).CopyHere ShellApp.Namespace(SourceZipPath).Items
    
    
    Set ShellApp = Nothing
    
End Sub

Public Function GetAllSubFolders(MainPath As String) As Variant
    
    Erase arr
    
    Counter = 0
    
    Dim myArr
    
    myArr = GetSubFolders(MainPath)
    
    GetAllSubFolders = myArr

End Function

Private Function GetSubFolders(RootPath As String) As Variant
    
    Dim fld As Object
    Dim sf As Object
    Dim myArr
    
    If Exist(RootPath) = False Then
        
        GetSubFolders = "INVALID PATH"
        Exit Function
        
    End If
    
    Set fld = FSO.GetFolder(RootPath)
    
    For Each sf In fld.SubFolders
        
        ReDim Preserve arr(Counter)
        
        arr(Counter) = sf.Path
        
        Counter = Counter + 1
        
        myArr = GetSubFolders(sf.Path)
        
        DoEvents
    
    Next
    
    GetSubFolders = arr
    
    Set sf = Nothing
    Set fld = Nothing
    
End Function

Public Function SubFolders(MainPath As String) As Variant
    
    Dim fld As Object
    Dim sf As Object
    
    Erase arr
    
    Counter = 0
    
    If Exist(MainPath) = False Then
        
        SubFolders = "INVALID PATH"
        Exit Function
        
    End If
    
    Set fld = FSO.GetFolder(MainPath)
    
    For Each sf In fld.SubFolders
        
        ReDim Preserve arr(Counter)
        
        arr(Counter) = sf.Path
        
        Counter = Counter + 1
        
        DoEvents
    
    Next
    
    SubFolders = arr
    
    Set sf = Nothing
    Set fld = Nothing
    
End Function

Public Function SubFoldersWithPattern(MainPath As String, pattern As String) As Variant
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                 ?        Any single character                                    '
'                                 *        Zero or more characters                                 '
'                                 #        Any single digit (0-9)                                  '
'                        [charlist]        Any single character in charlist                        '
'                       [!charlist]        Any single character not in charlist                    '
'                                                                                                  '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim fld As Object
    Dim sf As Object
    
    Erase Arr
    
    Counter = 0
    
     If Exist(MainPath) = False Then
        
        SubFoldersWithPattern = "INVALID PATH"
        Exit Function
        
    End If
    
    Set fld = FSO.GetFolder(MainPath)
    
    For Each sf In fld.SubFolders
        
        If IsLike(Folder.Name(sf.Path), pattern) = True Then
        ReDim Preserve Arr(Counter)
        
        Arr(Counter) = sf.Path
        
        Counter = Counter + 1
        
        DoEvents
        End If
    Next
    
    If Counter = 0 Then
    SubFoldersWithPattern = "No Folder in " & CStr(MainPath)
    Else
    SubFoldersWithPattern = Arr
    End If
    
    Set sf = Nothing
    Set fld = Nothing
    
End Function

Private Function IsLike(text As String, pattern As String) As Boolean

    IsLike = text Like pattern
    
End Function

Public Function FindFileName(FolderPath As String, SearchFileName As String) As String

    Dim result As String
    
    result = FindFilePath(FolderPath, SearchFileName)
    
    If result = "FILE NOT FOUND" Then
        FindFileName = result
    Else
        FindFileName = file.Name(result)
    End If

End Function

Public Function FindFilePath(FolderPath As String, SearchFileName As String) As String

    Dim files As Variant
    Dim search As Variant
    Dim i As Long, j As Long
    Dim m As Long
    
    files = SubFilesPaths(FolderPath)
    search = Split(SearchFileName, " ")
    
    For i = LBound(files) To UBound(files)
    
        m = LBound(search)
        
        For j = LBound(search) To UBound(search)
            
            If InStr(UCase(files(i)), UCase(search(j))) > 0 Then
                m = m + 1
            End If
            
        Next j
        
        If m = UBound(search) + 1 Then
                
                FindFilePath = files(j)
                Exit Function
                
        End If
        
    Next i
    
    FindFilePath = "FILE NOT FOUND"
    
End Function

Public Function IsThereFile(FolderPath As String, SearchFileName As String) As Boolean
    
    Dim filePath As String
    
    filePath = FindFilePath(FolderPath, SearchFileName)
    
    If filePath = "FILE NOT FOUND" Then
        
        IsThereFile = False
    
    Else
    
        IsThereFile = True
        
    End If
    
End Function
                                                                            
Public Function GetFolderPath()

Dim fd As FileDialog

Set fd = Application.FileDialog(msoFileDialogFolderPicker)

With fd
    .AllowMultiSelect = False
    .ButtonName = "Select"
    .Title = "Choose Folder"
    
    If .Show = -1 Then
        GetFolderPath = .SelectedItems(1)
    Else
        GetFolderPath = ""
    End If
    
End With

End Function

Public Function GetFolderName(FolderPath As String)
 
    If Len(FolderPath) = 0 Or InStr(FolderPath, "\") = 0 Then
        MsgBox "Choose a valid folder", vbInformation + vbOKOnly, "Warning"
        GetFolderName = ""
        Exit Function
    End If
    
    Dim i As Integer
    Dim result As String
    Dim chr As String * 1
    
    Do
        chr = Mid(FolderPath, Len(FolderPath) - i, 1)
        If chr <> "\" Then
            result = chr + result
        Else
            Exit Do
        End If
        i = i + 1
    
    Loop
    
    GetFolderName = result
 
End Function
    

