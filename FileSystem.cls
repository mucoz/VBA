Option Explicit

Public Function Exists(path As String) As Boolean
    Exists = (Dir(path, vbNormal Or vbDirectory) <> "")
End Function

Public Function Basename(path As String) As String
    Basename = Mid(path, InStrRev(path, "\") + 1)
End Function

Public Function Dirname(path As String) As String
    Dim pos As Long
    pos = InStrRev(path, "\")
    If pos > 0 Then
        Dirname = Left(path, pos - 1)
    Else
        Dirname = ""
    End If
End Function

Public Sub Mkdir(path As String)
    If Not Exists(path) Then
        VBA.Mkdir path
    End If
End Sub

Public Function JoinPath(path1 As String, path2 As String) As String
    If Right(path1, 1) = "\" Then
        JoinPath = path1 & path2
    Else
        JoinPath = path1 & "\" & path2
    End If
End Function

Public Function SplitExt(path As String) As Variant
    Dim fname As String, pos As Long, nameOnly As String, ext As String
    fname = Basename(path)
    pos = InStrRev(fname, ".")
    If pos > 0 Then
        nameOnly = Left(fname, pos - 1)
        ext = Mid(fname, pos)
    Else
        nameOnly = fname
        ext = ""
    End If
    SplitExt = Array(nameOnly, ext)
End Function

Public Function GetExtension(path As String) As String
    Dim parts As Variant
    parts = SplitExt(path)
    GetExtension = parts(1)
End Function

Public Sub Remove(path As String)
    If Exists(path) Then
        Kill path
    End If
End Sub

Public Sub Rmdir(path As String)
    If Exists(path) Then
        VBA.Rmdir path
    End If
End Sub

Public Function GetFileSize(path As String) As Long
    On Error GoTo ErrorHandler
    If Not Exists(path) Then
        GetFileSize = -1
        Exit Function
    End If
    GetFileSize = FileLen(path)
    Exit Function
ErrorHandler:
    GetFileSize = -1
End Function
