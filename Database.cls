Option Explicit

Private m_conn As ADODB.Connection
Private m_dbPath As String

Private Sub Class_Terminate()
    Call CloseConnection
    Set m_conn = Nothing
End Sub
Public Function ActiveConnection() As ADODB.Connection
    Set ActiveConnection = m_conn
End Function

Public Function Tables() As DatabaseTables
    If Tables Is Nothing Then
        Set Tables = New DatabaseTables
        Set Tables.Parent = Me
        Call Tables.RefreshMetadata
    End If
End Function

Public Sub ConnectTo(DBPath As String)
    m_dbPath = DBPath
    Call OpenConnection
End Sub

Public Function ExecuteQuery(ByVal queryString As String) As Boolean
    On Error GoTo Catch
    If OpenConnection Then
        m_conn.BeginTrans
        Call m_conn.Execute(queryString)
        m_conn.CommitTrans
    End If
    On Error GoTo 0
    ExecuteQuery = True
    Exit Function
Catch:
    Call Utils.ThrowError("Error while executing single query (Database.ExecuteQuery). Connection State : " & CStr(m_conn.State) & ". Query string : " & queryString & ". Err description : " & Err.Description)
    On Error GoTo 0
    ExecuteQuery = False
    m_conn.RollbackTrans
    Call CloseConnection
End Function

Public Function ExecuteMultipleQueries(ByVal queryCollection As Collection) As Boolean
    Dim tmpQuery As Variant
    On Error GoTo Catch
    If OpenConnection Then
        Call m_conn.BeginTrans
        For Each tmpQuery In queryCollection
            Call m_conn.Execute(CStr(tmpQuery))
            DoEvents
        Next tmpQuery
        Call m_conn.CommitTrans
        Call CloseConnection
        ExecuteMultipleQueries = True
    End If
    Exit Function
Catch:
    If tmpQuery = vbNullString Then
        Call Utils.ThrowError("Error while executing multiple queries. Rollback will be executed (Database.ExecuteMultipleQueries). Connection State : " & CStr(m_conn.State) & ". No query available.")
    Else
        Call Utils.ThrowError("Error while executing multiple queries. Rollback will be executed (Database.ExecuteMultipleQueries). Connection State : " & CStr(m_conn.State) & ". Err description : " & Err.Description)
    End If
    On Error GoTo 0
    ExecuteMultipleQueries = False
    Call m_conn.RollbackTrans
    Call CloseConnection
End Function

Private Sub OpenConnection()
    On Error GoTo Catch
    If m_conn Is Nothing Then
        Set m_conn = New ADODB.Connection
        m_conn.Provider = "Microsoft.ACE.OLEDB.12.0"
        m_conn.Open m_dbPath
    End If
    Exit Sub
Catch:
    On Error GoTo 0
    Call MsgBox("Cannot connect to database!", vbCritical + vbOKOnly, "Error")
End Sub

Private Sub CloseConnection()
    If Not m_conn Is Nothing Then
        If m_conn.State = 1 Then
            m_conn.Close
            Call Utils.Log.Info("DB Connection is closed.")
        End If
    End If
End Sub

Public Property Get Name() As String
    Name = FSO.GetFileName(m_dbPath)
End Property

Public Property Get Path() As String
    Path = FSO.GetParentFolderName(m_dbPath)
End Property

Public Property Get FullName() As String
    FullName = FSO.GetAbsolutePathName(m_dbPath)
End Property

Private Function FSO() As FileSystemObject
    Set FSO = New FileSystemObject
End Function
