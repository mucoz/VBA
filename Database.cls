Option Explicit

Private conn As ADODB.Connection
Private m_Arr As Variant
Private m_headers As Collection
Private m_dataTypes As Collection

Private Sub Class_Terminate()
    Call CloseConnection
    Set conn = Nothing
End Sub
Public Function ExecuteQuery(ByVal queryString As String) As Boolean
    On Error GoTo Catch
    If OpenConnection Then
        Call conn.Execute(queryString)
    End If
    On Error GoTo 0
    ExecuteQuery = True
    Exit Function
Catch:
    Call Utils.ThrowError("Error while executing single query (Database.ExecuteQuery). Connection State : " & CStr(conn.State) & ". Query string : " & queryString & ". Err description : " & Err.Description)
    On Error GoTo 0
    ExecuteQuery = False
    Call CloseConnection
End Function
Public Function ExecuteMultipleQueries(ByVal queryCollection As Collection) As Boolean
    Dim tmpQuery As Variant
    On Error GoTo Catch
    If OpenConnection Then
        Call conn.BeginTrans
        For Each tmpQuery In queryCollection
            Call conn.Execute(CStr(tmpQuery))
            DoEvents
        Next tmpQuery
    End If
    Call conn.CommitTrans
    Call CloseConnection
    On Error GoTo 0
    ExecuteMultipleQueries = True
    Exit Function
Catch:
    If tmpQuery = vbNullString Then
        Call Utils.ThrowError("Error while executing multiple queries. Rollback will be executed (Database.ExecuteMultipleQueries). Connection State : " & CStr(conn.State) & ". No query available.")
    Else
        Call Utils.ThrowError("Error while executing multiple queries. Rollback will be executed (Database.ExecuteMultipleQueries). Connection State : " & CStr(conn.State) & ". Err description : " & Err.Description)
    End If
    On Error GoTo 0
    ExecuteMultipleQueries = False
    Call conn.RollbackTrans
    Call CloseConnection
End Function
Public Function GetHeaders() As Collection
    Set GetHeaders = m_headers
End Function
Public Function GetDataTypes() As Collection
    Set GetDataTypes = m_dataTypes
End Function

Public Function GetRecords(ByVal queryString As String) As Variant '
    Dim rs As Recordset, arr As Variant, i As Integer 'queryString As String,
    Set rs = New ADODB.Recordset
    On Error GoTo Catch
    Set m_headers = New Collection
    Set m_dataTypes = New Collection
    If OpenConnection Then
        With rs
            .Source = queryString
            .ActiveConnection = conn
            .CursorLocation = adUseClient
            .LockType = adLockOptimistic
            .CursorType = adOpenForwardOnly
            .Open
        End With
        arr = rs.GetRows()
        For i = 0 To rs.Fields.Count - 1
            Call m_headers.Add(rs.Fields(i).Name, CStr(rs.Fields(i).Name))
            Call m_dataTypes.Add(rs.Fields(i).Type, CStr(rs.Fields(i).Name))
        Next i
        rs.Close
    End If
    GetRecords = arr
    Set rs = Nothing
    Call CloseConnection
    Call Utils.Log.Info("Collected data from database successfully.")
    Exit Function
Catch:
    If Err.Number = -2147467259 Then
        Call Utils.WarnUser("Table is empty. Query string : " & queryString)
    ElseIf Err <> 0 Then
        Call Utils.ThrowError("Error while getting records (Database.GetRecords). Query string : " & queryString & ". Err description : " & Err.Description)
    End If
    On Error GoTo 0
    Set rs = Nothing
    Call CloseConnection
End Function

Private Function OpenConnection() As Boolean
    On Error GoTo Catch
    If conn Is Nothing Then
        Set conn = New ADODB.Connection
        conn.ConnectionString = Replace(Constants.DB_CONNECTION_STRING, "<DB_PATH>", Constants.DB_PATH)
    End If
    If conn.State = 0 Then
        Call conn.Open
        Call Utils.Log.Info("DB Connection is open.")
    End If
    OpenConnection = True
    On Error GoTo 0
    Exit Function
Catch:
    On Error GoTo 0
    OpenConnection = False
    Call Utils.ThrowError("Cannot connect to database. Error description : " & Err.Description)
End Function
Private Sub CloseConnection()
    If Not conn Is Nothing Then
        If conn.State = 1 Then
            conn.Close
            Call Utils.Log.Info("DB Connection is closed.")
        End If
    End If
End Sub
