Option Explicit

' Database Class for Access DB
' Full mini ORM

Private m_conn As Object
Private m_dbPath As String

' Class Termination
Private Sub Class_Terminate()
    Call CloseConnection
    Set m_conn = Nothing
End Sub

' Connect to Database
Public Sub ConnectTo(DBPath As String)
    m_dbPath = DBPath
    Call OpenConnection
End Sub

' Open connection
Private Function OpenConnection() As Boolean
    On Error GoTo Catch
    If m_conn Is Nothing Then
        Set m_conn = CreateObject("ADODB.Connection")
        m_conn.Provider = "Microsoft.ACE.OLEDB.12.0"
        m_conn.Open m_dbPath
    End If
    OpenConnection = True
    Exit Function
Catch:
    OpenConnection = False
    On Error GoTo 0
End Function

' Close connection
Private Sub CloseConnection()
    If Not m_conn Is Nothing Then
        If m_conn.State = 1 Then
            m_conn.Close
        End If
    End If
End Sub

' Execute SQL without result (INSERT, UPDATE, DELETE)
Public Function ExecuteQuery(ByVal queryString As String) As Boolean
    On Error GoTo Catch
    If OpenConnection() Then
        m_conn.BeginTrans
        Call m_conn.Execute(queryString)
        m_conn.CommitTrans
        ExecuteQuery = True
    End If
    Exit Function
Catch:
    On Error GoTo 0
    ExecuteQuery = False
    m_conn.RollbackTrans
    Call CloseConnection
End Function

' Insert Record
Public Function Insert(ByVal data As Object) As Boolean
    Dim fields As String, values As String
    Dim key As Variant, value As Variant
    Dim sql As String
    On Error GoTo Catch

    If OpenConnection() Then
        fields = ""
        values = ""
        For Each key In data.Keys
            If key = "TABLE_NAME" Then
                GoTo NextKey
            End If
            fields = fields & "[" & key & "], "
            value = data(key)
            
            Select Case VarType(value)
                Case vbString
                    values = values & "'" & Replace(value, "'", "''") & "', "
                Case vbDate
                    values = values & "#" & Format(value, "yyyy-mm-dd hh:nn:ss") & "#, "
                Case vbNull
                    values = values & "Null, "
                Case Else
                    values = values & CStr(value) & ", "
            End Select
NextKey:
        Next key
        
        fields = Left(fields, Len(fields) - 2)
        values = Left(values, Len(values) - 2)
        
        sql = "INSERT INTO [" & data("TABLE_NAME") & "] (" & fields & ") VALUES (" & values & ");"
        Insert = ExecuteQuery(sql)
    End If
    Exit Function
Catch:
    On Error GoTo 0
    Insert = False
    Call CloseConnection
End Function

' Fetch multiple records as Array
Public Function FetchArray(ByVal queryString As String) As Variant
    Dim rs As Object
    On Error GoTo Catch
    If OpenConnection() Then
        Set rs = CreateObject("ADODB.Recordset")
        rs.Open queryString, m_conn, 1, 3 ' adOpenKeyset, adLockOptimistic
        If Not (rs.EOF And rs.BOF) Then
            FetchArray = rs.GetRows()
        Else
            FetchArray = Empty
        End If
        rs.Close
    End If
    Set rs = Nothing
    Exit Function
Catch:
    On Error GoTo 0
    FetchArray = Empty
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
End Function

' Fetch single row
Public Function FetchSingleRow(ByVal queryString As String) As Variant
    Dim data As Variant
    data = FetchArray(queryString)
    If IsEmpty(data) Then
        FetchSingleRow = Empty
    Else
        Dim i As Long
        Dim arr() As Variant
        ReDim arr(UBound(data, 1))
        For i = 0 To UBound(data, 1)
            arr(i) = data(i, 0)
        Next i
        FetchSingleRow = arr
    End If
End Function

' Fetch single value
Public Function FetchSingleValue(ByVal queryString As String) As Variant
    Dim data As Variant
    data = FetchArray(queryString)
    If IsEmpty(data) Then
        FetchSingleValue = Null
    Else
        FetchSingleValue = data(0, 0)
    End If
End Function

' Update Record
Public Function Update(ByVal data As Object, ByVal whereClause As String) As Boolean
    Dim setClause As String
    Dim key As Variant, value As Variant
    Dim sql As String
    On Error GoTo Catch

    If OpenConnection() Then
        setClause = ""
        For Each key In data.Keys
            If key = "TABLE_NAME" Then
                GoTo NextKey
            End If
            value = data(key)
            Select Case VarType(value)
                Case vbString
                    setClause = setClause & "[" & key & "] = '" & Replace(value, "'", "''") & "', "
                Case vbDate
                    setClause = setClause & "[" & key & "] = #" & Format(value, "yyyy-mm-dd hh:nn:ss") & "#, "
                Case vbNull
                    setClause = setClause & "[" & key & "] = Null, "
                Case Else
                    setClause = setClause & "[" & key & "] = " & CStr(value) & ", "
            End Select
NextKey:
        Next key
        
        setClause = Left(setClause, Len(setClause) - 2)
        sql = "UPDATE [" & data("TABLE_NAME") & "] SET " & setClause & " WHERE " & whereClause & ";"
        Update = ExecuteQuery(sql)
    End If
    Exit Function
Catch:
    On Error GoTo 0
    Update = False
    Call CloseConnection
End Function

' Delete Record
Public Function Delete(ByVal TableName As String, ByVal whereClause As String) As Boolean
    Dim sql As String
    On Error GoTo Catch

    If OpenConnection() Then
        sql = "DELETE FROM [" & TableName & "] WHERE " & whereClause & ";"
        Delete = ExecuteQuery(sql)
    End If
    Exit Function
Catch:
    On Error GoTo 0
    Delete = False
    Call CloseConnection
End Function

