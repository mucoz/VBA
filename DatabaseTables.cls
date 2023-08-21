Option Explicit

Private m_db As Database
Private m_rs As ADODB.Recordset
Private m_tableNames As Collection
Private m_tableCount As Long

Public Property Set Parent(db As Database)
    Set m_db = db
End Property

Public Property Get Parent() As Database
    Set Parent = m_db
End Property

Public Property Get Count() As Long
    If ValuesMissing Then Exit Sub
    Count = m_tableCount
End Property

Public Property Get Choose(tableName As String) As DatabaseTable
    If Choose Is Nothing Then
        Set Choose = New DatabaseTable
        Choose.Name = tableName
        Set Choose.Parent = Me
        Call Choose.RefreshMetadata
    End If
End Property

Public Sub Add(tableName As String) 'TODO : Optional dictionary will be added for fields and data types
    If ValuesMissing Then Exit Sub
    Dim sql As String
    sql = "CREATE TABLE " & tableName '& " (ID AUTOINCREMENT PRIMARY KEY, Name TEXT)"
    If m_db.ExecuteQuery(sql) Then
        Debug.Print "Table '" & tableName & "' added."
    Else
        Debug.Print "Failed to add table '" & tableName & "'."
    End If
End Sub

Public Sub Remove(tableName As String)
    If ValuesMissing Then Exit Sub
    Dim sql As String
    sql = "DROP TABLE [" & tableName & "]"
    If m_db.ExecuteQuery(sql) Then
        Debug.Print "Table '" & tableName & "' removed."
    Else
        Debug.Print "Failed to remove table '" & tableName & "'."
    End If
End Sub

Public Function Exists(tableName) As Boolean
    If ValuesMissing Then Exit Function
    Dim i As Long
    Exists = False
    If m_tableNames.Count > 0 Then
        For i = 1 To m_tableNames.Count
            If m_tableNames(i) = tableName Then
                Exists = True
                Exit Function
            End If
        Next i
    End If
End Function

Public Sub PrintNames()
    If ValuesMissing Then Exit Sub
    Dim i As Long
    If m_tableNames.Count > 0 Then
        For i = 1 To m_tableNames.Count
            Debug.Print m_tableNames(i)
        Next i
    End If
End Sub

Public Sub RefreshMetadata()
    Set m_rs = m_db.ActiveConnection.OpenSchema(adSchemaTables)
    Do While Not m_rs.EOF
        If m_rs.Fields("TABLE_TYPE").Value = "TABLE" Then
            m_tableCount = m_tableCount + 1
            m_tableNames.Add m_rs.Fields("TABLE_NAME").Value
        End If
        m_rs.MoveNext
    Loop
    m_rs.Close
    Set m_rs = Nothing
End Sub

Private Sub Class_Initialize()
    m_tableCount = 0
    Set m_tableNames = New Collection
    Set m_rs = New ADODB.Recordset
End Sub

Private Function ValuesMissing() As Boolean
    ValuesMissing = False
    If m_db Is Nothing Then
        Call MsgBox("Database not defined!", vbCritical + vbOKOnly, "Error")
        ValuesMissing = True
        Exit Function
    End If
'    If m_tableName = "" Then
'        Call MsgBox("Table name not defined!", vbCritical + vbOKOnly, "Error")
'        ValuesMissing = True
'        Exit Function
'    End If
End Function
