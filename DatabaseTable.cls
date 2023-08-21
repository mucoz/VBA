Option Explicit

Private m_tables As DatabaseTables
Private m_rs As ADODB.Recordset
Private m_cmd As ADODB.Command
Private m_fieldNames As Collection
Private m_fieldTypes As Collection
Private m_tableName As String
Private m_recordCount As Long

Public Property Set Parent(Tables As DatabaseTables)
    Set m_tables = Tables
End Property

Public Property Get Parent() As DatabaseTables
    Set Parent = m_tables
End Property

Public Property Let Name(tableName As String)
    m_tableName = tableName
End Property

Public Property Get Name() As String
    Name = m_tableName
End Property

Public Property Get RecordCount() As Long
    If ValuesMissing Then Exit Sub
    RecordCount = m_recordCount
End Property

Public Sub PrintFields()
    If ValuesMissing Then Exit Sub
'    Call RefreshMetadata
    Dim i As Long
    If m_fieldNames.Count > 0 Then
    For i = 1 To m_fieldNames.Count
        Debug.Print m_fieldNames(i) + " : " + CStr(m_fieldTypes(m_fieldNames(i)))
    Next i
    End If
End Sub
Private Function CollectionContainsCollections(ByVal coll As Collection) As Boolean
    Dim item As Variant
    For Each item In coll
        If TypeOf item Is Collection Then
            CollectionContainsCollections = True
            Exit Function
        End If
    Next item
    CollectionContainsCollections = False
End Function
Private Function AllFieldsComplete(ByRef FieldList As Variant) As Boolean
    Dim i As Integer
    If TypeName(FieldList) = "Variant()" Then
        For i = LBound(FieldList) To UBound(FieldList)
            If Not FieldExists(FieldList(i)) Then
                Call MsgBox("'" + FieldList(i) + "' table not found in database!", vbCritical + vbOKOnly, "Error")
                AllFieldsComplete = False
                Exit Function
            End If
        Next i
    ElseIf TypeName(FieldList) = "Collection" Then
        For i = 1 To FieldList.Count
            If Not FieldExists(FieldList(i)) Then
                Call MsgBox("'" + FieldList(i) + "' table not found in database!", vbCritical + vbOKOnly, "Error")
                AllFieldsComplete = False
                Exit Function
            End If
        Next i
    Else
        Call MsgBox("Unknown field list type!", vbCritical + vbOKOnly, "Error")
        AllFieldsComplete = False
        Exit Function
    End If
    AllFieldsComplete = True
End Function
Public Sub AddRecords(ByRef FieldList As Variant, ByRef ValueList As Variant)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'      Input   : Both parameters must have the same data types and the function accepts the following data types :                              '
'              : 'String' is used only for single column entry.                                                                                 '
'              : For example : AddRecords("EmployeeID", "123456")                                                                               '
'              : 'Array' is used for multiple columns entry.                                                                                    '
'              : For example : AddRecords(Array("EmployeeID", "EmployeeName", "EmployeeAge"), Array("123456", "Mustafa", "30"))                 '
'              : 'Collection' is used for multiple columns entry. ValueList collection1 can contain values or other collections                 '
'              : For Example : AddRecords(fieldColl, valueColl))                                                                                '
'      Output  : New record(s) are added. When it fails RollbackTrans method is executed.                                                       '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim i As Long, j As Long
    On Error GoTo Catch
    If TypeName(FieldList) <> TypeName(ValueList) Then
        Call MsgBox("FieldList and ValueList must be in the same data structure!", vbCritical + vbOKOnly, "Error")
        Exit Sub
    End If
    If ValuesMissing Then Exit Sub
    If m_rs Is Nothing Then Set m_rs = New ADODB.Recordset
    m_rs.Open m_tableName, ActiveConnection:=m_tables.Parent.ActiveConnection, CursorType:=adOpenForwardOnly, LockType:=adLockOptimistic
    m_tables.Parent.ActiveConnection.BeginTrans
    If TypeName(FieldList) = "String" Then
        If Not FieldExists(FieldList) Then
            Call MsgBox("'" + FieldList + "' table not found in the database!", vbCritical + vbOKOnly, "Error")
            Exit Sub
        End If
        m_rs.AddNew FieldList, ValueList
        m_rs.Update
    ElseIf TypeName(FieldList) = "Variant()" Then
        If UBound(FieldList) = UBound(ValueList) Then
            m_rs.AddNew
            For i = LBound(FieldList) To UBound(FieldList)
                m_rs.Fields(FieldList(i)) = ValueList(i)
            Next i
            m_rs.Update
        Else
            Call MsgBox("Field and Value Arrays do not have the same size!", vbCritical + vbOKOnly, "Error")
        End If
    ElseIf TypeName(FieldList) = "Collection" Then
        If CollectionContainsCollections(ValueList) = False Then
            m_rs.AddNew
            For i = 1 To ValueList.Count
                m_rs.Fields(FieldList(i)) = ValueList(i)
            m_rs.Update
            Next i
        Else
            If FieldList.Count = ValueList(1).Count Then
                For i = 1 To ValueList.Count
                    m_rs.AddNew
                    For j = 1 To FieldList.Count
                        m_rs.Fields(FieldList(j)) = ValueList(i)(j)
                    Next j
                Next i
                m_rs.Update
            Else
                Call MsgBox("Field and Value Collections do not have the same size!", vbCritical + vbOKOnly, "Error")
            End If
        End If
    Else
        Call MsgBox("Data types of field list and value list are not compatible", vbCritical + vbOKOnly, "Error")
    End If
    Debug.Print "Record(s) added successfully."
    m_tables.Parent.ActiveConnection.CommitTrans
    m_rs.Close
    Exit Sub
Catch:
    Call MsgBox(Err.Description, vbCritical + vbOKOnly, "Error")
    m_tables.Parent.ActiveConnection.RollbackTrans
    On Error GoTo 0
    Debug.Print "Adding records failed."
End Sub
Public Sub UpdateRecords(CriteriaField As String, CriteriaValue As String, ByRef FieldList As Variant, ByRef ValueList As Variant)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'      Input   : Both parameters must have the same data types and the function accepts the following data types :                              '
'              : 'String' is used only for single column entry.                                                                                 '
'              : For example : UpdateRecords("EmployeeID", "123456")                                                                            '
'              : 'Array' is used for multiple columns entry.                                                                                    '
'              : For example : UpdateRecords(Array("EmployeeID", "EmployeeName", "EmployeeAge"), Array("123456", "Mustafa", "30"))             '
'              : 'Collection' is used for multiple columns entry. ValueList collection can contain can contain values or other collections      '
'              : For Example : UpdateRecords(fieldColl, valueColl))                                                                             '
'              : Criteria (optional) is used to specify the condition for updating records.                                                   '
'      Output  : The record(s) are updated. When it fails RollbackTrans method is executed.                                                    '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim i As Long, j As Long
    On Error GoTo Catch
    If ValuesMissing Then Exit Sub
    
    If TypeName(FieldList) <> TypeName(ValueList) Then
        Call MsgBox("FieldList and ValueList must be in the same data structure!", vbCritical + vbOKOnly, "Error")
        Exit Sub
    End If
    
    If Not FieldExists(CriteriaField) Then
        Call MsgBox("'" + CriteriaField + "' table not found in the database!", vbCritical + vbOKOnly, "Error")
        Exit Sub
    End If
    
    If TypeName(FieldList) = "String" Then
        Set m_cmd = New ADODB.Command
        With m_cmd
            .ActiveConnection = m_tables.Parent.ActiveConnection
            .CommandText = "UPDATE [" & m_tableName & "] SET [" + FieldList + "] = ? WHERE [" & CriteriaField & "] = ?;"
            .CommandType = adCmdText
            .Parameters.Append .CreateParameter(, adVariant, adParamInput, , ValueList) 'm_fieldTypes(FieldList)
            .Parameters.Append .CreateParameter(, m_fieldTypes(CriteriaField), adParamInput, , CriteriaValue)
            .Execute , , adCmdText + adExecuteNoRecords
        End With
    ElseIf TypeName(FieldList) = "Variant()" Then
        If UBound(FieldList) = UBound(ValueList) Then
            m_rs.MoveFirst
            Do Until m_rs.EOF
                m_rs.Update
                For i = LBound(FieldList) To UBound(FieldList)
                    m_rs.Fields(FieldList(i)).Value = ValueList(i)
                Next i
                m_rs.Update
                m_rs.MoveNext
            Loop
        Else
            Call MsgBox("Field and Value Arrays do not have the same size!", vbCritical + vbOKOnly, "Error")
        End If
    ElseIf TypeName(FieldList) = "Collection" Then
        If CollectionContainsCollections(ValueList) = False Then
            m_rs.MoveFirst
            Do Until m_rs.EOF
                m_rs.UpdateBatch adAffectGroup
                For i = 1 To ValueList.Count
                    m_rs.Fields(FieldList(i)).Value = ValueList(i)
                Next i
                m_rs.Update
                m_rs.MoveNext
            Loop
        Else
            If FieldList.Count = ValueList(1).Count Then
                For i = 1 To ValueList.Count
                    m_rs.MoveFirst
                    Do Until m_rs.EOF
                        m_rs.UpdateBatch adAffectGroup
                        For j = 1 To FieldList.Count
                            m_rs.Fields(FieldList(j)).Value = ValueList(i)(j)
                        Next j
                        m_rs.Update
                        m_rs.MoveNext
                    Loop
                Next i
            Else
                Call MsgBox("Field and Value Collections do not have the same size!", vbCritical + vbOKOnly, "Error")
            End If
        End If
    Else
        Call MsgBox("Data types of field list and value list are not compatible", vbCritical + vbOKOnly, "Error")
    End If
    Debug.Print "Record(s) updated successfully."
    Exit Sub
Catch:
    Call MsgBox(Err.Description, vbCritical + vbOKOnly, "Error")
    On Error GoTo 0
    Debug.Print "Updating records failed."
End Sub

Public Function FieldExists(ByVal FieldName As String) As Boolean
    If ValuesMissing Then Exit Function
    Dim i As Integer
    FieldExists = False
    If m_fieldNames.Count > 0 Then
        For i = 1 To m_fieldNames.Count
            If m_fieldNames(i) = FieldName Then
                FieldExists = True
                Exit Function
            End If
        Next i
    End If
End Function

Public Sub DeleteRecords(ByVal FieldName As String, ByVal Value As String)
    If ValuesMissing Then Exit Sub
    On Error GoTo Catch
    If Not FieldExists(FieldName) Then
        Call MsgBox("'" + FieldName + "' table not found in the database!", vbCritical + vbOKOnly, "Error")
        Exit Sub
    End If
    Set m_cmd = New ADODB.Command
    If TypeName(FieldName) = "String" Then
    With m_cmd
        .ActiveConnection = m_tables.Parent.ActiveConnection
        .CommandText = "DELETE FROM [" & m_tableName & "] WHERE [" & FieldName & "] = ?;"
        .CommandType = adCmdText
        .Parameters.Append .CreateParameter(, m_fieldTypes(FieldName), adParamInput, , Value)
        .Execute , , adCmdText + adExecuteNoRecords
    End With
    End If
    Debug.Print "Record(s) deleted successfully."
    Exit Sub
Catch:
    Call MsgBox(Err.Description, vbCritical + vbOKOnly, "Error")
    On Error GoTo 0
    Debug.Print "Deleting records failed."
End Sub

Public Sub GetRecords()
    Debug.Print "not yet implemented"
End Sub

Public Sub RefreshMetadata()
    Set m_fieldNames = New Collection
    Set m_fieldTypes = New Collection
    m_rs.Open m_tableName, ActiveConnection:=m_tables.Parent.ActiveConnection, CursorType:=adOpenForwardOnly, LockType:=adLockOptimistic
    m_recordCount = m_rs.RecordCount
    Set m_rs = m_tables.Parent.ActiveConnection.OpenSchema(adSchemaColumns)
    Do While Not m_rs.EOF
        If m_rs.Fields("TABLE_NAME").Value = m_tableName Then
            Call m_fieldNames.Add(m_rs.Fields("COLUMN_NAME").Value)
            Call m_fieldTypes.Add(m_rs.Fields("DATA_TYPE").Value, m_rs.Fields("COLUMN_NAME").Value)
        End If
        m_rs.MoveNext
    Loop
    
    m_rs.Close
    Set m_rs = Nothing
End Sub

Private Sub Class_Initialize()
    If m_rs Is Nothing Then Set m_rs = New ADODB.Recordset
    m_rs.CursorLocation = adUseClient
End Sub

Private Function ValuesMissing() As Boolean
    ValuesMissing = False
    If m_tables Is Nothing Then
        Call MsgBox("Database not defined!", vbCritical + vbOKOnly, "Error")
        ValuesMissing = True
        Exit Function
    End If
    If m_tableName = "" Then
        Call MsgBox("Table name not defined!", vbCritical + vbOKOnly, "Error")
        ValuesMissing = True
        Exit Function
    End If
End Function
