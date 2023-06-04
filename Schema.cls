Option Explicit

Private m_tableName As String
Private m_sortField As String
Private m_db As Database
Private m_Arr As Variant
Private m_headers As Collection
Private m_dataTypes As Collection
Private m_dt As Integer 'data type in access database

Public Property Let TableName(Value As String)
    m_tableName = Value
End Property
Public Property Get TableName() As String
    TableName = m_tableName
End Property
Public Property Let SortField(Value As String)
    m_sortField = Value
End Property
Public Property Get SortField() As String
    SortField = m_sortField
End Property
Public Property Get GetHeaders() As Collection
    Set GetHeaders = m_headers
End Property
Public Function Refresh(Optional queryStr As Variant)
    Set m_db = New Database
    If IsMissing(queryStr) Then
        m_Arr = m_db.GetRecords("SELECT * FROM " & m_tableName & " ORDER BY " & m_sortField)
    Else
        m_Arr = m_db.GetRecords(queryStr)
    End If
    Set m_headers = m_db.GetHeaders()
    Set m_dataTypes = m_db.GetDataTypes()
End Function
Public Function Count() As Long
    Dim result As Long
    result = 0
    If IsEmpty(m_Arr) = False Then
        result = UBound(m_Arr, 2)
    End If
    Count = result
End Function
Public Sub PrintFields(Fields As Variant)
    Dim i As Long, j As Long
    Dim message As String, nextField As Variant
    If IsEmpty(m_Arr) = True Then Exit Sub
    For i = LBound(m_Arr, 2) To UBound(m_Arr, 2)
        message = ""
        For j = LBound(Fields, 1) To UBound(Fields, 1)
            nextField = Fields(j)
            message = message + CStr(m_Arr(nextField, i)) + vbTab
        Next j
        Debug.Print message
    Next i
End Sub
Public Function FindAll() As Collection
    Dim itemsColl As New Collection
    Dim i As Long, j As Long
    If IsEmpty(m_Arr) = False Then
        For i = LBound(m_Arr, 2) To UBound(m_Arr, 2)
            For j = LBound(m_Arr, 1) To UBound(m_Arr, 1)
                Dim userColl As New Collection
                Call userColl.Add(m_Arr(j, i))
            Next j
            Call itemsColl.Add(userColl)
            Set userColl = Nothing
        Next i
    End If
    Set FindAll = itemsColl
End Function
Public Function FindBy(FieldIndex As Integer, Value As Variant) As Collection
    Dim itemsColl As New Collection
    Dim i As Long, j As Long
    If IsEmpty(m_Arr) = False Then
        For i = LBound(m_Arr, 2) To UBound(m_Arr, 2)
            If m_Arr(FieldIndex, i) = Value Then
                For j = LBound(m_Arr, 1) To UBound(m_Arr, 1)
                    Dim userColl As New Collection
                    Call userColl.Add(m_Arr(j, i))
                Next j
                Call itemsColl.Add(userColl)
                Set userColl = Nothing
            End If
        Next i
    End If
    Set FindBy = itemsColl
End Function

Public Function FindFirst(FieldIndex As Integer, Value As Variant) As Collection
    Dim itemsColl As New Collection
    Dim i As Long, j As Long
    If IsEmpty(m_Arr) = False Then
        For i = LBound(m_Arr, 2) To UBound(m_Arr, 2)
            If m_Arr(FieldIndex, i) = Value Then
                For j = LBound(m_Arr, 1) To UBound(m_Arr, 1)
                    Dim userColl As New Collection
                    Call userColl.Add(m_Arr(j, i))
                Next j
                Call itemsColl.Add(userColl)
                Set userColl = Nothing
                Set FindFirst = itemsColl
                Exit Function
            End If
        Next i
    End If
    Set FindFirst = itemsColl
End Function

Public Function Delete(FieldIndex As Integer, Value As Variant) As Boolean
    Dim itemsToDelete As New Collection
    Dim deleteQueries As New Collection
    Dim deleteQuery As String
    Dim singleColl  As Collection
    Dim answer As Integer
    Set itemsToDelete = FindBy(FieldIndex, Value)
    If itemsToDelete.Count = 0 Then
        Call MsgBox("Record does not exist!", vbExclamation + vbOKOnly, "Warning")
        Delete = False
        Exit Function
    End If
    If itemsToDelete.Count > 1 Then
        answer = MsgBox("More than one record found. Delete all?", vbYesNo + vbInformation, "Warning")
        If answer = vbNo Then Exit Function
    End If
    For Each singleColl In itemsToDelete
         deleteQuery = "DELETE FROM " & m_tableName & " WHERE " & m_headers(FieldIndex + 1) & "='" & Trim(singleColl(FieldIndex + 1)) & "';"
         Call deleteQueries.Add(deleteQuery)
    Next singleColl
    If m_db.ExecuteMultipleQueries(deleteQueries) = True Then
        Delete = True
    Else
        Delete = False
    End If
End Function

Public Function UpdateRecord(Fields As Collection, Record As Collection, FieldIndex As Integer, Value As Variant) As Boolean
    Dim updateQueries As New Collection
    Dim i As Long
    Dim updateQuery As String
    
    If Fields.Count <> Record.Count Then
        Call MsgBox("Number of fields and record values mismatch.", vbCritical + vbOKOnly, "Error")
        UpdateRecord = False
        Exit Function
    End If
    
    updateQuery = "UPDATE " & m_tableName & " SET "
    
    For i = 1 To Fields.Count
        m_dt = m_dataTypes(Fields(i))
        If m_dt = 3 Or m_dt = 5 Or m_dt = 14 Or m_dt = 11 Or m_dt = 20 Then
            updateQuery = updateQuery & "[" & Fields(i) & "]=" & Record(i) & ""
        Else
            updateQuery = updateQuery & "[" & Fields(i) & "]='" & Record(i) & "'"
        End If
        If i <> Fields.Count Then
            updateQuery = updateQuery & ","
        End If
    Next i

    m_dt = m_dataTypes(Fields(FieldIndex + 1))
    If m_dt = 3 Or m_dt = 5 Or m_dt = 14 Or m_dt = 11 Or m_dt = 20 Then
        updateQuery = updateQuery & " WHERE " & m_headers(FieldIndex + 1) & "=" & Value & ";"
    Else
        updateQuery = updateQuery & " WHERE " & m_headers(FieldIndex + 1) & "='" & Value & "';"
    End If
    Call updateQueries.Add(updateQuery)
    
    If m_db.ExecuteMultipleQueries(updateQueries) Then
        UpdateRecord = True
    Else
        UpdateRecord = False
    End If
End Function
Public Function AddRecord(Fields As Collection, Values As Collection) As Boolean
    Dim insertQuery As String
    Dim i As Long
    
    If Fields.Count <> Values.Count Then
        Call MsgBox("Number of fields and values mismatch.", vbCritical + vbOKOnly, "Error")
        AddRecord = False
        Exit Function
    End If
    insertQuery = "INSERT INTO " & m_tableName & " ("
    For i = 1 To Fields.Count
        insertQuery = insertQuery & "[" & Fields(i) & "]"
        If i <> Fields.Count Then
            insertQuery = insertQuery & ","
        End If
    Next i

    
    insertQuery = insertQuery & ") VALUES ("
    For i = 1 To Values.Count
        m_dt = m_dataTypes(Fields(i))
        If m_dt = 3 Or m_dt = 5 Or m_dt = 14 Or m_dt = 11 Or m_dt = 20 Then '3:Number; 20:LargeNumber; 11:Boolean
            insertQuery = insertQuery & Values(i)
        Else '202:shorttext; 203:longtext
            insertQuery = insertQuery & "'" & Values(i) & "'"
        End If
        If i <> Values.Count Then
            insertQuery = insertQuery & ","
        End If
    Next i
    insertQuery = insertQuery & ");"
    If m_db.ExecuteQuery(insertQuery) Then
        AddRecord = True
    Else
        AddRecord = False
    End If
End Function

Public Function AddRecords(Fields As Collection, Records As Collection) As Boolean
    Dim insertQueries As New Collection
    Dim i As Long, j As Long
    Dim insertQuery As String
    If Records.Count = 0 Then
        Call MsgBox("No records to add!", vbExclamation + vbOKOnly, "Warning")
        AddRecords = False
        Exit Function
    End If
    If Fields.Count <> Records(1).Count Then
        Call MsgBox("Number of fields and records mismatch.", vbExclamation + vbOKOnly, "Warning")
        AddRecords = False
        Exit Function
    End If
    
    For j = LBound(Records, 1) To UBound(Records, 1)
        insertQuery = "INSERT INTO " & m_tableName & " ("
        
        For i = 1 To Fields.Count
            insertQuery = insertQuery & "[" & Fields(i) & "]"
            If i <> Fields.Count Then
                insertQuery = insertQuery & ","
            End If
        Next i
        
        insertQuery = insertQuery & ") VALUES ("
        
        For i = 1 To Fields.Count
            insertQuery = insertQuery & "'" & Records(j)(i) & "'"
            If i <> Fields.Count Then
                insertQuery = insertQuery & ","
            End If
        Next i
        
        insertQuery = insertQuery & ");"
        
        Call insertQueries.Add(insertQuery)
    Next j
    
    If m_db.ExecuteMultipleQueries(insertQueries) Then
        AddRecords = True
    Else
        AddRecords = False
    End If
End Function

Private Sub Class_Terminate()
    Set m_db = Nothing
End Sub
