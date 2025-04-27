Option Explicit

Public UserName As String
Public UserEmail As String
Public UserPassword As String

Public Function GetAll(db As Database) As Variant
    Dim arr As Variant
    arr = db.FetchArray("SELECT * FROM Users")
    GetAll = arr
End Function

Public Sub Create(db As Database)
    Dim record As Object
    Set record = CreateObject("Scripting.Dictionary")
    
    record.Add "TABLE_NAME", "Users"
    record.Add "USER_NAME", UserName
    record.Add "USER_EMAIL", UserEmail
    record.Add "USER_PASSWORD", UserPassword
    Call db.Insert(record)
End Sub

Public Sub Edit(db As Database, whereClause As String)
    Dim record As Object
    Set record = CreateObject("Scripting.Dictionary")
    
    record.Add "TABLE_NAME", "Users"
    record.Add "USER_NAME", UserName
    record.Add "USER_EMAIL", UserEmail
    record.Add "USER_PASSWORD", UserPassword
    Call db.Update(record, whereClause)
End Sub

Public Sub Delete(db As Database, whereClause As String)
    Call db.Delete("Users", whereClause)
End Sub

Public Function GetUser() As ModelUser
    Set GetUser = Me
End Function


