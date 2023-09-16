Option Explicit

' Hash a password before storing it
Public Function HashPassword(plainPassword As String) As String
    Dim sha256 As Object, hashedPassword As String, bytes() As Byte, hashedBytes() As Byte
    
    Set sha256 = CreateObject("System.Security.Cryptography.SHA256Managed")
    bytes = StrConv(plainPassword, vbFromUnicode)
    hashedBytes = sha256.ComputeHash_2((bytes))
    hashedPassword = StrConv(hashedBytes, vbUnicode)
    Set sha256 = Nothing
    HashPassword = hashedPassword
End Function

' Verify a password when a user logs in
Public Function VerifyPassword(plainPassword As String, hashedPassword As String) As Boolean
    Dim hashedAttempt As String
    
    hashedAttempt = HashPassword(plainPassword)
    VerifyPassword = (hashedAttempt = hashedPassword)
End Function
