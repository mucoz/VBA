Option Explicit

Private Declare Function InternetCheckConnection Lib "wininet.dll" Alias "InternetCheckConnectionA" (ByVal lpszUrl As String, ByVal dwFlags As Long, ByVal dwReserved As Long) As Long

Const FLAG_ICC_FORCE_CONNECTION = &H1

Public Function IsConnected() As Boolean

    Dim msg As Boolean
    Dim status As Boolean
    Dim address As String
    
    address = "http://www.yahoo.com/"
    
    status = InternetCheckConnection(address, FLAG_ICC_FORCE_CONNECTION, 0&)
    
    If status Then
        
        msg = True
    
    Else
        
        msg = False
        
    End If
    
    IsConnected = msg
    
End Function

