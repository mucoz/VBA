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


Public Function isVPNConnected() As Boolean

Dim PingResults As Object
Dim PingResult As Variant
Dim Query As String
Dim Ping As Boolean

Dim Host As String


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                  In order to find Host Name, you can use "ipconfig /all" command                 '
'                 There might be more than one name in cmd. Try one by one to check                '
'      You can also open Control Panel -> Network and Sharing Center to see the Domain network     '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Host = "hitachivantara.com"

Query = "SELECT * FROM Win32_PingStatus WHERE Address = '" & Host & "'"

Set PingResults = GetObject("winmgmts://./root/cimv2").ExecQuery(Query)

For Each PingResult In PingResults
    If Not IsObject(PingResult) Then
        Ping = False
    ElseIf PingResult.StatusCode = 0 Then
        Ping = True
    Else
        Ping = False
    End If
Next

    isVPNConnected = Ping
    
End Function
