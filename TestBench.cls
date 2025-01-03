Option Explicit

'The commented code below will be added to Utils module
'
'Private m_tester As TestBench
'
'Public Function Tester() As TestBench
'    If m_tester Is Nothing Then
'        Set m_tester = New TestBench
'    End If
'    Set Tester = m_tester
'End Function


#If VBA7 Then
    Private Declare PtrSafe Function CoCreateGuid Lib "ole32" (id As Any) As Long
#Else
    Private Declare Function CoCreateGuid Lib "ole32" (id As Any) As Long
#End If


Private m_counter           As Integer
Private m_errors            As Integer
Private m_coll              As Collection
Private m_startTimer        As Double
Private m_endTimer          As Double
Private Const CHAR_LENGTH   As Integer = 10

Private Sub ResetCounters()
    m_counter = 0
    m_errors = 0
End Sub

Private Function CreateGUID() As String
    Const S_OK As Long = 0
    Dim id(0 To 15) As Byte
    Dim Cnt As Long, GUID As String
    If CoCreateGuid(id(0)) = S_OK Then
        For Cnt = 0 To 15
            CreateGUID = CreateGUID & IIf(id(Cnt) < 16, "0", "") + Hex$(id(Cnt))
        Next Cnt
        CreateGUID = Left$(CreateGUID, 8) & "-" & _
                     Mid$(CreateGUID, 9, 4) & "-" & _
                     Mid$(CreateGUID, 13, 4) & "-" & _
                     Mid$(CreateGUID, 17, 4) & "-" & _
                     Right$(CreateGUID, 12)
    Else
        MsgBox "Error while creating GUID!"
    End If
End Function

Public Sub StartTests(title As String, description As String)
    Dim startTag As String, dashTag As String, titleTag As String, descriptionTag As String, idTag As String, timeStampTag As String
    Call ResetCounters
    Set m_coll = New Collection
    startTag = String(CHAR_LENGTH, "=") + " TEST STARTS (" + title + ")" + String(CHAR_LENGTH, "=")
    dashTag = String(CHAR_LENGTH * 2 + 15 + Len(title), "-")
    titleTag = "Title: " + title
    descriptionTag = "Description: " + description
    idTag = "ID: " + CreateGUID
    timeStampTag = "Timestamp: " + Format(Now, "dd.mm.yyyy - hh:mm:ss")
    Call m_coll.Add(startTag)
    Call m_coll.Add(dashTag)
    Call m_coll.Add(titleTag)
    Call m_coll.Add(descriptionTag)
    Call m_coll.Add(idTag)
    Call m_coll.Add(timeStampTag)
    Call m_coll.Add(dashTag)
    Debug.Print startTag
    Debug.Print dashTag
    Debug.Print titleTag
    Debug.Print descriptionTag
    Debug.Print idTag
    Debug.Print timeStampTag
    Debug.Print dashTag
    m_startTimer = Timer
End Sub

Public Sub EndTests()
    Dim endTag As String, dashTag As String, summaryTag As String
    
    m_endTimer = Timer
    summaryTag = "Summary: " + CStr(CInt((m_counter - m_errors) / m_counter * 100)) + "% -> " + _
                CStr(m_counter - m_errors) + " passed, " + CStr(m_errors) + " failed in " + _
                Format(m_endTimer - m_startTimer, "0.0") + " seconds."
    dashTag = String(Len(summaryTag), "-")
    endTag = String((Len(summaryTag) - 14) / 2, "=") + " END OF TESTS " + String((Len(summaryTag) - 14) / 2, "=")
    Call m_coll.Add(dashTag)
    Call m_coll.Add(summaryTag)
    Call m_coll.Add(dashTag)
    Call m_coll.Add(endTag)
    Debug.Print dashTag
    Debug.Print summaryTag
    Debug.Print dashTag
    Debug.Print endTag
    Call ResetCounters
End Sub

Public Sub Assert(actualValue As Variant, expectedValue As Variant)
    Dim message As String
    m_counter = m_counter + 1
    If actualValue = expectedValue Then
        message = CStr(m_counter) + ") OK."
    Else
        message = CStr(m_counter) + ") Expected " + CStr(expectedValue) + " but got " + CStr(actualValue)
        m_errors = m_errors + 1
    End If
    Call m_coll.Add(message)
    Debug.Print message
End Sub

Public Sub Raise(macroName As String, errorCode As Integer)
    Dim message As String
    m_counter = m_counter + 1
    Call Application.Run(macroName)
    If Err.Number = errorCode Then
        message = CStr(m_counter) + ") OK."
    Else
        message = CStr(m_counter) + ") Expected error code is " + CStr(errorCode) + " but got " + CStr(Err.Number)
        m_errors = m_errors + 1
    End If
    Call m_coll.Add(message)
    Debug.Print message
End Sub

Public Sub SaveResults(fileName As String)
    Dim FSO As Object
    Dim textFile As Object
    Dim item As Variant

    Set FSO = CreateObject("Scripting.FileSystemObject")
    On Error GoTo ErrorHandler
    Set textFile = FSO.CreateTextFile(fileName, True)
    For Each item In m_coll
        textFile.WriteLine item
    Next item
    textFile.Close
    Exit Sub
ErrorHandler:
    MsgBox "Error writing to file: " & Err.description, vbCritical
    If Not textFile Is Nothing Then textFile.Close
End Sub
