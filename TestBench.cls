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


Private m_counter As Integer
Private m_errors As Integer

Private Sub ResetCounters()
    m_counter = 0
    m_errors = 0
End Sub

Public Sub StartTests(message As String)
    Call ResetCounters
    Debug.Print String(10, "=") + " TEST STARTS (" + message + ")" + String(10, "=")
End Sub

Public Sub EndTests()
    Debug.Print String(10, "=") + " END OF TESTS (" + CStr(m_counter - m_errors) + " passed, " + CStr(m_errors) + " failed)" + String(10, "=")
    Call ResetCounters
End Sub

Public Sub Assert(actualValue As Variant, expectedValue As Variant)
    m_counter = m_counter + 1
    If actualValue = expectedValue Then
        Debug.Print CStr(m_counter) + ") OK."
    Else
        Debug.Print CStr(m_counter) + ") Expected " + CStr(expectedValue) + " but got " + CStr(actualValue)
        m_errors = m_errors + 1
    End If
End Sub

Public Sub Raise(macroName As String, errorCode As Integer)
    m_counter = m_counter + 1
    Call Application.Run(macroName)
    If Err.Number = errorCode Then
        Debug.Print CStr(m_counter) + ") OK."
    Else
        Debug.Print CStr(m_counter) + ") Expected error code is " + CStr(errorCode) + " but got " + CStr(Err.Number)
        m_errors = m_errors + 1
    End If
End Sub
