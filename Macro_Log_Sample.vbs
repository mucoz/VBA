'File Module===============================================================================================================================

Option Explicit

Enum LogStatus
    
    Start
    Finish
    Fail
    Blank
    
End Enum


Private ProcessDuration As Double
Private LastTime As Double

Private Function FSO() As Object
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
End Function

Public Sub CreateLog(FilePath As String)
    
    Dim f As Object
 
    Set f = FSO.CreateTextFile(FilePath, True)
    
    StartTimer
    
    Set f = Nothing
    
    File.LogTo FilePath, UCase("Process has been started by : ") & Globals.UserName, Blank
    
End Sub


Public Sub TerminateLog(FilePath As String)
    
    If File.Exist(FilePath) = True Then
        
        File.LogTo FilePath, UCase("Process has been completed in " & ProcessTime & " seconds."), Blank
    
    End If
    
End Sub

Public Sub LogTo(ByVal Path As String, Text As String, Optional Status As LogStatus) 'Optional ErrorMessage As Boolean = False)
            
    Dim message As String
    
    Do
    DoEvents
    Loop Until IsTXTOpen(Path) = False
    
    Open Path For Append As #1
        
        If Status = Start Then
            message = Text & " process has been started."
        ElseIf Status = Finish Then
            message = Text & " has been completed successfully" & " in " & CStr(Format(Timer - LastTime, "00.00")) & " seconds."
        ElseIf Status = Blank Or Fail Then
            message = Text
        End If
        
        If Status = Fail Then
            Print #1, String(Len(CStr(Now) & "   ->   " & Text), "=")
        End If
        
        Print #1, CStr(Now) & "   ->   " & message
        
        If Status = Fail Then
            Print #1, String(Len(CStr(Now) & "   ->   " & Text), "=")
        End If
        Print #1, vbNewLine
    
    Close #1
    
    LastTime = Timer
    
End Sub

Public Function IsTXTOpen(ByVal FileName As String) As Boolean
    Dim iFilenum As Long
    Dim iErr As Long
     
    On Error Resume Next
    iFilenum = FreeFile()
    Open FileName For Input Lock Read As #iFilenum
    Close iFilenum
    iErr = Err
    'On Error GoTo 0
     
    Select Case iErr
    Case 0:    IsTXTOpen = False
    Case 70:   IsTXTOpen = True
    Case Else: Error iErr
    End Select
     
End Function

Public Function Exist(FilePath As String) As Boolean

    If FSO.fileexists(FilePath) = True Then
        
        Exist = True
    
    Else
    
        Exist = False
        
    End If
    
End Function

Private Function ProcessTime() As String
    
    ProcessTime = CStr(Format(Timer - ProcessDuration, "00.00"))
    
End Function

Private Sub StartTimer()

    ProcessDuration = Timer
    LastTime = ProcessDuration
    
End Sub

'Globals Module===============================================================================================================================

Option Explicit

'This is the path of the Log File which will be used during runtime
Public Function LogFile() As String
    
    LogFile = Environ("userprofile") & "\Desktop\MacroLog_" & CStr(Format(Now, "ddmmyyyy")) & ".txt"

End Function

Public Function UserName() As String
    
    UserName = Environ("computername")
    
End Function

'Main Module===============================================================================================================================

Option Explicit

'Entry Point

Sub RunMacro()
    
    On Error GoTo ErrorHandler
    
    'Create Log File for the macro and keep the process time
    Call File.CreateLog(LogFile)
    
    
    Call Task1
    
    Call Task2
    
    Call Task3
    
    
ErrorHandler:
    
    Call HandleError(Err)

End Sub

Sub HandleError(error As ErrObject)

    If error.Number <> 0 Then
        File.LogTo LogFile, "ERROR : " & UCase(error.Description), Fail
        Call File.TerminateLog(LogFile)
        MsgBox "Something went wrong. Please check the log file.", vbCritical + vbOKOnly, "Prompt"
        On Error GoTo -1
        Exit Sub
    End If
    
    Call File.TerminateLog(LogFile)
    
    MsgBox "Macro finished processing successfully.", vbInformation + vbOKOnly, "Prompt"
    Debug.Print "finished"
    
End Sub

Sub Task1()
    File.LogTo LogFile, "Task1", Start
    
    Dim i As Long
    
    
    For i = 0 To 10000000
    
        i = i + 1
    
    Next i
    
    File.LogTo LogFile, "Task1", Finish
End Sub

Sub Task2()
    File.LogTo LogFile, "Task2", Start
    
    Dim i As Long
    
    
    For i = 0 To 10000000
    
        i = i + 1
    
    Next i
    
    'Err.Raise 9
    
    File.LogTo LogFile, "Task2", Finish
End Sub

Sub Task3()
    File.LogTo LogFile, "Task3", Start
    
    Dim i As Long
    
    
    For i = 0 To 10000000
    
        i = i + 1
    
    Next i
    
    File.LogTo LogFile, "Task3", Finish
End Sub
