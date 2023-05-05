Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'      Author  : Mustafa Can Öztürk                                                                     '
'      Purpose : It creates a log file in text format                                                   '
'      Input   : Use Start and Finish methods at the beginning and end of the process respectively      '
'              : Use "Info", "Warn", "Error" methods to write appropriate messages                      '
'              : Use SetFolderPath method to create the file in a different location                    '
'      Output  : Structured log file with time stamps, info level, and messages                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private m_logFolderPath As String
Private g_errorOccurred As Boolean

Enum Status
    START_LOG
    FINISH_LOG
    WRITE_INFO
    WRITE_WARN
    WRITE_ERROR
End Enum

Private Const LOG_FOLDER_NAME = "Log"
Private Const LOG_FILE_NAME = "Log"
Private Const MSG_SEP As String = " | "
Private Const MSG_INFO As String = "INFO"
Private Const MSG_WARN As String = "WARN"
Private Const MSG_ERROR As String = "ERROR"
Private Const MSG_CHAR As String = "="
Private Const CHAR_LEN As Integer = 25

Public Sub SetFolderPath(Path As String)
    m_logFolderPath = Path
End Sub

Public Sub Start()
    Call WriteToFile("PROCESS STARTED BY " + GetUsername(), START_LOG)
End Sub

Public Sub Finish()
    Call WriteToFile("END OF THE PROCESS", FINISH_LOG)
End Sub

Public Sub Info(Text As String)
    Call WriteToFile(Text, WRITE_INFO)
End Sub

Public Sub Warn(Text As String)
    Call WriteToFile(Text, WRITE_WARN)
End Sub

Public Sub Error(Text As String)
    Call WriteToFile(Text, WRITE_ERROR)
End Sub

Private Sub WriteToFile(Text As String, Stat As Status)
On Error GoTo error_handler
    Dim logFolderPath As String, logFilePath As String, logFile As Object, timeStamp As String
    If g_errorOccurred = True Then
        Exit Sub
    End If
    If ThisWorkbook.Path = "" And m_logFolderPath = "" Then
        Call MsgBox("Path of the workbook is undefined." + vbNewLine + "Set folder path or save the workbook first!", vbCritical, "Logger Error")
        g_errorOccurred = True
        Exit Sub
    End If
    If m_logFolderPath <> "" Then
        If FSO.FolderExists(m_logFolderPath) = True Then
            logFolderPath = FSO.BuildPath(m_logFolderPath, LOG_FOLDER_NAME)
            logFilePath = FSO.BuildPath(logFolderPath, LOG_FILE_NAME & "_" & Format(Date, "yyyy_mm_dd") & ".txt")
        Else
            Call MsgBox("""" & m_logFolderPath & """ is not a folder!", vbCritical, "Logger Error")
            g_errorOccurred = True
            Exit Sub
        End If
    Else
        logFolderPath = ThisWorkbook.Path & "\" & LOG_FOLDER_NAME
        logFilePath = logFolderPath & "\" & LOG_FILE_NAME & "_" & Format(Date, "yyyy_mm_dd") & ".txt"
    End If
    timeStamp = Format(Now, "dd.mm.yyyy HH:MM:SS")
    If Dir(logFolderPath, vbDirectory) = "" Then Call MkDir(logFolderPath)
    If Dir(logFilePath) = "" Then
        Set logFile = FSO.CreateTextFile(logFilePath, True)
    Else
        Set logFile = FSO.OpenTextFile(logFilePath, 8)
    End If
    If Stat = START_LOG Then
        Call logFile.WriteLine(timeStamp & MSG_SEP & MSG_INFO & MSG_SEP & String(CHAR_LEN, MSG_CHAR) & Text & String(CHAR_LEN, MSG_CHAR))
    ElseIf Stat = FINISH_LOG Then
        Call logFile.WriteLine(timeStamp & MSG_SEP & MSG_INFO & MSG_SEP & String(CHAR_LEN, MSG_CHAR) & Text & String(CHAR_LEN, MSG_CHAR))
    ElseIf Stat = WRITE_INFO Then
        Call logFile.WriteLine(timeStamp & MSG_SEP & MSG_INFO & MSG_SEP & Text)
    ElseIf Stat = WRITE_WARN Then
        Call logFile.WriteLine(timeStamp & MSG_SEP & MSG_WARN & MSG_SEP & Text)
    ElseIf Stat = WRITE_ERROR Then
        Call logFile.WriteLine(timeStamp & MSG_SEP & MSG_ERROR & MSG_SEP & Text)
    End If
    Call logFile.Close
    Set logFile = Nothing
    Exit Sub
error_handler:
    Call MsgBox("Error occurred while writing log file!", vbCritical)
    g_errorOccurred = True
End Sub

Private Function GetUsername() As String
    GetUsername = Environ("USERNAME")
End Function
Private Function FSO() As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
End Function
