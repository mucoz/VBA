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

Private Const LOG_FOLDER_NAME = "Logs"
Private Const LOG_FILE_NAME = "Log"
Private Const MSG_SEP As String = " | "

Private Const MSG_START As String = "START"
Private Const MSG_FINISH As String = "FINISH"
Private Const MSG_INFO As String = "INFO"
Private Const MSG_WARN As String = "WARNING"
Private Const MSG_ERROR As String = "ERROR"

Private Const MSG_CHAR As String = "="
Private Const CHAR_LEN As Integer = 25
Private Const LOG_LEVEL_WIDTH As Integer = 7


Public Sub SetFolderPath(Path As String)
    m_logFolderPath = Path
End Sub

Public Sub Start()
    Call WriteToFile("PROCESS STARTED BY " + GetUsername(), MSG_START)
End Sub

Public Sub Finish()
    Call WriteToFile("END OF THE PROCESS", MSG_FINISH)
End Sub

Public Sub Info(Text As String)
    Call WriteToFile(Text, MSG_INFO)
End Sub

Public Sub Warn(Text As String)
    Call WriteToFile(Text, MSG_WARN)
End Sub

Public Sub Error(Text As String)
    Call WriteToFile(Text, MSG_ERROR)
End Sub

Private Sub WriteToFile(Text As String, Stat As String)
    On Error GoTo error_handler
    Dim logFolderPath As String
    Dim logFilePath As String
    Dim logFile As Object
    Dim timeStamp As String
    Dim logMessage As String
    If ThisWorkbook.Path = "" Then
        Call MsgBox("Path of the workbook is undefined." + vbNewLine + "Set folder path or save the workbook first!", vbCritical, "Logger Error")
        Exit Sub
    End If
    If m_logFolderPath <> "" Then
        If FSO.FolderExists(m_logFolderPath) = True Then
            logFolderPath = FSO.BuildPath(m_logFolderPath, LOG_FOLDER_NAME)
            logFilePath = FSO.BuildPath(logFolderPath, LOG_FILE_NAME & "_" & Format(Date, "yyyy_mm_dd") & ".txt")
        Else
            Call MsgBox("""" & m_logFolderPath & """ is not a folder!", vbCritical, "Logger Error")
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
    logMessage = ""
    If Stat = MSG_START Or Stat = MSG_FINISH Then
        logMessage = timeStamp & MSG_SEP & FormatLogLevel(MSG_INFO) & MSG_SEP & String(CHAR_LEN, MSG_CHAR) & Text & String(CHAR_LEN, MSG_CHAR)
    Else
        logMessage = timeStamp & MSG_SEP & FormatLogLevel(Stat) & MSG_SEP & Text
    End If
    Call logFile.WriteLine(logMessage)
    Call logFile.Close
    Debug.Print logMessage
    Set logFile = Nothing
    Exit Sub
error_handler:
    Call MsgBox("Error occurred while writing log file!", vbCritical)
End Sub

Private Function FormatLogLevel(level As String) As String
    
    Dim paddedLevel As String
    paddedLevel = level
    
    Do While Len(paddedLevel) < LOG_LEVEL_WIDTH
        paddedLevel = paddedLevel & " "
    Loop
    
    FormatLogLevel = paddedLevel
End Function

Private Function GetUsername() As String
    GetUsername = Environ("USERNAME")
End Function
Private Function FSO() As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
End Function
