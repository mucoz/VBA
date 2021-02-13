Option Explicit

'Created by Mustafa Can Ozturk on 10.02.2021
'===========================================
'The functions below will be used for file operations
'It includes the most common functions
'It needs to be used in "Window" module
'Functions:
'ActivateWindow, CloseWindow, FindAppNameFromText, MinimizeWindow, MaximizeWindow, MoveWindow, ResizeWindow, GetAppNames
'GetAppNames, Show, Hide, Sleep, PrintAppNames, OpenApplication
'Properties:
'GetActiveAppHandle, GetActiveAppName, GetAppNameWithHandle, GetHandle, IsActive, IsOpen, IsVisible,
'===========================================


'====API's to get application names=============
'#If Win64 Or VBA7 Then

'Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
'Declare PtrSafe Function GetDesktopWindow Lib "user32" () As Long
'Declare PtrSafe Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
'Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
'Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal aint As Long) As Long

'#Else

Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As Long
Private Declare PtrSafe Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal aint As Long) As Long
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As Long
Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare PtrSafe Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Public Declare PtrSafe Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
'#End If


'===================================================================================
'CONSTANTS
'These are the constants for Windows API's
Const GW_CHILD = 5
Const GW_HWNDNEXT = 2
Const GWL_STYLE = (-16)
Const WS_VISIBLE = &H10000000
Const MAX_LEN = 255

Const WM_CLOSE = &H10
Const WM_ACTIVATE = &H6
Const WM_SHOWWINDOW = &H18
Const WA_ACTIVE = 1
Const WM_ACTIVATEAPP = &H1C
Const WM_SETFOCUS = &H7
Const WM_ENABLE = &HA

' ShowWindow() Commands
Const SW_HIDE = 0
Const SW_SHOWNORMAL = 1
Const SW_NORMAL = 1
Const SW_SHOWMINIMIZED = 2
Const SW_SHOWMAXIMIZED = 3
Const SW_MAXIMIZE = 3
Const SW_SHOWNOACTIVATE = 4
Const SW_SHOW = 5
Const SW_MINIMIZE = 6
Const SW_SHOWMINNOACTIVE = 7
Const SW_SHOWNA = 8
Const SW_RESTORE = 9
Const SW_SHOWDEFAULT = 10
Const SW_MAX = 10


Global Const WM_SETTEXT As Long = &HC


Public Property Get GetActiveAppName() As String
    
    GetActiveAppName = GetAppNameWithHandle(GetForegroundWindow)
    
End Property

Public Property Get GetActiveAppHandle() As Long

    GetActiveAppHandle = GetForegroundWindow
    
End Property

Public Property Get GetAppNameWithHandle(ByVal hwnd As Long) As String
    
    GetAppNameWithHandle = String(GetWindowTextLength(hwnd) + 1, Chr$(0))
    
    Call GetWindowText(hwnd, GetAppNameWithHandle, Len(GetAppNameWithHandle))
    
End Property

Public Function GetAppNames() As Variant
    
    Dim data As New Collection                '() As String
    Dim DataArr() As Variant
    Dim xStr As String
    Dim xStrLen As Long
    Dim xHandle As Long
    Dim xHandleStr As String
    Dim xHandleLen As Long, xHandleStyle As Long
    
    On Error Resume Next
    
    Dim i As Long
    
    i = 0
    xHandle = GetWindow(GetDesktopWindow(), GW_CHILD)
    
    Do While xHandle <> 0
        xStr = String$(MAX_LEN - 1, 0)
        xStrLen = GetWindowText(xHandle, xStr, MAX_LEN)
        If xStrLen > 0 Then
            xStr = Left$(xStr, xStrLen)
            xHandleStyle = GetWindowLong(xHandle, GWL_STYLE)
            If xHandleStyle And WS_VISIBLE Then
'                ReDim Preserve DataArr(i)
'                DataArr(i) = xStr
                data.Add xStr, CStr(xStr)
'                i = i + 1
            End If
        End If
        xHandle = GetWindow(xHandle, GW_HWNDNEXT)
    Loop
    
    For i = 1 To data.Count

        ReDim Preserve DataArr(i - 1)

        DataArr(i - 1) = data.Item(i)

    Next i
    
    GetAppNames = DataArr

End Function

Public Sub PrintAppNames()
    
    Dim appList As Variant
    Dim i As Long
    
    appList = GetAppNames
    
    For i = LBound(appList) To UBound(appList)
    
        Debug.Print appList(i)
        
    Next i
    
End Sub

Public Function CloseWindow(WindowName As String) As Long

    Dim hwnd As Long
    Dim result As Long
    
    hwnd = FindWindow(vbNullString, WindowName)
    
    result = SendMessage(hwnd, WM_CLOSE, 0, 0)
    
    CloseWindow = result
    
End Function

Public Function ActivateWindow(WindowName As String) As Long

    Dim hwnd As Long
    Dim result As Long
    
    hwnd = FindWindow(vbNullString, WindowName)
    
    result = SetForegroundWindow(hwnd)
    
    result = ShowWindow(hwnd, SW_RESTORE)
    
    ActivateWindow = result
    
End Function

Public Function MinimizeWindow(WindowName As String) As Long
    
    Dim hwnd As Long
    Dim result As Long
    
    hwnd = FindWindow(vbNullString, WindowName)
    
    result = ShowWindow(hwnd, SW_MINIMIZE)
    
    MinimizeWindow = result
    
End Function

Public Function MaximizeWindow(WindowName As String) As Long

    Dim hwnd As Long
    Dim result As Long
    
    hwnd = FindWindow(vbNullString, WindowName)
    
    result = ShowWindow(hwnd, SW_MAXIMIZE)
    
    MaximizeWindow = result
    
End Function

Public Function ResizeWindow(WindowName As String, PositionX As Long, PositionY As Long, Width As Long, Height As Long) As Long
    
    Dim hwnd As Long
    Dim result As Long
    
    hwnd = FindWindow(vbNullString, WindowName)
    
    result = MoveWindow(hwnd, PositionX, PositionY, Width, Height, 0)
    
    ResizeWindow = result
    
End Function

Public Property Get GetHandle(WindowName As String) As Long
    
    Dim hwnd As Long
    
    hwnd = FindWindow(vbNullString, WindowName)
    
    GetHandle = hwnd
    
End Property

Public Property Get IsActive(WindowNameOrHandle As Variant) As Boolean

    Dim hwnd As Long
    Dim WindowName As String
    Dim status As Boolean
    
    If IsNumeric(WindowNameOrHandle) = True Then
    
        hwnd = WindowNameOrHandle
        
        If GetAppNameWithHandle(GetForegroundWindow) = GetAppNameWithHandle(hwnd) Then
        
            status = True
            
        Else
         
            status = False
            
        End If
        
        IsActive = status
        
    Else
    
        WindowName = WindowNameOrHandle
        hwnd = FindWindow(vbNullString, WindowName)
        
        If GetAppNameWithHandle(GetForegroundWindow) = GetAppNameWithHandle(hwnd) Then
        
            status = True
            
        Else
         
            status = False
            
        End If
        
        IsActive = status
        
    End If



End Property
Public Property Get IsVisible(WindowNameOrHandle As Variant) As Boolean

    Dim hwnd As Long
    Dim WindowName As String
    Dim status As Boolean
    
    If IsNumeric(WindowNameOrHandle) = True Then
    
        hwnd = WindowNameOrHandle
        status = IsWindowVisible(hwnd)
        IsVisible = status
        
    Else
    
        WindowName = WindowNameOrHandle
        hwnd = FindWindow(vbNullString, WindowName)
        status = IsWindowVisible(hwnd)
        IsVisible = status
        
    End If
    
End Property

Public Property Get IsOpen(WindowNameOrHandle As Variant) As Boolean
    
    Dim hwnd As Long
    Dim WindowName As String
    Dim status As Boolean
    
    If IsNumeric(WindowNameOrHandle) = True Then
    
        hwnd = WindowNameOrHandle
        status = IsWindowEnabled(hwnd)
        IsOpen = status
        
    Else
    
        WindowName = WindowNameOrHandle
        hwnd = FindWindow(vbNullString, WindowName)
        status = IsWindowEnabled(hwnd)
        IsOpen = status
        
    End If

End Property

Public Function FindAppNameFromText(PartialName As String) As String
    
    Dim list As Variant
    Dim i As Long
    
    list = GetAppNames
    
    For i = LBound(list) To UBound(list)
        
        If InStr(UCase(list(i)), UCase(PartialName)) > 0 Then
            
            FindAppNameFromText = list(i)
            Exit Function
        
        End If
        
    Next i
    
    FindAppNameFromText = "Not Found"
    
End Function

Public Function OpenApplication(WindowName As String, AppFullPath As String) As Long
    
    'Returns Application's Handle
    
    Dim hwnd As Long
    Dim result As Long
    
    hwnd = FindWindow(vbNullString, WindowName)
    
    result = ShellExecute(hwnd, "open", AppFullPath, vbNullString, File.ParentFolder(AppFullPath), SW_NORMAL)
    
    If result = 2 Or result = 3 Then Exit Function
    
    Do
        DoEvents
        hwnd = FindWindow(vbNullString, WindowName)
    Loop Until hwnd > 0
    
    OpenApplication = hwnd
    
End Function

Public Function Hide(WindowName As String) As Long

    Dim hwnd As Long
    Dim result As Long
    
    hwnd = FindWindow(vbNullString, WindowName)
    
    result = ShowWindow(hwnd, SW_HIDE)
    
    Hide = result
    
End Function

Public Function Show(WindowName As String) As Long
    
    Dim hwnd As Long
    Dim result As Long
    
    hwnd = FindWindow(vbNullString, WindowName)
    
    result = ShowWindow(hwnd, SW_SHOW)
    
    Show = result

End Function

Public Function HexToDec(Hex As String) As Double
     
    Dim i               As Long
    Dim j               As Variant
    Dim k               As Long
    Dim n               As Long
    Dim HexArray()      As Double
     
    n = Len(Hex)
    k = -1
    ReDim HexArray(1 To n)
    For i = n To 1 Step -1
        j = Mid(Hex, i, 1)
        k = k + 1
        Select Case j
        Case 0 To 9
            HexArray(i) = j * 16 ^ (k)
        Case Is = "A"
            HexArray(i) = 10 * 16 ^ (k)
        Case Is = "B"
            HexArray(i) = 11 * 16 ^ (k)
        Case Is = "C"
            HexArray(i) = 12 * 16 ^ (k)
        Case Is = "D"
            HexArray(i) = 13 * 16 ^ (k)
        Case Is = "E"
            HexArray(i) = 14 * 16 ^ (k)
        Case Is = "F"
            HexArray(i) = 15 * 16 ^ (k)
        End Select
    Next i
    
    HexToDec = Application.WorksheetFunction.Sum(HexArray)
     
End Function
