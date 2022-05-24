Option Explicit

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const MOUSEEVENTF_LEFTUP = &H4
'Private Const MOUSEEVENTF_RIGHTDOWN As Long = &H8
'Private Const MOUSEEVENTF_RIGHTUP As Long = &H10

Private Const tolerance As Long = 20

Type POINTAPI
         Xpos As Long
         Ypos As Long
End Type
'Create a button on the sheet and assign this macro
Sub RunMacro()
    
    Dim firstPoint As POINTAPI
    Dim lastPoint As POINTAPI
    Dim rng As Range
    
    Set rng = Sheet1.Range("D3")
    
    Call GetCursorPos(firstPoint)
    
    rng.Interior.Color = vbGreen
    
    Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0&, 0&, 0&, 0&)
    Call mouse_event(MOUSEEVENTF_LEFTUP, 0&, 0&, 0&, 0&)
    
    Sleep 4000
    
    Call GetCursorPos(lastPoint)
    
    If firstPoint.Xpos < lastPoint.Xpos + tolerance Or firstPoint.Xpos > lastPoint.Xpos - tolerance Then
        
        If firstPoint.Ypos < lastPoint.Ypos + tolerance Or firstPoint.Ypos > lastPoint.Ypos - tolerance Then
        
            rng.Interior.Color = vbRed
            Exit Sub
            
        End If
    
    End If
    
End Sub
