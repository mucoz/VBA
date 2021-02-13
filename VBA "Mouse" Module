Option Explicit

'Created by Mustafa Can Ozturk on 13.02.2021
'===========================================
'The functions below will be used for mouse operations
'It includes the most common functions
'It needs to be used in "Mouse" module
'Functions:
'GetPosition, SetPosition, GetPixelColor
'===========================================

Private Declare PtrSafe Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare PtrSafe Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare PtrSafe Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long

Type POINTAPI
    x As Long
    y As Long
End Type

Type RGB
    Red As Long
    Green As Long
    Blue As Long
End Type

Public Property Get GetPixelColor() As RGB
    
    Dim r As Long, g As Long, b As Long
    Dim mousePos As POINTAPI
    Dim longDC As Long
    Dim pixel As Long
    Dim sTmp As String
    
    longDC = GetWindowDC(0)
    
    mousePos = GetPosition
    
    pixel = GetPixel(longDC, mousePos.x, mousePos.y)
    sTmp = Right$("000000" & Hex(pixel), 6)
    
    GetPixelColor.Red = CLng(HexToDec(Right$(sTmp, 2)))
    GetPixelColor.Green = CLng(HexToDec(Mid$(sTmp, 3, 2)))
    GetPixelColor.Blue = CLng(HexToDec(Left$(sTmp, 2)))
    
End Property

Public Property Get GetPosition() As POINTAPI

    Dim newPoint As POINTAPI
    
    GetCursorPos newPoint
    
    GetPosition.x = newPoint.x
    GetPosition.y = newPoint.y
    
End Property

Public Sub SetPosition(x As Long, y As Long)

    SetCursorPos x, y

End Sub

Private Function HexToDec(ByVal Hex As String) As Double
     
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
