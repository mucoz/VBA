Option Explicit

'Created by Mustafa Can Ozturk on 22.02.2021
'===========================================
'The functions below will be used for miscellaneous operations
'It includes the most common functions
'It needs to be used in "Misc" module
'Functions:
'Regular expressions (GetTextWithPattern, ReplaceTextWithPattern), Random, PrintFormLayout, GetTextFromClipBoard

'For more info on Regular Expressions :
'https://stackoverflow.com/questions/22542834/how-to-use-regular-expressions-regex-in-microsoft-excel-both-in-cell-and-loops
'https://regexr.com/
'===========================================

'=============================Set text to clipboard=============================
Option Explicit
Private Declare Function OpenClipboard Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function EmptyClipboard Lib "user32.dll" () As Long
Private Declare Function CloseClipboard Lib "user32.dll" () As Long
Private Declare Function IsClipboardFormatAvailable Lib "user32.dll" (ByVal wFormat As Long) As Long
Private Declare Function GetClipboardData Lib "user32.dll" (ByVal wFormat As Long) As Long
Private Declare Function SetClipboardData Lib "user32.dll" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyW" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long

Public Sub SetClipboard(sUniText As String)
    Dim iStrPtr As Long
    Dim iLen As Long
    Dim iLock As Long
    Const GMEM_MOVEABLE As Long = &H2
    Const GMEM_ZEROINIT As Long = &H40
    Const CF_UNICODETEXT As Long = &HD
    OpenClipboard 0&
    EmptyClipboard
    iLen = LenB(sUniText) + 2&
    iStrPtr = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, iLen)
    iLock = GlobalLock(iStrPtr)
    lstrcpy iLock, StrPtr(sUniText)
    GlobalUnlock iStrPtr
    SetClipboardData CF_UNICODETEXT, iStrPtr
    CloseClipboard
End Sub

'=======================================================================


Public Function Random(StartingFrom As Integer, UpTo As Integer) As Integer
    
    Randomize
    
    Random = Int((UpTo - StartingFrom + 1) * Rnd + StartingFrom)
    
End Function



Public Function GetTextWithPattern(Text As String, Pattern As String) As String
    
    Dim RegEx As Object 'WILL BE CHANGED TO LATE BINDING
    Dim matchColl As Object
    
    Set RegEx = CreateObject("vbscript.regexp")
    
    With RegEx
        .MultiLine = False
        .Global = False
        .IgnoreCase = False
        .Pattern = Pattern
    End With
    
    Set matchColl = RegEx.Execute(Text)
    
    If matchColl.Count = 0 Then
    
        GetTextWithPattern = ""
    
    Else
        
        GetTextWithPattern = matchColl.Item(0)
    
    End If

End Function

Function ReplaceTextWithPattern(TextToSearch As String, SearchPattern As String, ReplaceWith As String, _
                      Optional GlobalReplace As Boolean = True, _
                      Optional IgnoreCase As Boolean = False, _
                      Optional MultiLine As Boolean = False) As String
    
    Dim RE As Object

    Set RE = CreateObject("vbscript.regexp")
    
    With RE
        .MultiLine = MultiLine
        .Global = GlobalReplace
        .IgnoreCase = IgnoreCase
        .Pattern = SearchPattern
    End With

    ReplaceTextWithPattern = RE.Replace(TextToSearch, ReplaceWith)
    
End Function

Public Function GetTextFromClipBoard() As String
    

    Dim objData As New DataObject
    Dim strText

       objData.GetFromClipboard
       strText = objData.GetText

       GetTextFromClipBoard = strText

End Function

Public Sub PrintFormLayout(Form As UserForm)
    
    Dim c As Control
    
    For Each c In Form.Controls
        Debug.Print c.Name & ".Left = " & c.Left
        Debug.Print c.Name & ".Top = " & c.Top
        Debug.Print c.Name & ".Height = " & c.Height
        Debug.Print c.Name & ".Width = " & c.Width
        
    Next c
    
End Sub

Function LevenshteinDistance(s1 As String, s2 As String) As Integer
    Dim m As Integer, n As Integer, costMatrix() As Integer, i As Integer, j As Integer
    m = Len(s1)
    n = Len(s2)
    ReDim costMatrix(0 To m, 0 To n)
    For i = 0 To m
        costMatrix(i, 0) = i
    Next i
    For j = 0 To n
        costMatrix(0, j) = j
    Next j
    For i = 1 To m
        For j = 1 To n
            If LCase(Mid(s1, i, 1)) = LCase(Mid(s2, j, 1)) Then
                costMatrix(i, j) = costMatrix(i - 1, j - 1)
            Else
                costMatrix(i, j) = Application.WorksheetFunction.Min( _
                    costMatrix(i - 1, j), _
                    costMatrix(i, j - 1), _
                    costMatrix(i - 1, j - 1) _
                ) + 1
            End If
        Next j
    Next i
    LevenshteinDistance = costMatrix(m, n)
End Function
