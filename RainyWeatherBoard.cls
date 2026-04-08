Option Explicit

Private m_ws As Worksheet
Private m_row As Integer
Private m_col As Integer

Private Sub Class_Initialize()
    Set m_ws = ThisWorkbook.Sheets("Sheet1")
    m_row = 0
    m_col = 0
End Sub

Private Sub Class_Terminate()
    Set m_ws = Nothing
End Sub

Public Sub CreateNew(width As Double, height As Double)
    m_ws.Cells.Delete
    m_ws.Columns.ColumnWidth = width
    m_ws.Rows.RowHeight = height
End Sub

Private Function GetLastVisibleCellAddress() As String
    Dim visRange As Range
    Set visRange = ActiveWindow.VisibleRange
    GetLastVisibleCellAddress = visRange.Cells(visRange.Rows.Count, visRange.Columns.Count).Address
End Function

Private Function SetupBoardAsColl() As Collection
    Dim lastVisibleCellAddress As String
    lastVisibleCellAddress = GetLastVisibleCellAddress
    m_row = m_ws.Range(lastVisibleCellAddress).Row
    m_col = m_ws.Range(lastVisibleCellAddress).Column
End Function

Private Function GenerateNewRainDrops() As Collection
    Dim i As Integer
    Dim j As Integer
    Dim pxl As Pixel
    Dim pxlColl As New Collection
    For i = 1 To m_col
        If Rnd < 0.005 Then
            Set pxl = New Pixel
            pxl.x = 1
            pxl.y = i
            pxl.color = 12
            pxlColl.Add pxl
        End If
    Next i
    Set GenerateNewRainDrops = pxlColl
End Function

Public Sub StartRain(color As Integer)
    Call SetupBoardAsColl
    Dim i As Integer
    Dim j As Integer
    Dim pxlColl As Collection
    Dim dropColl As Collection
    Set pxlColl = New Collection
    
    On Error Resume Next
    Do
        Set dropColl = GenerateNewRainDrops
        pxlColl.Add dropColl
        For i = 1 To pxlColl.Count
            For j = 1 To pxlColl(i).Count
                pxlColl(i)(j).x = pxlColl(i)(j).x + 1
                m_ws.Cells(pxlColl(i)(j).x, pxlColl(i)(j).y).Interior.ColorIndex = color
            Next j
        Next i
        If pxlColl(1)(1).x >= m_row Then
            pxlColl.Remove 1
        End If
        m_ws.Cells.Delete
        DoEvents
    Loop
End Sub
