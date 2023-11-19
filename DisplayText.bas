Option Explicit

Private Sub MakeFullScreen()
    Application.DisplayFullScreen = True
    Application.DisplayFormulaBar = False
    ActiveWindow.DisplayWorkbookTabs = False
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.DisplayGridlines = False
    Application.DisplayStatusBar = False
End Sub

Private Sub CancelFullScreen()
    Application.DisplayFullScreen = False
    Application.DisplayFormulaBar = True
    ActiveWindow.DisplayWorkbookTabs = True
    ActiveWindow.DisplayHeadings = True
    ActiveWindow.DisplayGridlines = True
    Application.DisplayStatusBar = True
End Sub

Private Sub ClearBoard()
    Dim r As Range
    Set r = Range("A1:DD1")
    r.Value = ""
    r.ColumnWidth = 4
End Sub

Private Sub initializeColorPalette(ByRef orangePalette As Collection, ByRef blackPalette As Collection)
    ' gray to orange
    Call orangePalette.Add(RGB(110, 104, 95))
    Call orangePalette.Add(RGB(123, 106, 88))
    Call orangePalette.Add(RGB(136, 108, 81))
    Call orangePalette.Add(RGB(148, 110, 75))
    Call orangePalette.Add(RGB(161, 112, 68))
    Call orangePalette.Add(RGB(174, 114, 61))
    Call orangePalette.Add(RGB(187, 116, 54))
    Call orangePalette.Add(RGB(200, 118, 47))
    Call orangePalette.Add(RGB(213, 120, 40))
    Call orangePalette.Add(RGB(225, 122, 34))
    Call orangePalette.Add(RGB(238, 124, 27))
    Call orangePalette.Add(RGB(251, 126, 20))
    
    'gray to black
    Call blackPalette.Add(RGB(110, 104, 95))
    Call blackPalette.Add(RGB(100, 95, 86))
    Call blackPalette.Add(RGB(90, 85, 78))
    Call blackPalette.Add(RGB(80, 76, 69))
    Call blackPalette.Add(RGB(70, 66, 60))
    Call blackPalette.Add(RGB(60, 57, 52))
    Call blackPalette.Add(RGB(50, 47, 43))
    Call blackPalette.Add(RGB(40, 38, 35))
    Call blackPalette.Add(RGB(30, 28, 26))
    Call blackPalette.Add(RGB(20, 19, 17))
    Call blackPalette.Add(RGB(10, 9, 9))
    Call blackPalette.Add(RGB(0, 0, 0))
End Sub

Sub RunMacro()
    Call MakeFullScreen
    Call WriteText("HELLO GORDON. MY NAME IS MUSTAFA", 2, 5)
    Call WriteText("THIS IS HALF-LIFE STYLE TEXT", 17, 15)
    Call CancelFullScreen
End Sub

Private Sub WriteText(message As String, rowIndex As Long, colIndex As Long)
    Dim i As Integer, j As Integer, orangePalette As New Collection, blackPalette As New Collection, letterIndex As Long
    Call initializeColorPalette(orangePalette, blackPalette)
    Dim counter As Integer, colorCounter As Integer
    Call ClearBoard
    Dim t As Double
    t = 0.05
    letterIndex = 1
    For i = colIndex To colIndex + Len(message)
        Sheet1.Cells(rowIndex, i) = Mid(message, letterIndex, 1)
        Sheet1.Cells(rowIndex, i).Font.Color = orangePalette(12)
        Sheet1.Cells(rowIndex, i).Font.Size = 20
        Sheet1.Cells(rowIndex, i).Font.Bold = True
        letterIndex = letterIndex + 1
        If i - 1 > 0 Then
        counter = 0
            For j = i To 1 Step -1
                If counter >= 12 Then
                    Sheet1.Cells(rowIndex, j).Font.Color = orangePalette(1)
                Else
                    Sheet1.Cells(rowIndex, j).Font.Color = orangePalette(12 - counter)
                End If
                counter = counter + 1
            Next j
        End If
        TimeOut (t)
    Next i
    Dim decrease As Integer
    If Len(message) > 12 Then
        decrease = 12
        counter = 12
    Else
        decrease = Len(message)
        counter = Len(message)
    End If
    ' Make the last letters gray gradually
    For i = Len(message) + colIndex - decrease To Len(message) + colIndex
        colorCounter = 1
        For j = i To colIndex + Len(message)
            If colorCounter <= counter Then
                Sheet1.Cells(rowIndex, j).Font.Color = orangePalette(colorCounter)
                colorCounter = colorCounter + 1
            End If
        Next j
        Call TimeOut(t)
    Next i
    ' Make the text disappear
    Call TimeOut(1)
    For j = 1 To 12
    For i = colIndex To colIndex + Len(message)
        Sheet1.Cells(rowIndex, i).Font.Color = blackPalette(j)
    Next i
        Call TimeOut(t)
    Next j
End Sub

Private Sub TimeOut(duration As Double)
    Dim st As Double
    st = Timer
    Do
        DoEvents
    Loop Until (Timer - st) >= duration
End Sub
