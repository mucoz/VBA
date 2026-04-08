Option Explicit

Public Sub EntryPoint()

    Dim brd As Board
    Set brd = New Board
    Call brd.CreateNew(0.5, 15)
    Call brd.StartRain(23)
    
End Sub
