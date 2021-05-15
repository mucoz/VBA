Option Explicit

'Created by Mustafa Can Ozturk on 10.02.2021
'===========================================
'The functions below will be used for Sheet operations
'It includes the most common functions

'Functions:
'GetUsedRange, FindHeader, ArrayLookUp, WriteArray, IsSheetEmpty, SheetExists, LR, LRWithColumn, LRUsedRange
'StoreExcelData, DeletePart, GetRowsBetween, GetDataBetween, Clear, ClearAllSheets, FastMode, NormalMode
'MoveUsedRange, CopyUsedRange, GetUniqueItems, DeleteEmptyColumns, IsColumnEmpty, AreRangesSame, FindValueRow(Returns row number)
'===========================================

Type Header
    column As Long
    row As Long
End Type

Public Function GetUniqueItems(Sheet As Worksheet, ColumnNumber As Long, Optional InitialRowNumber As Variant) As Variant

    Dim coll As New Collection
    Dim column As Long, row As Long, i As Long
    Dim data() As Variant
    
    If IsMissing(InitialRowNumber) = True Then
    
        row = 1
    
    Else
    
        row = InitialRowNumber
    
    End If
    
    column = ColumnNumber
    
    Do Until Sheet.Cells(row, column) = ""
        
        On Error Resume Next
        coll.add Sheet.Cells(row, column).Value, CStr(Sheet.Cells(row, column).Value)
        row = row + 1
        
    Loop
    
    For i = 1 To coll.Count
    
        ReDim Preserve data(i - 1)
        data(i - 1) = coll.Item(i)
        
    Next i
    
    GetUniqueItems = data
    
    Set coll = Nothing
    
End Function

Public Sub CopyUsedRange(SourceSheet As Worksheet, TargetSheet As Worksheet, Optional InitialRow As Variant, Optional InitialColumn As Variant)
    
    Dim rngArr As Variant
    Dim startRow As Long
    Dim startColumn As Long
    
    If IsMissing(InitialRow) = True And IsMissing(InitialColumn) = True Then
        
        startRow = 1
        startColumn = 1
    
    Else
    
        startRow = InitialRow
        startColumn = InitialColumn
        
    End If
    
    rngArr = GetUsedRange(SourceSheet)
    
    WriteArray rngArr, TargetSheet, startRow, startColumn
    
End Sub

Public Sub MoveUsedRange(SourceSheet As Worksheet, TargetSheet As Worksheet, Optional InitialRow As Variant, Optional InitialColumn As Variant)
    
    Dim rngArr As Variant
    Dim startRow As Long
    Dim startColumn As Long
    
    If IsMissing(InitialRow) = True And IsMissing(InitialColumn) = True Then
        
        startRow = 1
        startColumn = 1
    
    Else
    
        startRow = InitialRow
        startColumn = InitialColumn
        
    End If
    
    rngArr = GetUsedRange(SourceSheet)
    
    Clear SourceSheet
    
    WriteArray rngArr, TargetSheet, startRow, startColumn
    
End Sub

Public Function GetDataBetween(Sheet As Worksheet, column As String, FirstText As String, LastText As String) As Variant
    
    Dim j As Long
    Dim cell As Range, subcell As Range, datacell As Range
    Dim rng As Range
    Dim FirstTextAddress As String, LastTextAddress As String
    Dim rngData(1) As Variant, data() As Variant
    
    Set rng = Sheet.Range(column & 1, column & Sheet.Rows.Count)
    
    j = 0
    
    For Each cell In rng
        
        If InStr(UCase(cell.Value), UCase(FirstText)) > 0 Then
            
            rngData(0) = cell.row
            Set rng = Sheet.Range(column & rngData(0) + 1, column & Sheet.Rows.Count)
            
            For Each subcell In rng
                
                If InStr(UCase(subcell.Value), UCase(LastText)) > 0 Then
                    
                    rngData(1) = subcell.row
                    Set rng = Sheet.Range(column & rngData(0) + 1, column & rngData(1) - 1)
                    
                    For Each datacell In rng
                        ReDim Preserve data(j)
                        data(j) = datacell.Value
                        j = j + 1
                    Next datacell
                    
                    GetDataBetween = data
                    Exit Function
                
                End If
            
            Next subcell
        
        End If
    
    Next cell
    
    GetDataBetween = "Empty"

End Function

Public Function GetRowsBetween(Sheet As Worksheet, column As String, FirstText As String, LastText As String) As Variant
    
    Dim j As Long
    Dim cell As Range, subcell As Range, datacell As Range
    Dim rng As Range
    Dim FirstTextAddress As String, LastTextAddress As String
    Dim rngData(1) As Variant, data() As Variant
    
    Set rng = ActiveSheet.Range(column & 1, column & Sheet.Rows.Count)
    
    j = 0
    
    For Each cell In rng
        
        If InStr(UCase(cell.Value), UCase(FirstText)) > 0 Then
            
            rngData(0) = cell.row
            Set rng = ActiveSheet.Range(column & rngData(0) + 1, column & Sheet.Rows.Count)
            
            For Each subcell In rng
                
                If InStr(UCase(subcell.Value), UCase(LastText)) > 0 Then
                    
                    rngData(1) = subcell.row
                    Set rng = ActiveSheet.Range(column & rngData(0) + 1, column & rngData(1) - 1)
                    
                    For Each datacell In rng
                        ReDim Preserve data(j)
                        data(j) = datacell.Address
                        j = j + 1
                    Next datacell
                    
                    GetRowsBetween = data
                    Exit Function
                
                End If
            
            Next subcell
        
        End If
    
    Next cell
    
    GetRowsBetween = "Empty"

End Function

Sub DeletePart(Worksheet As Worksheet, startRow As Long, FinishRow As Long)
    
    Dim ws As Worksheet
    
    Set ws = Worksheet
    
    Dim i As Long
    
    If FinishRow < startRow Then Exit Sub
    
    For i = startRow To FinishRow
        Worksheet.Rows(startRow).EntireRow.Delete
    Next i

End Sub
Public Sub WriteArray(DataArray As Variant, SheetToWrite As Worksheet, Optional InitialRow As Variant, Optional InitialColumn As Variant)

    Dim i As Long
    Dim j As Long
    Dim startRow As Long
    Dim startColumn As Long
    Dim row As Long
    Dim column As Long
    
    If IsMissing(InitialRow) = True And IsMissing(InitialColumn) = True Then
        
        startRow = 1
        startColumn = 1
    
    Else
    
        startRow = InitialRow
        startColumn = InitialColumn
        
    End If
    
    row = startRow
    
    For i = 1 To UBound(DataArray, 1)
    
        column = startColumn
        
        For j = 1 To UBound(DataArray, 2)
    
            SheetToWrite.Cells(row, column) = DataArray(i, j)
            column = column + 1
        
        Next j
        
        row = row + 1
    
    Next i

End Sub
Function StoreExcelData(filePath As String, Optional SheetName As Variant, Optional Password As Variant) As Variant
    
    Dim WBdata As Workbook
    Dim WSdata As Worksheet
    Dim rngData As Range
    Dim arr() As Variant
    Dim LR As Long, LC As Long, i As Long, j As Long
    
    If filePath = "" Then
        
        StoreExcelData = Array("")
        GoTo exitFunction
    
    End If
    
    FastMode
    
    'Set workbook with/without password
    If IsMissing(Password) = True Then
        Set WBdata = Workbooks.Open(filePath, , True)
    Else
        Set WBdata = Workbooks.Open(filePath, , True, , Password)
    End If
    
    'Set worksheet with/without name
    
    If IsMissing(SheetName) = True Then
        
        Set WSdata = WBdata.Sheets(1)
    
    Else
        
        Set WSdata = WBdata.Sheets(SheetName)
    
    End If
    
    Set rngData = WSdata.UsedRange
    LR = rngData.Rows.Count
    LC = rngData.Columns.Count
    ReDim arr(1 To LR, 1 To LC)
    
    For i = 1 To LR
        
        For j = 1 To LC
            arr(i, j) = WSdata.Cells(i, j).Value
        Next j
    
    Next i
    
    StoreExcelData = arr()
    
    WBdata.Close
    
    Set WBdata = Nothing
    Set WSdata = Nothing
    Set rngData = Nothing
    
    NormalMode

exitFunction:
End Function

Public Sub FastMode()

    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With

End Sub
Public Sub NormalMode()
    
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
    End With

End Sub

Public Sub Clear(Optional Sheet As Variant)
    
    Dim ws As Object
    
    If IsMissing(Sheet) = True Then
        
        Set ws = ActiveSheet
    
    Else
    
        Set ws = Sheet
    
    End If
    
    ws.Cells.Delete
    
End Sub

Public Sub ClearAllSheets(Optional WB As Variant)
    
    Dim wbMain As Workbook
    Dim ws As Worksheet
    
    Dim answer As Long
    
    answer = MsgBox("Do you want to clear every sheet in this workbook?", vbQuestion + vbYesNo, "Prompt")
    
    If answer = vbNo Then
        
        Exit Sub
    
    Else
        
    If IsMissing(WB) = True Then
    
        Set wbMain = ActiveWorkbook
        
    Else
        
        Set wbMain = WB
    
    End If
        
        For Each ws In wbMain.Worksheets
            ws.Cells.Delete
        Next ws
    
    End If

End Sub

Public Function SheetExists(WB As Workbook, SheetName As String) As Boolean

    Dim ws As Worksheet

    SheetExists = False
    
    For Each ws In WB.Sheets
        
        If ws.Name = SheetName Then
            
            SheetExists = True
        
        End If
    
    Next ws

End Function

Public Function IsSheetEmpty(Sheet As Worksheet) As Boolean
    'Adress
    Dim add As String
    add = "$A$1"
    
    If Sheet.UsedRange.Address = add And Sheet.Range("A1") = "" Then
        IsSheetEmpty = True
        Exit Function
    End If
    
    IsSheetEmpty = False
    
End Function

Public Function GetUsedRange(Sheet As Worksheet) As Variant

    GetUsedRange = Sheet.UsedRange
    
End Function

Public Function ArrayLookUp(ByVal SearchItem As Variant, LookUpRange As Variant, FirstColumnNumber As Long, LastColumnNumber As Long) As Variant
    
    Dim i As Long
    
    For i = 1 To UBound(LookUpRange, 1)
        
        If IsNumeric(LookUpRange(i, FirstColumnNumber)) = True Then
            
            If SearchItem = Val(LookUpRange(i, FirstColumnNumber)) Then
                ArrayLookUp = LookUpRange(i, LastColumnNumber)
                Exit Function
            End If
        
        Else
            
            If SearchItem = CStr(LookUpRange(i, FirstColumnNumber)) Then
                ArrayLookUp = LookUpRange(i, LastColumnNumber)
                Exit Function
            End If
        
        End If
    
    Next i
    
    ArrayLookUp = "NOT FOUND"

End Function

Public Function FindHeader(LookUpRange As Variant, Text As String, Optional ExactMatch As Boolean = True) As Header

    Dim i As Long
    Dim j As Long
    
    For i = 1 To UBound(LookUpRange, 1)
        
        For j = 1 To UBound(LookUpRange, 2)
        
            If ExactMatch = True Then
        
                If LookUpRange(i, j) = Text Then
                
                    FindHeader.column = j
                    FindHeader.row = i
                    Exit Function
                    
                End If
            
            Else
            
                If InStr(UCase(LookUpRange(i, j)), UCase(Text)) > 0 Then
                    
                    FindHeader.column = j
                    FindHeader.row = i
                    Exit Function
                    
                End If
                
            End If
        
        Next j
        
    Next i

    FindHeader.column = -1
    FindHeader.row = -1
    
End Function
                        
Public Sub DeleteEmptyColumns(Rng As Range)

    Dim i As Long
    Dim j As Long
    
    
    For j = Rng.Columns(1).Column To Rng.Columns.Count
        
        For i = Rng.Rows(1).Row To Rng.Rows.Count
        
            If Trim(Rng.Parent.Cells(i, j)) <> "" Then
            
                GoTo nextJ
                
            End If
            
        Next i
        
        Rng.Parent.Columns(j).EntireColumn.Delete
nextJ:
    Next j

End Sub

Public Function IsColumnEmpty(Sheet As Worksheet, ColumnNameOrNumber As String) As Boolean

    Dim i As Long
    Dim lr As Long
    
    
    If IsNumeric(ColumnNameOrNumber) = False Then
        lr = Sheet.Range(ColumnNameOrNumber & Rows.Count).End(xlUp).Row
    Else
        lr = Sheet.Cells(Rows.Count, ColumnNameOrNumber).End(xlUp).Row
    End If
    
    IsColumnEmpty = True
    
    For i = 1 To lr
    
        If Trim(Sheet.Range(ColumnNameOrNumber & i)) <> vbNullString Then
        
            IsColumnEmpty = False
            Exit Function
        End If
    
    Next i

End Function

Public Function AreRangesSame(Rng1 As Range, Rng2 As Range) As Boolean

    Dim r1 As Long, c1 As Long, r2 As Long, c2 As Long
    
    r1 = Rng1.Rows.Count
    c1 = Rng1.Columns.Count
    
    r2 = Rng2.Rows.Count
    c2 = Rng2.Columns.Count
    
    If Rng1 Is Nothing Then Debug.Print "First Range is not set": Exit Function
    If Rng2 Is Nothing Then Debug.Print "Second Range is not set": Exit Function
    
    If r1 <> r2 And c1 <> c2 Then
    
        AreRangesSame = False
        Exit Function
        
    End If
    
    Dim i As Long, j As Long
    Dim counter As Long
    counter = 0
    For i = 1 To r1
        For j = 1 To c1
            If Rng1(i, j) = Rng2(i, j) Then
                counter = counter + 1
            End If
        Next j
    Next i
    
    If counter = r1 * c1 Then
        AreRangesSame = True
    Else
        AreRangesSame = False
    End If

End Function
            
Public Function FindValueRow(Sheet As Worksheet, Value As String) As Long
    Dim FindString As String
    Dim Rng As Range
    FindString = Value
    If Trim(FindString) <> "" Then
        With Sheet.Range("A:Z")
            Set Rng = .Find(What:=FindString, _
                            After:=.Cells(.Cells.Count), _
                            LookIn:=xlValues, _
                            LookAt:=xlWhole, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlNext, _
                            MatchCase:=False)
            If Not Rng Is Nothing Then
                FindValueRow = Rng.Row
            End If
        End With
    End If
End Function

