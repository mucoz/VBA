Option Explicit


'OpenGetUsedRange


Public Function XLConnectionString(Path As String) As String
'https://www.connectionstrings.com/excel/
    XLConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                         "Data Source=" & Path & ";" & _
                         "Extended Properties=""Excel 12.0;IMEX=1" 'IMEX : Treat all data as text.Safer way to get data
                    
End Function

Sub main()
    
    Dim a As Variant
    Dim st As Double
    
    st = Timer
    
    a = GetUsedRange("C:\Users\moeztuerk\Desktop\Tax\Bot Requirements\680.14135 23500 VAT UK Dec-20.xlsb", "Daybook 680")
    
    WriteArray "C:\Users\moeztuerk\Desktop\deneme.xlsx", a, , 2, 2
    
    Debug.Print Format(Timer - st, "00.00")

'    Stop

End Sub

Function OpenAndGetData(FilePath As String, Optional SheetName As Variant, Optional Password As Variant) As Variant
    
    Dim WBdata As Workbook
    Dim WSdata As Worksheet
    Dim status As Boolean
    
    If FilePath = "" Then
        OpenAndGetData = Array("")
        Exit Function
    End If
    
    status = IsOpen(FilePath)
    
    FastMode
    
    'Set workbook with/without password
    If IsMissing(Password) = True Then
        Set WBdata = Workbooks.Open(FilePath, , True)
    Else
        Set WBdata = Workbooks.Open(FilePath, , True, , Password)
    End If
    
    'Set worksheet with/without name
    If IsMissing(SheetName) = True Then
        Set WSdata = WBdata.Sheets(1)
    Else
        Set WSdata = WBdata.Sheets(SheetName)
    End If
    
    OpenAndGetData = WSdata.UsedRange
    
    'If the workbook was not open at the beginning, close it
    If status = False Then WBdata.Close
    
    Set WBdata = Nothing
    Set WSdata = Nothing
    
    NormalMode

End Function

Public Sub WriteArray(WorkBookPath As String, DataArray As Variant, Optional SheetName As Variant, Optional InitialRow As Variant, Optional InitialColumn As Variant)
    
    Dim main As Workbook
    Dim sheet As String
    Dim r As Long, c As Long, cMemo As Long
    Dim i As Long, j As Long
    
    If IsOpen(File.Name(WorkBookPath)) = True Then
        Set main = GetObject(WorkBookPath)
    Else
        Set main = Workbooks.Open(WorkBookPath)
    End If
    
    If IsMissing(InitialRow) = True Then
        r = 1
    Else
        r = InitialRow
    End If
    
    If IsMissing(InitialColumn) = True Then
        c = 1
    Else
        c = InitialColumn
    End If
    
    If IsMissing(SheetName) = True Then
        sheet = "Sheet1"
    Else
        sheet = SheetName
    End If
    
    cMemo = c
    
    For i = 1 To UBound(DataArray, 1)
        c = cMemo
        For j = 1 To UBound(DataArray, 2)
            main.Sheets(sheet).Cells(r, c) = DataArray(i, j)
            c = c + 1
        Next j
        r = r + 1
    Next i
End Sub

Public Function GetUsedRange(Path As String, Optional SheetName As Variant, Optional IncludeHeaders As Boolean = True) As Variant
    
    Dim rs As Object 'Object 'ADODB.Recordset
    Dim strConn As String
    Dim strSQL As String
    Dim data As Variant
    
    Set rs = CreateObject("ADODB.Recordset")        'CreateObject("ADODB.Recordset")    New ADODB.Recordset
    
    If IncludeHeaders = True Then
        strConn = XLConnectionString(Path) + ";HDR=No"""
    Else
        strConn = XLConnectionString(Path) + ";HDR=Yes"""
    End If
    
    If IsMissing(SheetName) = True Then
        strSQL = "Select * from [Sheet1$]"
    Else
        strSQL = "Select * from [" & SheetName & "$]"
    End If
    
    rs.Open strSQL, strConn, 3, 1   'adOpenStatic, adLockReadOnly -> 3  and  1
    
    data = rs.GetRows(, 1)

    GetUsedRange = ConvertArray(data)
    
    Set rs = Nothing
    
End Function

Private Function ConvertArray(DataArray As Variant) As Variant
    
    Dim i As Long, j As Long
    Dim data() As Variant
    Dim LastRow As Long, LastColumn As Long
    
    LastRow = UBound(DataArray, 2)
    LastColumn = UBound(DataArray, 1)
    
    ReDim data(1 To LastRow + 1, 1 To LastColumn + 1)
    
    For i = 0 To LastRow
        For j = 0 To LastColumn

            If TypeName(DataArray(j, i)) = "Null" Then
                data(i + 1, j + 1) = vbNullString
            Else
                data(i + 1, j + 1) = DataArray(j, i)
            End If
            
        Next j
    Next i
    
    ConvertArray = data
    
End Function

Public Function IsOpen(WBNameOrPath) As Boolean
'For parameter, you need to enter only the name of the workbook, not the full path

    Dim a As Workbook
    On Error Resume Next
    
    If File.IsPath(WBNameOrPath) = False Then
        Set a = Workbooks(WBNameOrPath)
    Else
        Set a = Workbooks(File.Name(WBNameOrPath))
    End If
    
    If Err.Number = 0 Then
        
        IsOpen = True
        
    Else
    
        IsOpen = False
        
    End If

End Function

Public Sub ExportCSV(CSVFileName, ExportRange As Range)
    
    Dim ExpBook As Workbook
    ' Exports a range to CSV file
    If ExportRange Is Nothing Then
    Debug.Print "Not ExportRange"
    Exit Sub
    End If
    
    On Error GoTo ErrHandle
    
    Application.ScreenUpdating = False
    
    Set ExpBook = Workbooks.Add(xlWorksheet)
    
    ExportRange.Copy
    Application.DisplayAlerts = False
    
    With ExpBook
    .Sheets(1).Paste
    .SaveAs Filename:=CSVFileName, FileFormat:=xlCSV
    .Close SaveChanges:=False
    End With
    
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    Exit Sub
ErrHandle:
    ExpBook.Close SaveChanges:=False
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Debug.Print "Error " & Err & vbCrLf & vbCrLf & Error(Err), _
    vbCritical, "Export Method Error"
End Sub

Public Sub ImportCSV(CSVFileName, ImportRange As Range)
    
    Dim CSVFile As Workbook
    ' Imports a CSV file to a range
    If ImportRange Is Nothing Then
    Debug.Print "Not ImportRange"
    Exit Sub
    End If
    
    If CSVFileName = "" Then
    Debug.Print "Import FileName not specified"
    Exit Sub
    End If
    
    On Error GoTo ErrHandle
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Workbooks.Open CSVFileName
    
    Set CSVFile = ActiveWorkbook
    
    ActiveSheet.UsedRange.Copy Destination:=ImportRange
    CSVFile.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub

ErrHandle:
    CSVFile.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Debug.Print "Error " & Err & vbCrLf & vbCrLf & Error(Err), _
    vbCritical, "Import Method Error"
End Sub

Private Sub FastMode()

    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With

End Sub

Private Sub NormalMode()
    
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
    End With

End Sub

