Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                            Created by Mustafa Can Ozturk on 18.03.2021                           '
'                                                                                                  '
'                      The functions below will be used for array operations                       '
'                                It needs to be used in "Arr" module                               '
'                                                                                                  '
'       Functions : WriteArray, DeleteColumnValues, LookUp, Print2D, Print1D, BubbleSort           '
'           QuickSort, 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Enum ArrayValue
    
    Include
    Exclude

End Enum


Public Function DeleteColumnValues(DataArray As Variant, ColumnNumber As Long, value As Variant, IncludeOrExclude As ArrayValue, Optional IsThereHeader As Boolean = True) As Variant

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'            This function deletes the rows of an array that contain a value in a column           '
'                                                                                                  '
'                        e.g.      Call DeleteColumnValues(Arr, 3, "Ipad")                         '
'              Function will delete the 2nd row as it contains "Ipad" in 3rd column                '
'                                          Column Numbers                                          '
'                                                                                                  '
'                    1     2         3         4    ->  1     2       3         4                  '
'                    ----------------------------      ---------------------------                 '
'                    1   Apple   Microsoft     20   ->  1   Apple  Microsoft    20                 '
'                    2   Apple   Ipad          56   ->  3   Apple   Iphone      13                 '
'                    3   Apple   Iphone        13   ->  4   Apple   Icar        34                 '
'                    4   Apple   Icar          34   ->                                             '
'                                                                                                  '
'                                                                                                  '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim data() As Variant
    Dim i As Long, j As Long
    Dim m As Long
    Dim counter As Long
    Dim increaseM As Boolean
    
    If IsArray(DataArray) = False Then
    
        DeleteColumnValues = "VALUE NOT FOUND"
        Exit Function
    
    End If
    
    For i = LBound(DataArray, 1) To UBound(DataArray, 1)
    
        If UCase(Trim(DataArray(i, ColumnNumber))) = UCase(Trim(value)) Then
            
            counter = counter + 1
            
        End If
    
    Next i
    
    If counter = 0 Then
        
        If IncludeOrExclude = Exclude Then
            DeleteColumnValues = DataArray 'Debug.Print "There is no """ & CStr(value) & """ in the column " & CStr(ColumnNumber)
            Exit Function
        Else
            DeleteColumnValues = "VALUE NOT FOUND"
            Exit Function
        End If
    
    Else
        If IncludeOrExclude = Exclude Then
        
            ReDim data(1 To UBound(DataArray, 1) - counter, 1 To UBound(DataArray, 2))
            m = 1
        
        Else
        
            If IsThereHeader = True Then
                
                ReDim data(1 To counter + 1, 1 To UBound(DataArray, 2))
                m = 2
            
            Else
            
                ReDim data(1 To counter, 1 To UBound(DataArray, 2))
                m = 1
                
            End If
        
        End If
        
    End If

    increaseM = False
    
    For i = m To UBound(DataArray, 1)
        
        For j = LBound(DataArray, 2) To UBound(DataArray, 2)
        
            If IncludeOrExclude = Include Then
                If Trim(CStr(DataArray(i, ColumnNumber))) = CStr(value) Then
                
                    data(m, j) = DataArray(i, j)
                    increaseM = True
                
                End If
            Else
                If Trim(CStr(DataArray(i, ColumnNumber))) <> CStr(value) Then
                
                    data(m, j) = DataArray(i, j)
                    increaseM = True
                
                End If
            End If
            
        Next j
        
        If increaseM = True Then
            m = m + 1
            increaseM = False
        End If
        
    Next i
    
    'if there is header, fill it
    If IsThereHeader = True Then
        For i = LBound(DataArray, 2) To UBound(DataArray, 2)
            
            data(1, i) = DataArray(1, i)
        
        Next i
    End If
    
    DeleteColumnValues = data
    
End Function

Public Sub WriteArray(Arr As Variant, Rng As Range)

    Dim Destination As Range
    Set Destination = Rng
    Destination.Resize(UBound(Arr, 1), UBound(Arr, 2)).value = Arr

End Sub

Public Function ClearHeaders(DataArray As Variant) As Variant

    Dim i As Long, j As Long
    Dim data() As Variant
    Dim lower As Long, upper As Long
    Dim m As Long
    
    If IsArray(DataArray) = False Then
        'If it s not an array, return the same array
        ClearHeaders = DataArray
        Exit Function
    End If
    
    lower = LBound(DataArray, 1)
    upper = UBound(DataArray, 1)
    
    If upper <= 1 Then
        'If there is only one row, return the same array
        ClearHeaders = DataArray
        Exit Function
    End If
    
    ReDim data(lower To upper - 1, LBound(DataArray, 2) To UBound(DataArray, 2))
    
    m = 1

    For i = lower + 1 To upper
        For j = LBound(DataArray, 2) To UBound(DataArray, 2)
        
            data(m, j) = DataArray(i, j)
            
        Next j
        m = m + 1
    Next i
    
    ClearHeaders = data
    
End Function
    
Public Function LookUp(ByVal SearchItem As Variant, LookUpRange As Variant, FirstColumnNumber As Long, LastColumnNumber As Long) As Variant
    
    Dim i As Long
    
    For i = 1 To UBound(LookUpRange, 1)
        
        If IsNumeric(LookUpRange(i, FirstColumnNumber)) = True Then
            
            If SearchItem = Val(LookUpRange(i, FirstColumnNumber)) Then
                LookUp = LookUpRange(i, LastColumnNumber)
                Exit Function
            End If
        
        Else
            
            If SearchItem = CStr(LookUpRange(i, FirstColumnNumber)) Then
                LookUp = LookUpRange(i, LastColumnNumber)
                Exit Function
            End If
        
        End If
    
    Next i
    
    LookUp = "NOT FOUND"

End Function

Public Function LookUp2(ByVal FirstSearchItem As Variant, _
                        ByVal SecondSearchItem As Variant, _
                        ByRef LookUpRange As Variant, _
                        FirstItemColumnNumber As Long, _
                        SecondItemColumnNumber As Long, _
                        GetDataFromColumnNumber) As Variant
    
    Dim i As Long
    
    If FirstSearchItem <> "" Or SecondSearchItem <> "" Then
        
        For i = 1 To UBound(LookUpRange, 1)
            
            If CStr(LookUpRange(i, FirstItemColumnNumber)) = CStr(FirstSearchItem) And CStr(LookUpRange(i, SecondItemColumnNumber)) = CStr(SecondSearchItem) Then
                LookUp2 = LookUpRange(i, GetDataFromColumnNumber)
                Exit Function
            End If
            
        Next i
        
    Else
        
        LookUp2 = ""
        Exit Function
    
    End If
    
    LookUp2 = ""

End Function
        
Public Function FindTextInColumn(arr As Variant, ColumnnNameOrNumber As Variant, Text As Variant) As String

    Dim i As Long
    Dim colN As Long
    
    If IsNumeric(ColumnnNameOrNumber) = True Then
        colN = ColumnnNameOrNumber
    Else
        colN = Range(ColumnnNameOrNumber & 1).Column
    End If
    
    For i = LBound(arr) To UBound(arr)
        
        If InStr(UCase(arr(i, colN)), UCase(Text)) > 0 Then
            
            FindTextInColumn = arr(i, colN)
            Exit Function
        
        End If
    
    Next i
    
    FindTextInColumn = "VALUE NOT FOUND"    '""
    
End Function        
        
Public Sub Print2D(DataArray As Variant)

    Dim i As Long, j As Long
    Dim line As String
    
    For i = LBound(DataArray, 1) To UBound(DataArray, 1)
        line = ""
        For j = LBound(DataArray, 2) To UBound(DataArray, 2)
            line = line + CStr(DataArray(i, j)) + "       "
        Next j
        Debug.Print line
    Next i
    
End Sub

Public Sub Print1D(DataArray As Variant)
    
    Dim i As Long
    For i = LBound(DataArray, 1) To UBound(DataArray, 1)
        Debug.Print CStr(DataArray(i))
    Next i
    
End Sub

Public Function SumOfColumn(DataArray As Variant, ColumnNumber As Long) As Variant

    Dim sum As Variant
    Dim i As Long
    
    sum = 0
    
    For i = LBound(DataArray, 1) To UBound(DataArray, 1)
    
        If IsNumeric(DataArray(i, ColumnNumber)) = True Then
            sum = sum + DataArray(i, ColumnNumber)
        End If
        
    Next i
    
    SumOfColumn = sum
    
End Function

Public Function GetHeaderColumn(DataArray As Variant, HeaderText As String, Optional HeaderText2 As Variant) As Long
    
    Dim i As Long, j As Long
    
    If IsMissing(HeaderText2) = True Then
        For i = LBound(DataArray, 1) To UBound(DataArray, 1)
            For j = LBound(DataArray, 2) To UBound(DataArray, 2)
                If InStr(UCase(CStr(DataArray(i, j))), UCase(CStr(Trim(HeaderText)))) > 0 Then
                    GetHeaderColumn = j
                    Exit Function
                End If
            Next j
        Next i
    Else
        For i = LBound(DataArray, 1) To UBound(DataArray, 1)
            For j = LBound(DataArray, 2) To UBound(DataArray, 2)
                If InStr(UCase(CStr(DataArray(i, j))), UCase(CStr(Trim(HeaderText)))) > 0 And _
                   InStr(UCase(CStr(DataArray(i, j))), UCase(CStr(Trim(HeaderText2)))) > 0 Then
                    GetHeaderColumn = j
                    Exit Function
                End If
            Next j
        Next i
    End If
    
    GetHeaderColumn = -1
    
End Function

Public Function GetDimension(DataArray As Variant) As Integer
    
    Dim dimNumber As Integer
    Dim result As Integer
    
    On Error Resume Next
    
    Do
        dimNumber = dimNumber + 1
        result = UBound(DataArray, dimNumber)
    Loop Until Err.Number <> 0
    
    Err.Clear
    
    GetDimension = dimNumber - 1

End Function



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                An array with 100.000 elements was sorted in               '
'               403,89 seconds by using Bubble Sort algorithm               '
'                00,53 seconds by using Quick Sort algorithm                '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Public Sub BubbleSort(ByRef DataArray As Variant)
    
    Dim first As Long, last As Long
    Dim i As Long, j As Long, temp As Long
    
    first = LBound(DataArray)
    last = UBound(DataArray)
    
    For i = first To last - 1
        For j = i + 1 To last
            If DataArray(i) > DataArray(j) Then
                temp = DataArray(j)
                DataArray(j) = DataArray(i)
                DataArray(i) = temp
            End If
        Next j
    Next i
    
End Sub

Public Sub QuickSort(ByRef DataArray As Variant)

    Dim l As Long, u As Long
    Dim arr As Variant
    
    l = LBound(DataArray)
    u = UBound(DataArray)
    Call Util(DataArray, l, u)
    
End Sub

Private Sub Util(ByRef DataArray As Variant, lower As Long, upper As Long)

    If upper <= lower Then
        Exit Sub
    End If
    
    Dim pivot As Long, start As Long, finish As Long
    
    pivot = DataArray(lower)
    start = lower
    finish = upper
    
    Do While lower < upper
        Do While (DataArray(lower) <= pivot And lower < upper)
            
            lower = lower + 1
        
        Loop
    
        Do While (DataArray(upper) > pivot And lower <= upper)
        
            upper = upper - 1
            
        Loop
        
        If lower < upper Then
            
            Swap DataArray, upper, lower
        
        End If
    Loop
    
    Swap DataArray, upper, start
    
    Util DataArray, start, upper - 1
    Util DataArray, upper + 1, finish
    
End Sub

Private Sub Swap(ByRef DataArray As Variant, first As Long, second As Long)
    
    Dim temp As Long
    
    temp = DataArray(first)
    
    DataArray(first) = DataArray(second)
    DataArray(second) = temp
    
End Sub

'===============================END OF QUICK SORT============================
