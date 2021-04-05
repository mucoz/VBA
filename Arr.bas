Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                            Created by Mustafa Can Ozturk on 18.03.2021                           '
'                                                                                                  '
'                      The functions below will be used for array operations                       '
'                                It needs to be used in "Arr" module                               '
'                                                                                                  '
'       Functions : WriteArray, DeleteColumnValues, LookUp, Print2D, Print1D, BubbleSort            '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub WriteArray(Arr As Variant, Rng As Range)

    Dim Destination As Range
    Set Destination = Rng
    Destination.Resize(UBound(Arr, 1), UBound(Arr, 2)).value = Arr

End Sub

Public Function DeleteColumnValues(DataArray As Variant, ColumnNumber As Long, value As Variant) As Variant

    Dim Data() As Variant
    Dim i As Long, j As Long
    Dim m As Long, n As Long
    Dim counter As Long
    
    If IsArray(DataArray) = False Then Exit Function
    
    For i = LBound(DataArray, 1) To UBound(DataArray, 1)
    
        If Trim(DataArray(i, ColumnNumber)) = value Then
            
            counter = counter + 1
        
        End If
    
    Next i
    
    If counter = 0 Then
    
        Debug.Print "There is no """ & CStr(value) & """ in the column " & CStr(ColumnNumber)
        Exit Function
    
    Else
    
        ReDim Data(1 To UBound(DataArray, 1) - counter, 1 To UBound(DataArray, 2))
        
    End If
    
    m = 1
    
    For i = LBound(DataArray, 1) To UBound(DataArray, 1)
        
        n = 1
        
        For j = LBound(DataArray, 2) To UBound(DataArray, 2)
        
            If Trim(CStr(DataArray(i, ColumnNumber))) <> CStr(value) Then
            
                Data(m, n) = DataArray(i, j)
                n = n + 1
            
            End If
            
        Next j
        
        If Trim(CStr(DataArray(i, ColumnNumber))) <> CStr(value) Then
            m = m + 1
        End If
        
    Next i
    
    DeleteColumnValues = Data
    
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


Public Function BubbleSort(DataArray As Variant) As Variant
    
    Dim arr As Variant
    Dim first As Long, last As Long
    Dim i As Long, j As Long, temp As Long
    
    arr = DataArray
    first = LBound(DataArray)
    last = UBound(DataArray)
    
    For i = first To last - 1
        For j = i + 1 To last
            If arr(i) > arr(j) Then
                temp = arr(j)
                arr(j) = arr(i)
                arr(i) = temp
            End If
        Next j
    Next i
    
    BubbleSort = arr
    
End Function
