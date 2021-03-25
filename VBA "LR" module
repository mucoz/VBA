Option Explicit

Public Function OfArray(DataArray As Variant) As Long

    OfArray = UBound(DataArray, 1)

End Function

Public Function Active() As Long

    Active = Split(ActiveSheet.UsedRange.Address, "$")(4)
    
End Function

Public Function UsedRange(Sheet As Worksheet) As Long

    UsedRange = Split(Sheet.UsedRange.Address, "$")(4)

End Function

Public Function LastRowOfColumn(Sheet As Worksheet, ColumnnNameOrNumber As Variant) As Long
    
    If IsNumeric(ColumnnNameOrNumber) = True Then
    
        LastRowOfColumn = Sheet.Cells(Rows.Count, ColumnnNameOrNumber).End(xlUp).Row
    
    Else
    
        LastRowOfColumn = Sheet.Range(ColumnnNameOrNumber & Rows.Count).End(xlUp).Row
    
    End If
    
End Function
