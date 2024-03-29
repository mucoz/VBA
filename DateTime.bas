Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                            Created by Mustafa Can Ozturk on 18.03.2021                           '
'                                                                                                  '
'                       The functions below will be used for date operations                       '
'                               It needs to be used in "DateTime" module                           '
'                                                                                                  '
'                Functions : FormatDateInArray, MonthToNumber, DateNow, NameOfMonth                '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Enum Language
    English
    Spanish
    German
End Enum

Public Sub FormatDateInArray(ByRef DataArray As Variant, ColumnNumberOfDates As Integer)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                  This function changes the date format in a column of an array                   '
'    It takes 2 arguments such as a 2D array and the number of the column that holds the dates     '
'                                                                                                  '
'                        e.g. Consider a 2D array with 3 rows and 2 columns                        '
'                                                                                                  '
'                                  Call FormatDateInArray(Array, 2)                                '
'                                                                                                  '
'                                              Array                                               '
'                        --------------------------------------------------                        '
'                                                                                                  '
'                            1     11-NOV-20        ->        11.11.2020                           '
'                            2     13-JAN-20        ->        13.01.2020                           '
'                            3     15-FEB-20        ->        15.02.2020                           '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim i As Long
    
    For i = LBound(DataArray) To UBound(DataArray)
        
        DataArray(i, ColumnNumberOfDates) = MonthToNumber(DataArray(i, ColumnNumberOfDates))
    
    Next i
    
End Sub

Public Function MonthToNumber(value As Variant) As Variant
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'            This function is used to convert "dd-MMM-yy" format to "dd.mm.yyyy" format            '
'                                                                                                  '
'                                  e.g.  13-OCT-20 ->  13.10.2020                                  '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim months As Variant
    Dim newMonth As String
    Dim i As Long
    Dim converted As Boolean
    
    converted = False
    
    months = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")

    For i = LBound(months) To UBound(months)
    
        If InStr(UCase(value), UCase(months(i))) > 0 Then
        
            If Val(months(i)) < 10 Then
                newMonth = Left(value, 2) & CStr(".0" & (i + 1) & ".") & Right(value, 2)
            Else
                newMonth = Left(value, 2) & CStr("." & (i + 1) & ".") & Right(value, 2)
            End If
            
            converted = True
            Exit For
            
        End If
        
    Next i
    
    If converted = True Then
    MonthToNumber = CDate(newMonth)
    Else
    MonthToNumber = value
    End If
End Function

Public Function DateNow() As Date

    DateNow = Format(Now, "dd.mm.yyyy")

End Function

Public Function NameOfMonth(ByVal FullDate As Variant, ResultLanguage As Language) As String
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'       Give this function a full date (e.g. "13.10.2021") and the language (e.g. "English")       '
'              It will give you the name of the month (e.g. October) in that language              '
'                                                                                                  '
'                                "13.10.2021"     ->     "October"                                 '
'                                                                                                  '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim EnglishArray As Variant
    Dim SpanishArray As Variant
    Dim GermanArray As Variant
    Dim DataArray As Variant
    Dim i As Long
    Dim d As Date
    
    On Error GoTo exitF
    d = CDate(FullDate)
    
    EnglishArray = Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
    SpanishArray = Array("Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre")
    GermanArray = Array("Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember")
    
    If ResultLanguage = English Then
        DataArray = EnglishArray
    ElseIf ResultLanguage = Spanish Then
        DataArray = SpanishArray
    ElseIf ResultLanguage = German Then
        DataArray = GermanArray
    End If
    
    For i = LBound(DataArray) To UBound(DataArray)
    
        If Month(d) = i + 1 Then
        
            NameOfMonth = DataArray(i)
            Exit Function
            
        End If
        
    Next i

exitF:
    NameOfMonth = FullDate

End Function

