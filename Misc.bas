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

'Taking picture

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
