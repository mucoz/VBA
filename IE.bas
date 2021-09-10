Option Explicit

Private Declare Function InternetCheckConnection Lib "wininet.dll" Alias "InternetCheckConnectionA" (ByVal lpszUrl As String, ByVal dwFlags As Long, ByVal dwReserved As Long) As Long

Const FLAG_ICC_FORCE_CONNECTION = &H1

Public Function IsConnected() As Boolean

    Dim msg As Boolean
    Dim status As Boolean
    Dim address As String
    
    address = "http://www.yahoo.com/"
    
    status = InternetCheckConnection(address, FLAG_ICC_FORCE_CONNECTION, 0&)
    
    If status Then
        
        msg = True
    
    Else
        
        msg = False
        
    End If
    
    IsConnected = msg
    
End Function


Public Function isVPNConnected() As Boolean

Dim PingResults As Object
Dim PingResult As Variant
Dim Query As String
Dim Ping As Boolean

Dim Host As String


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                  In order to find Host Name, you can use "ipconfig /all" command                 '
'                 There might be more than one name in cmd. Try one by one to check                '
'      You can also open Control Panel -> Network and Sharing Center to see the Domain network     '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Host = "hitachivantara.com"

Query = "SELECT * FROM Win32_PingStatus WHERE Address = '" & Host & "'"

Set PingResults = GetObject("winmgmts://./root/cimv2").ExecQuery(Query)

For Each PingResult In PingResults
    If Not IsObject(PingResult) Then
        Ping = False
    ElseIf PingResult.StatusCode = 0 Then
        Ping = True
    Else
        Ping = False
    End If
Next

    isVPNConnected = Ping
    
End Function
    
    
    
    
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'      Purpose : The codes below control Internet Explorer. You need to include Internet Controls library      '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''    
' pauses the execution of your VBA code while waiting for IE to finish loading
Sub waitFor(IE As InternetExplorer)
    Do
        Do
            Application.Wait Now + TimeValue("00:00:01")
            attach IE
            DoEvents
        Loop Until Not IE.Busy And IE.readystate = 4
        Application.Wait Now + TimeValue("00:00:01")
    Loop Until Not IE.Busy And IE.readystate = 4
End Sub



'Connect IE to the most recently opened Internet Explorer windows
' if urlPart is supplied, will only attach on to an explorer that has that string as a part of the URL
Function attach(IE As Object, Optional urlPart As String) As Boolean
    Dim o As Object
    Dim x As Long
    Dim explorers As Object
    Dim name As String
    Set explorers = CreateObject("Shell.application").Windows
    For x = explorers.Count - 1 To 0 Step -1
       name = "Empty"
       On Error Resume Next
       name = explorers((x)).name
       On Error GoTo 0
       If name = "Internet Explorer" Then
          If InStr(1, explorers((x)).LocationURL, urlPart, vbTextCompare) Then
               Set IE = explorers((x))
               attach = True
               Exit For
          End If
       End If
    Next
    
End Function

' Returns the number of the HTML element specified by tagname and identifying text
Function getTagNumber(IE As InternetExplorer, tagName As String, Optional identifyingText As String, Optional startAtTagNumber As Long = 0) As Long
  Dim x As Long
  Dim t As Object
  For x = startAtTagNumber To IE.Document.all.Length - 1
     Set t = IE.Document.all(x)
     If UCase(t.tagName) = UCase(tagName) Then
        'we found the right kind of tag, check to see if it has the right text
        'Debug.Print t.outerHTML
        If InStr(1, t.outerhtml, identifyingText) > 0 Then
          'we found the right kind fo tag with the right identifying text, return the number
          getTagNumber = x
          Exit Function
        End If
     End If
  Next
  getTagNumber = -1 ' sentinal value indicating the tag was not found
End Function

'returns a reference to a tag object given its number.  Used in conjunction with GetTagNumber
Function getTag(IE As InternetExplorer, tagNumber) As Object
  Set getTag = IE.Document.all(tagNumber)
End Function

Public Sub showpage(IE As InternetExplorer)
    savePage IE
    ThisWorkbook.FollowHyperlink ThisWorkbook.Path & "\source.html"
End Sub


Public Sub savePage(IE As InternetExplorer, Optional filePath As String)
  'saves a local copy of the document in Internet Explorer as currently rendered
  Dim x As Long
  Dim len1 As Long
  Dim len2 As Long
  
  Dim ff As Integer
  ff = FreeFile
    If filePath = "" Then
       Open ThisWorkbook.Path & "\source.html" For Output As #ff
   Else
       Open filePath For Output As #ff
   End If
   
   For x = 0 To IE.Document.all.Length - 1
     Print #ff, IE.Document.all(x).outerhtml
     If UCase(IE.Document.all(x).tagName) = "HTML" Then Exit For
   Next
   
   Close #ff
End Sub

'Uses the WebQuery Wizard to import data from the current page in IE
Public Sub importPage(IE As InternetExplorer, newSheetName As String, Optional wb As Workbook)
  Dim ff As Integer
  Dim S As Worksheet
  
  If TypeName(wb) = "Nothing" Then
    Set wb = ThisWorkbook
  End If
  
  ff = FreeFile
  
   Open ThisWorkbook.Path & "\localWebPageAgentFile.html" For Output As #ff
   Print #ff, "<html><head><title>Saved Page</title></head>"
   Print #ff, IE.Document.body.outerhtml
   Print #ff, "</html>"
   Close #ff

  
  
  On Error Resume Next
    Application.DisplayAlerts = False
       wb.Sheets(newSheetName).Delete
    Application.DisplayAlerts = True
  On Error GoTo 0
  
  
  Set S = wb.Worksheets.Add
  S.name = newSheetName
  
      With S.QueryTables.Add(Connection:= _
        "URL;file:///" & Replace(ThisWorkbook.Path, "\", "/") & "/localWebPageAgentFile.html", Destination:=S.Range("$A$1"))
        .name = "localWebPageAgentFile"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .WebSelectionType = xlEntirePage
        .WebFormatting = xlWebFormattingNone
        .WebPreFormattedTextToColumns = True
        .WebConsecutiveDelimitersAsOne = True
        .WebSingleBlockTextImport = False
        .WebDisableDateRecognition = False
        .WebDisableRedirections = False
        .Refresh BackgroundQuery:=False
    End With
    
    S.QueryTables(1).Delete
    
  Kill ThisWorkbook.Path & "\localWebPageAgentFile.html"
  
End Sub

