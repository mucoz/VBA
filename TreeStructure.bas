Option Explicit

Public Sub TreeStructure()
    
    Dim t As New Tree
    Call t.SetRoot("Root")
    Call t.AddNode("Root", "Child 1")
    Call t.AddNode("Root", "Child 2")
    Call t.AddNode("Child 1", "Child 1.1")
    Call t.AddNode("Child 1", "Child 1.2")
    Call t.AddNode("Child 1", "Child 1.3")
    Call t.AddNode("Child 2", "Child 2.1")
    Call t.AddNode("Child 2", "Child 2.2")
    Call t.Display
    
End Sub

'"Node class"
'
'Option Explicit
'
'Private m_value As Variant
'Private m_children As Collection
'
'Private Sub Class_Initialize()
'    Set m_children = New Collection
'End Sub
'
'Public Sub SetValue(Value As Variant)
'    m_value = Value
'End Sub
'
'Public Function GetValue() As String
'    GetValue = m_value
'End Function
'
'Public Sub AddChild(ChildNode As Variant)
'    Call m_children.Add(ChildNode)
'End Sub
'
'Public Function GetChildren() As Collection
'    Set GetChildren = m_children
'End Function


'"Tree class"
'
'
'Option Explicit
'
'Private m_root As Node
'
'Public Sub SetRoot(RootValue As Variant)
'    Set m_root = New Node
'    Call m_root.SetValue(RootValue)
'End Sub
'
'Public Sub AddNode(ParentValue As Variant, ChildValue As Variant)
'    Dim parentNode As Node
'    Dim nodeObject As Node
'    Set parentNode = FindNode(m_root, ParentValue)
'    If Not parentNode Is Nothing Then
'        Set nodeObject = New Node
'        Call nodeObject.SetValue(ChildValue)
'        Call parentNode.AddChild(nodeObject)
'    Else
'        Debug.Print "Parent node with value '" + ParentValue + "' not found"
'    End If
'End Sub
'
'Public Function FindNode(ByRef nodeObject As Node, Value As Variant) As Variant
'    Dim child As Node
'    Dim result As Variant
'    If nodeObject.GetValue = Value Then
'        Set FindNode = nodeObject
'        Exit Function
'    End If
'    For Each child In nodeObject.GetChildren
'        Set result = FindNode(child, Value)
'        If Not result Is Nothing Then
'            Set FindNode = result
'            Exit Function
'        End If
'    Next child
'    Set FindNode = Nothing
'End Function
'
'Public Sub Display(Optional nodeObject As Node, Optional Level As Integer = 0)
'    Dim child As Node
'    If nodeObject Is Nothing Then
'        Set nodeObject = m_root
'    End If
'    Debug.Print String(Level, vbTab) + "-" + nodeObject.GetValue
'    For Each child In nodeObject.GetChildren
'        Call Display(child, Level + 1)
'    Next child
'End Sub

