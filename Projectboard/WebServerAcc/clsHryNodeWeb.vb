Imports ProjectBoardDefinitions
Public Class clsHryNodeWeb
    Public Property hryNodeKey As String
    Public Property hryNode As clsHierarchyNodeDB
    Public Sub New()
        _hryNodeKey = ""
        _hryNode = New clsHierarchyNodeDB
    End Sub
End Class
