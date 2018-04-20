Imports ProjectBoardDefinitions
Public Class clsHrchyNodeWeb
    Public Property elemId As String
    Public Property node As clsHierarchyNodeDB

    Sub New()
        _elemId = ""
        _node = New clsHierarchyNodeDB
    End Sub
End Class
