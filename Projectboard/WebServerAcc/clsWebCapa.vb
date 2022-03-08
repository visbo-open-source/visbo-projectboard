Imports ProjectBoardDefinitions

Public Class clsWebCapa

    Inherits clsWebOutput
    Public Property count As Integer
    Public Property capacity As List(Of clsCapa)

    Sub New()
        _count = 0
        _capacity = New List(Of clsCapa)
    End Sub
End Class
