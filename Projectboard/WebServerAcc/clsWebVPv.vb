Imports ProjectBoardDefinitions

Public Class clsWebVPv

    Inherits clsWebOutput
    Public Property vpv As List(Of clsProjektWebShort)

    Sub New()
        _vpv = New List(Of clsProjektWebShort)
    End Sub
End Class
