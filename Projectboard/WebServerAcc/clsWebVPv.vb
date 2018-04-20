Imports ProjectBoardDefinitions

Public Class clsWebVPv

    Public Property state As String
    Public Property message As String
    Public Property vpv As List(Of clsProjektWeb)

    Sub New()
        _state = "unknown"
        _message = "not yet any response"
        _vpv = New List(Of clsProjektWeb)
    End Sub
End Class
