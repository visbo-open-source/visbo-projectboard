Public Class clsKostenartDefinition

    Public name As String
    Public farbe As Object

    Private _budget() As Double
    Private _uuid As Integer

    Public Property UID() As Integer
        Get
            UID = _uuid
        End Get
        Set(value As Integer)
            _uuid = value
        End Set
    End Property

    Public Sub New()

    End Sub
End Class
