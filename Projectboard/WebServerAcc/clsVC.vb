Public Class clsVC
    Public Property _id As String
    Public Property Name As String
    Public Property Users As List(Of clsVCuser)
    Sub New()
        _id = ""
        _Name = "not named"
        _Users = New List(Of clsVCuser)
    End Sub
End Class
