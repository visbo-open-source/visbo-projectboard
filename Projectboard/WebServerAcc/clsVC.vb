Public Class clsVC
    Public Property _id As String
    Public Property name As String
    Public Property users As List(Of clsVCuser)
    Public Property updatedAt As String
    Public Property createdAt As String

    Sub New()
        _id = ""
        _name = "not named"
        _users = New List(Of clsVCuser)
        _updatedAt = Date.MinValue.ToString
        _createdAt = Date.MinValue.ToString
    End Sub

End Class
