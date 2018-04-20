Public Class clsVC
    Public Property _id As String
    Public Property name As String
    Public Property users As List(Of clsUser)
    'Public Property updatedAt As String
    Public Property createdAt As String

    Sub New()
        _id = ""
        _name = "not named"
        _users = New List(Of clsUser)
        '_updatedAt = Date.MinValue.ToString
        _createdAt = Date.MinValue.ToString
    End Sub

End Class
