Public Class clsVC
    Public Property _id As String
    Public Property name As String
    'Public Property users As List(Of clsUser)
    Public Property updatedAt As Date
    Public Property createdAt As Date

    Sub New()
        _id = ""
        _name = "not named"
        '_users = New List(Of clsUser)
        _updatedAt = Date.MinValue
        _createdAt = Date.MinValue
    End Sub

End Class
