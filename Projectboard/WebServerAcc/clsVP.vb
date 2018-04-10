Public Class clsVP
    Public Property _id As String
    Public Property name As String
    Public Property vcid As String
    Public Property users As List(Of clsUser)
    Public Property updatedAt As String
    Public Property createdAt As String

    Sub New()
        _id = ""
        _name = "not named"
        _vcid = "not yet defined"
        _users = New List(Of clsUser)
        _updatedAt = Date.MinValue.ToString
        _createdAt = Date.MinValue.ToString
    End Sub
End Class
