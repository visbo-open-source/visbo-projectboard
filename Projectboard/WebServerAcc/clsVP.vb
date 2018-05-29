Public Class clsVP
    Public Property _id As String
    Public Property name As String
    Public Property vcid As String
    Public Property users As List(Of clsUser)
    Public Property updatedAt As String
    Public Property createdAt As String
    Public Property lock As List(Of clsVPLock)
    Public Property [Variant] As List(Of clsVPvariant)

    Sub New()
        _id = ""
        _name = "not named"
        _vcid = "not yet defined"
        _users = New List(Of clsUser)
        _updatedAt = Date.MinValue.ToString
        _createdAt = Date.MinValue.ToString
        _lock = New List(Of clsVPLock)
        _Variant = New List(Of clsVPvariant)
    End Sub
End Class
