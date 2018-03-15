Public Class clsVCplus
    Inherits clsVC
    Public Property _v As Integer
    Public Property updatedAt As String
    Public Property createdAt As String

    Sub New()
        _v = 0
        _updatedAt = Date.MinValue.ToString
        _createdAt = Date.MinValue.ToString
    End Sub


End Class
