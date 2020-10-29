Public Class clsVPvariant
    Public Property _id As String
    Public Property variantName As String
    Public Property email As String
    Public Property createdAt As Date
    Public Property vpvCount As Integer


    Sub New()
        _id = ""
        _variantName = ""
        _email = "someone@visbo.de"
        _createdAt = Date.MinValue
        _vpvCount = 0
    End Sub
End Class
