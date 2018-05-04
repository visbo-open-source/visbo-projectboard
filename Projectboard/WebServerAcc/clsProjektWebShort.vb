Public Class clsProjektWebShort
    Public Property _id As String
    Public Property name As String
    Public Property updatedAt As Date
    Public Property createdAt As Date
    Public Property vpid As String
    Public Property timestamp As Date
    Public Property Erloes As Integer
    Public Property endDate As Date
    Public Property variantName As String

    Public Sub New()
        _id = ""
        name = "Project Name"
        vpid = ""
        timestamp = Date.MinValue
        Erloes = 0
        endDate = Date.MinValue
        variantName = ""
    End Sub
End Class
