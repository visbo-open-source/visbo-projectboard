Public Class clsWebVCTSOrganisation
    Inherits clsWebOutput
    Public Property organisation As List(Of clsTSOOrganisationWeb)
    Public Property updatedAt As String
    Public Property createdAt As String

    Sub New()
        _organisation = New List(Of clsTSOOrganisationWeb)
        _updatedAt = ""
        _createdAt = ""
    End Sub
End Class
