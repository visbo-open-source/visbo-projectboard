Public Class clsWebVCSettingOrganisation

    Inherits clsWebOutput
    Public Property vcsetting As List(Of clsVCSettingOrganisation)
    Public Property updatedAt As String
    Public Property createdAt As String

    Sub New()
        _vcsetting = New List(Of clsVCSettingOrganisation)
        _updatedAt = ""
        _createdAt = ""
    End Sub
End Class
