Public Class clsWebVCSettingAppearance

    Inherits clsWebOutput
    Public Property vcsetting As List(Of clsVCSettingAppearance)
    Public Property updatedAt As String
    Public Property createdAt As String

    Sub New()
        _vcsetting = New List(Of clsVCSettingAppearance)
        _updatedAt = ""
        _createdAt = ""
    End Sub
End Class
