Public Class clsWebVCSettingCustomization

    Inherits clsWebOutput
    Public Property vcsetting As List(Of clsVCSettingCustomization)
    Public Property updatedAt As String
    Public Property createdAt As String

    Sub New()
        _vcsetting = New List(Of clsVCSettingCustomization)
        _updatedAt = ""
        _createdAt = ""
    End Sub
End Class
