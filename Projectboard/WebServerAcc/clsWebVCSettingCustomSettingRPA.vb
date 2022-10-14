Public Class clsWebVCSettingCustomSettingRPA
    Inherits clsWebOutput
    Public Property vcsetting As List(Of clsVCSettingCustomSettingsRPA)
    Public Property updatedAt As String
    Public Property createdAt As String

    Sub New()
        _vcsetting = New List(Of clsVCSettingCustomSettingsRPA)
        _updatedAt = ""
        _createdAt = ""
    End Sub
End Class
