Public Class clsWebVCSettingCustomfields

    Inherits clsWebOutput
    Public Property vcsetting As List(Of clsVCSettingCustomfields)
    Public Property updatedAt As String
    Public Property createdAt As String

    Sub New()
        _vcsetting = New List(Of clsVCSettingCustomfields)
        _updatedAt = ""
        _createdAt = ""
    End Sub
End Class
