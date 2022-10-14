Public Class clsWebVCSettingconfiguration
    Inherits clsWebOutput
    Public Property vcsetting As List(Of clsVCSettingConfiguration)
    Public Property updatedAt As String
    Public Property createdAt As String

    Sub New()
        _vcsetting = New List(Of clsVCSettingConfiguration)
        _updatedAt = ""
        _createdAt = ""
    End Sub
End Class