
Public Class clsWebVCSettingCustomroles

    Inherits clsWebOutput
    Public Property vcsetting As List(Of clsVCSettingCustomroles)
    Public Property updatedAt As String
    Public Property createdAt As String

    Sub New()
        _vcsetting = New List(Of clsVCSettingCustomroles)
        _updatedAt = ""
        _createdAt = ""
    End Sub

End Class

