Public Class clsWebVCSettingEverything

    Inherits clsWebOutput
    Public Property vcsetting As List(Of clsVCSettingEverything)
    Public Property updatedAt As String
    Public Property createdAt As String

    Sub New()
        _vcsetting = New List(Of clsVCSettingEverything)
        _updatedAt = ""
        _createdAt = ""
    End Sub
End Class
