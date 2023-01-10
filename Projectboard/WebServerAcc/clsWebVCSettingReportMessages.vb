Public Class clsWebVCSettingReportMessages

    Inherits clsWebOutput
    Public Property vcsetting As List(Of clsVCSettingReportMsg)
    Public Property updatedAt As String
    Public Property createdAt As String

    Sub New()
        _vcsetting = New List(Of clsVCSettingReportMsg)
        _updatedAt = ""
        _createdAt = ""
    End Sub

End Class
