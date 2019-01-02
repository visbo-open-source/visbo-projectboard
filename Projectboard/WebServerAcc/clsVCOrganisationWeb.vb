Public Class clsVCOrganisationWeb
    Public Property vcroles As List(Of clsVCrole)
    Public Property vccosts As List(Of clsVCcost)
    Sub New()
        _vcroles = New List(Of clsVCrole)
        _vccosts = New List(Of clsVCcost)
    End Sub
End Class
