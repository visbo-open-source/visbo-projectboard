Public Class clsVCOrganisationWeb
    ' besser? as list(of clsRollendefinitionWeb)
    Public Property vcroles As List(Of clsVCrole)
    Public Property vccosts As List(Of clsVCcost)
    Public Property validFrom As Date
    Sub New()
        _vcroles = New List(Of clsVCrole)
        _vccosts = New List(Of clsVCcost)
        _validFrom = Date.Now.Date
    End Sub
End Class
