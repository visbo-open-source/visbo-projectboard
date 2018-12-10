Public Class clsWebVCcost

    Inherits clsWebOutput
    Public Property vccost As List(Of clsVCcost)
    Public Property validFrom As Date

    Sub New()
        _vccost = New List(Of clsVCcost)
        _validFrom = "#1.1.2010#"
    End Sub
End Class
