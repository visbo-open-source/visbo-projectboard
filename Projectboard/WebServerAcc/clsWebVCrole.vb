Public Class clsWebVCrole

    Inherits clsWebOutput
    Public Property vcrole As List(Of clsVCrole)
    Public Property validFrom As Date

    Sub New()
        _vcrole = New List(Of clsVCrole)
        _validFrom = "#1.1.2010#"
    End Sub
End Class
