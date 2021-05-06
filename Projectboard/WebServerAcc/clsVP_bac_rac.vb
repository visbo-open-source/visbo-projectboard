Public Class clsVP_bac_rac
    ' definition of the input of createProjectFromTemplate
    Inherits clsVP
    Public Property bac As Double
    Public Property rac As Double
    Public Property description As String
    Public Property startDate As Date
    Public Property endDate As Date


    Sub New()
        _bac = 0.0
        _rac = 0.0
        _description = "not yet defined"
        _startDate = Date.MinValue
        _endDate = Date.MinValue
    End Sub
End Class
