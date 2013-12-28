Public Class clsOptimizationObject
    Public Property projectName As String
    Public Property startOffset As Integer
    Public Property bestValue As Double
    Public Sub New()
        _bestValue = -1.0
    End Sub
End Class
