Public Class clsOptimizationObject
    Public Property projectName As String
    ' Änderung tk: kann jetzt eine Reihe von Werten aufnehmen 
    Public Property offset As Integer()
    Public Property bestValue As Double
    Public Sub New()
        _bestValue = -1.0
    End Sub
End Class
