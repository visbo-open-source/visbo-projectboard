Imports xlNS = Microsoft.Office.Interop.Excel
Public Class clsMeilensteinDefinition

    Public Property name As String
    Public Property schwellWert As Integer
    Public Property shapeVorlage As xlNS.ShapeRange
    Public Property UID As Long

    Public Sub New()

    End Sub

End Class
