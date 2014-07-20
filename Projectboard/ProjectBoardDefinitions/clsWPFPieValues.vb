Public Class clsWPFPieValues

    Public Property name As String
    Public Property value As Double
    Public Property color As UInt32
    Public Property toolTip As String


    Sub New()
        name = ""
        toolTip = ""
        value = 0.0
        color = CType(RGB(255, 255, 255), UInt32)
    End Sub


End Class
