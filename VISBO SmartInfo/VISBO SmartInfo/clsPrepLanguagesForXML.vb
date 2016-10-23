Public Class clsPrepLanguagesForXML

    Public sprachArray As String()
    Public dimen1 As Integer
    Public dimen2 As Integer

    Sub New(ByVal dm1 As Integer, ByVal dm2 As Integer)
        dimen1 = dm1
        dimen2 = dm2
        ReDim sprachArray(dm1 * (dm2 + 1) - 1)
    End Sub

End Class
