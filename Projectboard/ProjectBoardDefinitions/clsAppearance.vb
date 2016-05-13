Imports xlNS = Microsoft.Office.Interop.Excel

''' <summary>
''' beschreibt, wie ein Element dieser Darstellungsklasse beschrieben werden soll  
''' </summary>
''' <remarks></remarks>
Public Class clsAppearance

    Public Property name As String
    Public Property isMilestone As Boolean
    Public Property form As xlNS.Shape

    Public Sub New()
        _name = ""
        _isMilestone = False
        _form = Nothing
    End Sub
End Class
