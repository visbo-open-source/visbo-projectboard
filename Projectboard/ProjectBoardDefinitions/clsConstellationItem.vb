
Public Class clsConstellationItem

    Public Property projectName As String
    Public Property variantName As String
    Public Property Start As Date
    Public Property show As Boolean
    Public Property zeile As Integer

    Private _reasonToInclude As String = ""
    Public Property reasonToInclude As String
        Get
            reasonToInclude = _reasonToInclude
        End Get
        Set(value As String)
            If Not IsNothing(value) Then
                _reasonToInclude = value
            End If
        End Set
    End Property

    Private _reasonToExclude As String = ""
    Public Property reasonToExclude As String
        Get
            reasonToExclude = _reasonToExclude
        End Get
        Set(value As String)
            If Not IsNothing(value) Then
                _reasonToExclude = value
            End If
        End Set
    End Property


    Sub New()

        _projectName = ""
        _variantName = ""
        _Start = StartofCalendar.AddMonths(-1)
        _show = True
        _zeile = 0
        _reasonToInclude = ""
        _reasonToExclude = ""

    End Sub
End Class
