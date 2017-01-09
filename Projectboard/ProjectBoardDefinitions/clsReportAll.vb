Imports ProjectBoardDefinitions
Imports System.Xml
Imports System.Xml.Schema

<Serializable()> _
Public Class clsReportAll
    Inherits clsReport

    Private reportVersion As String = "1.0"
    Private reportProjectsWithNoMPmayPass As Boolean
    Private reportBeschreibung As String

    ''' <summary>
    ''' schreibt/liest ob Phasen mit Namen beschriftet werden sollen
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property projectsWithNoMPmayPass As Boolean
        Get
            projectsWithNoMPmayPass = reportProjectsWithNoMPmayPass
        End Get
        Set(value As Boolean)
            reportProjectsWithNoMPmayPass = value
        End Set
    End Property
    Public Property description As String
        Get
            description = reportBeschreibung
        End Get
        Set(value As String)
            reportBeschreibung = value
        End Set
    End Property

    Sub New()

        projectsWithNoMPmayPass = True
        description = ""
    End Sub
End Class
