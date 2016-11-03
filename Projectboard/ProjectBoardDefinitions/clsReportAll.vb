Imports ProjectBoardDefinitions
Imports System.Xml
Imports System.Xml.Schema

<Serializable()> _
Public Class clsReportAll
    Inherits clsReport

    Private reportVersion As String = "1.0"
    Private reportProjectsWithNoMPmayPass As Boolean

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
  
    Sub New()

        projectsWithNoMPmayPass = True
    
    End Sub
End Class
