Imports ProjectBoardDefinitions
Imports System.Xml
Imports System.Xml.Schema


<Serializable()> _
Public Class clsReport

    Private _name As String

    ' Definitionen, die beim BerechneFormat bestimmt werden und für den Report-Ablauf benötigt werden,
    ' müssen nicht beim Reportprofil in DB abgespeichert werden
    Private reportProjects As SortedList(Of Double, String)
    Private reportCalendarVon As Date
    Private reportCalendarBis As Date

    ' Definitionen, die in DB für das ReportProfil gespeichert werden müssen
    Private reportIsMpp As Boolean
    Private reportPPTTemplate As String

    Private reportPhase As SortedList(Of String, String)
    Private reportMilestone As SortedList(Of String, String)
    Private reportRolle As SortedList(Of String, String)
    Private reportCost As SortedList(Of String, String)
    Private reportTyp As SortedList(Of String, String)
    Private reportBU As SortedList(Of String, String)

    Private reportProjectline As Boolean
    Private reportAllIfOne As Boolean
    Private reportAmpeln As Boolean
    Private reportUseAbbreviation As Boolean
    'Private reportProjectsWithNoMPmayPass as boolean

    Private reportPhName As Boolean
    Private reportPhDate As Boolean
    Private reportMSName As Boolean
    Private reportMSDate As Boolean

    Private reportKwInMilestone As Boolean
    Private reportUseOriginalNames As Boolean

    Private reportVLinien As Boolean
    Private reportShowHorizontals As Boolean
    Private reportLegend As Boolean
    Private reportOnePage As Boolean

    Private reportSortedDauer As Boolean
    Private reportExtendedMode As Boolean

    Private reportFullyContained As Boolean

    Private reportVon As Date
    Private reportBis As Date



    ''' <summary>
    ''' prüft ob irgendein Report gesetzt ist 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property isEmpty As Boolean
        Get
            Dim sum As Integer = reportPhase.Count + reportMilestone.Count + _
                                 reportRolle.Count + reportCost.Count + _
                                 reportTyp.Count + reportBU.Count

            If sum = 0 Then
                isEmpty = True
            Else
                isEmpty = False
            End If

        End Get
    End Property


    ''' <summary>
    ''' kopiert die Angaben vom aktuellen Report ein einen Neuen (übergebenen)
    ''' </summary>
    ''' <param name="newReport"></param>
    ''' <remarks></remarks>
    Public Sub CopyTo(ByRef newReport As clsReport)
        Try

            With newReport
                .name = Me._name
                For Each kvp As KeyValuePair(Of Double, String) In Me.Projects
                    .Projects.Add(kvp.Key, kvp.Value)
                Next
                .CalendarVonDate = Me.reportCalendarVon
                .CalendarBisDate = Me.reportCalendarBis
                .isMpp = Me.reportIsMpp
                If Not IsNothing(Me.reportPPTTemplate) Then
                    .PPTTemplate = Me.reportPPTTemplate
                End If
                .Phases = copyList(Me.reportPhase)
                .Milestones = copyList(Me.reportMilestone)
                .Roles = copyList(Me.reportRolle)
                .Costs = copyList(Me.reportCost)
                .Typs = copyList(Me.reportTyp)
                .BUs = copyList(Me.reportBU)
                .calcRepVonBis(Me.reportVon, Me.reportBis)
                .ProjectLine = Me.reportProjectline
                .Ampeln = Me.reportAmpeln
                .AllIfOne = Me.reportAllIfOne
                .PhName = Me.reportPhName
                .PhDate = Me.reportPhDate
                .MSName = Me.reportMSName
                .MSDate = Me.reportMSDate
                .VLinien = Me.reportVLinien
                .Legend = Me.reportLegend
                .SortedDauer = Me.reportSortedDauer
                .OnePage = Me.reportOnePage
                .ExtendedMode = Me.reportExtendedMode
                .KwInMilestone = Me.reportKwInMilestone
            End With

        Catch ex As Exception
            Throw New ArgumentException("Fehler in der Property für clsReport")
        End Try
    End Sub




    ''' <summary>
    ''' schreibt/liest das Datum des Beginn des ausgewählten zeitl Bereiches
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property CalendarVonDate As Date
        Get
            CalendarVonDate = reportCalendarVon
        End Get
        Set(value As Date)
            If value >= StartofCalendar Then
                reportCalendarVon = value
            Else
                reportCalendarVon = StartofCalendar
                'Throw New ArgumentException("Datum muss nach StartofCalendar liegen")
            End If

        End Set
    End Property
    ''' <summary>
    ''' schreibt/liest das Datum des Endes des ausgewählten zeitl Bereiches
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property CalendarBisDate As Date
        Get
            CalendarBisDate = reportCalendarBis
        End Get
        Set(value As Date)
            If value >= StartofCalendar And value > Me.CalendarVonDate Then
                reportCalendarBis = value
            Else
                reportCalendarBis = Date.Now
                'Throw New ArgumentException("Datum muss nach StartofCalendar und vor heute liegen")
            End If

        End Set
    End Property

    ''' <summary>
    ''' schreibt/liest die Report Collection der BUs
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property BUs() As SortedList(Of String, String)
        Get
            BUs = reportBU
        End Get
        Set(value As SortedList(Of String, String))

            If Not IsNothing(value) Then
                reportBU = value
            End If

        End Set
    End Property

    ''' <summary>
    ''' schreibt/liest die Report Collection der Typen
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Typs() As SortedList(Of String, String)
        Get
            Typs = reportTyp
        End Get
        Set(value As SortedList(Of String, String))

            If Not IsNothing(value) Then
                reportTyp = value
            End If

        End Set
    End Property

    ''' <summary>
    ''' schreibt/liest die Report sortierte Liste der Projekte
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Projects() As SortedList(Of Double, String)
        Get
            Projects = reportProjects
        End Get
        Set(value As SortedList(Of Double, String))

            If Not IsNothing(value) Then
                reportProjects = value
            End If

        End Set
    End Property

    ''' <summary>
    ''' schreibt/liest die Report Collection der Phasen
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Phases() As SortedList(Of String, String)
        Get
            Phases = reportPhase
        End Get
        Set(value As SortedList(Of String, String))

            If Not IsNothing(value) Then
                reportPhase = value
            End If

        End Set
    End Property

    ''' <summary>
    ''' schreibt/liest die Report Collection der Meilensteine
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Milestones() As SortedList(Of String, String)
        Get
            Milestones = reportMilestone
        End Get
        Set(value As SortedList(Of String, String))

            If Not IsNothing(value) Then
                reportMilestone = value
            End If

        End Set
    End Property

    ''' <summary>
    ''' schreibt/liest die Report Collection der Rolle
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Roles() As SortedList(Of String, String)
        Get
            Roles = reportRolle
        End Get
        Set(value As SortedList(Of String, String))

            If Not IsNothing(value) Then
                reportRolle = value
            End If

        End Set
    End Property

    ''' <summary>
    ''' schreibt/liest die Report Collection der Kostenart
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Costs() As SortedList(Of String, String)
        Get
            Costs = reportCost
        End Get
        Set(value As SortedList(Of String, String))

            If Not IsNothing(value) Then
                reportCost = value
            End If

        End Set
    End Property


    ''' <summary>
    ''' liest bzw. schreibt den Namen des Reports
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property name As String

        Get
            name = _name
        End Get
        Set(value As String)

            If Not IsNothing(value) Then
                If value.Trim.Length > 0 Then
                    _name = value
                Else
                    _name = "XXX"
                End If
            Else
                _name = "XXX"
            End If

        End Set
    End Property
    ''' <summary>
    ''' schreibt/liest ob Report eine Constellation betrifft
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property isMpp As Boolean
        Get
            isMpp = reportIsMpp
        End Get
        Set(value As Boolean)
            reportIsMpp = value
        End Set
    End Property
    ''' <summary>
    ''' schreibt/liest ob ProjektLinie gezeichnet werden soll
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ProjectLine As Boolean
        Get
            ProjectLine = reportProjectline
        End Get
        Set(value As Boolean)
            reportProjectline = value
        End Set
    End Property
    ''' <summary>
    ''' schreibt/liest ob Ampel gezeichnet werden soll
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Ampeln As Boolean
        Get
            Ampeln = reportAmpeln
        End Get
        Set(value As Boolean)
            reportAmpeln = value
        End Set
    End Property
    ''' <summary>
    ''' schreibt/liest ob alle Planelemente  gezeichnet werden soll, wenn nur eines in den ausgewählten Zeitraum fällt
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property AllIfOne As Boolean
        Get
            AllIfOne = reportAllIfOne
        End Get
        Set(value As Boolean)
            reportAllIfOne = value
        End Set
    End Property
    ''' <summary>
    ''' schreibt/liest ob Phasen mit Namen beschriftet werden sollen
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property PhName As Boolean
        Get
            PhName = reportPhName
        End Get
        Set(value As Boolean)
            reportPhName = value
        End Set
    End Property
    ''' <summary>
    ''' schreibt/liest ob Phasen mit Datum beschriftet werden sollen
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property PhDate As Boolean
        Get
            PhDate = reportPhDate
        End Get
        Set(value As Boolean)
            reportPhDate = value
        End Set
    End Property

    ''' <summary>
    ''' schreibt/liest ob Meilenstein mit Namen beschriftet werden sollen
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property MSName As Boolean
        Get
            MSName = reportMSName
        End Get
        Set(value As Boolean)
            reportMSName = value
        End Set
    End Property
    ''' <summary>
    ''' schreibt/liest ob Meilenstein mit Datum beschriftet werden sollen
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property MSDate As Boolean
        Get
            MSDate = reportMSDate
        End Get
        Set(value As Boolean)
            reportMSDate = value
        End Set
    End Property
    ''' <summary>
    ''' schreibt/liest ob vertikale Linien werden sollen
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property VLinien As Boolean
        Get
            VLinien = reportVLinien
        End Get
        Set(value As Boolean)
            reportVLinien = value
        End Set
    End Property
    ''' <summary>
    ''' schreibt/liest ob eine Legende gezeichnet werden soll
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Legend As Boolean
        Get
            Legend = reportLegend
        End Get
        Set(value As Boolean)
            reportLegend = value
        End Set
    End Property
    ''' <summary>
    ''' schreibt/liest ob die Projekte sortiert nach Dauer gezeichnet werden sollen
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property SortedDauer As Boolean
        Get
            SortedDauer = reportSortedDauer
        End Get
        Set(value As Boolean)
            reportSortedDauer = value
        End Set
    End Property
    ''' <summary>
    ''' schreibt/liest ob alles auf eine Seite gezeichnet werden soll
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property OnePage As Boolean
        Get
            OnePage = reportOnePage
        End Get
        Set(value As Boolean)
            reportOnePage = value
        End Set
    End Property
    ''' <summary>
    ''' schreibt/liest ob im extended Mode  gezeichnet werden soll
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ExtendedMode As Boolean
        Get
            ExtendedMode = reportExtendedMode
        End Get
        Set(value As Boolean)
            reportExtendedMode = value
        End Set
    End Property
    ''' <summary>
    ''' schreibt/liest ob eine Phase ganz gezeichnet werden soll, wenn im TimeRange enthalten
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property FullyContained As Boolean
        Get
            FullyContained = reportFullyContained
        End Get
        Set(value As Boolean)
            reportFullyContained = value
        End Set
    End Property
    ''' <summary>
    ''' schreibt/liest ob horizontale Linien (siehe BHTC) gezeichnet werden soll
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ShowHorizontals As Boolean
        Get
            ShowHorizontals = reportShowHorizontals
        End Get
        Set(value As Boolean)
            reportShowHorizontals = value
        End Set
    End Property
    ''' <summary>
    ''' schreibt/liest ob nur die Abkürzungen verwendet werden sollen
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property UseAbbreviation As Boolean
        Get
            UseAbbreviation = reportUseAbbreviation
        End Get
        Set(value As Boolean)
            reportUseAbbreviation = value
        End Set
    End Property
    ''' <summary>
    ''' schreibt/liest ob die Original Namen verwendet werden sollen
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property UseOriginalNames As Boolean
        Get
            UseOriginalNames = reportUseOriginalNames
        End Get
        Set(value As Boolean)
            reportUseOriginalNames = value
        End Set
    End Property
    ''' <summary>
    ''' schreibt/liest ob die Original Namen verwendet werden sollen
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property KwInMilestone As Boolean
        Get
            KwInMilestone = reportKwInMilestone
        End Get
        Set(value As Boolean)
            reportKwInMilestone = value
        End Set
    End Property


    ''' <summary>
    ''' schreibt/liest den Namen des verwendeten Template.pptx
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property PPTTemplate As String
        Get
            PPTTemplate = reportPPTTemplate
        End Get
        Set(value As String)
            reportPPTTemplate = value
        End Set
    End Property

    ''' <summary>
    ''' liest das Datum des Beginn des ausgewählten zeitl Bereiches
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property VonDate As Date
        Get
            VonDate = reportVon
        End Get

    End Property
    ''' <summary>
    ''' liest das Datum des Endes des ausgewählten zeitl Bereiches
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property BisDate As Date
        Get
            BisDate = reportBis
        End Get

    End Property

    ''' <summary>
    ''' setzt die Range von und bis als Datum im Report-Profil. wenn sowohl von als auch bis mit dem Wert StartoFCalendar aufgerufen wird, dann 
    ''' ist keine Range gesetzt 
    ''' </summary>
    ''' <param name="von"></param>
    ''' <param name="bis"></param>
    ''' <remarks></remarks>
    Public Sub calcRepVonBis(ByVal von As Date, ByVal bis As Date)

        Try
            If von > StartofCalendar And bis > StartofCalendar Then

                If DateDiff(DateInterval.Day, von, bis) > 0 Then
                    Me.reportVon = von
                    Me.reportBis = bis

                Else
                    Throw New ArgumentException("Datum 'von' muss vor Datum 'bis' liegen")
                End If

            ElseIf getColumnOfDate(von) = getColumnOfDate(StartofCalendar) And getColumnOfDate(bis) = getColumnOfDate(StartofCalendar) Then
                ' das ist das Zeichen, dass kein Range definiert ist 
                Me.reportVon = StartofCalendar
                Me.reportBis = StartofCalendar
            Else
                Throw New ArgumentException("Datum 'von' und 'bis' müssen nach dem ' " & StartofCalendar.ToString & " ' liegen")
            End If


        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try
    End Sub



    ''' <summary>
    ''' fügt dem Business Unit Report einen Eintrag hinzu
    ''' wenn der Eintrag  vorhanden ist, wird nichts eingefügt
    ''' aber auch keine Fehlermeldung geworfen
    ''' </summary>
    ''' <param name="businessUnit"></param>
    ''' <remarks></remarks>
    Public Sub addBU(ByVal businessUnit As String)

        If reportBU.ContainsKey(businessUnit) Then
            ' nichts tun ..
        Else

            If Not IsNothing(businessUnit) Then
                reportBU.Add(businessUnit, businessUnit)
            End If

        End If

    End Sub

    ''' <summary>
    ''' entfernt aus dem Business Unit Report einen Eintrag
    ''' wenn der Eintrag nicht vorhanden ist, wird nichts entfernt
    ''' aber auch keine Fehlermeldung geworfen 
    ''' </summary>
    ''' <param name="businessUnit"></param>
    ''' <remarks></remarks>
    Public Sub removeBU(ByVal businessUnit As String)

        If Not IsNothing(businessUnit) Then
            If reportBU.ContainsKey(businessUnit) Then
                reportBU.Remove(businessUnit)
            Else
                ' nichts tun ..
            End If
        End If

    End Sub




    Sub New()
        _name = "Dummy"

        ' Definitionen, die beim BerechneFormat bestimmt werden und für den Report-Ablauf benötigt werden,
        ' müssen nicht beim Reportprofil in DB abgespeichert werden
        reportProjects = New SortedList(Of Double, String)
        reportCalendarVon = StartofCalendar
        reportCalendarBis = Date.Now

        ' Definitionen, die in DB für das ReportProfil gespeichert werden müssen
        reportIsMpp = False
        reportPPTTemplate = ""
        reportBU = New SortedList(Of String, String)
        reportPhase = New SortedList(Of String, String)
        reportMilestone = New SortedList(Of String, String)
        reportTyp = New SortedList(Of String, String)
        reportRolle = New SortedList(Of String, String)
        reportCost = New SortedList(Of String, String)
        reportVon = StartofCalendar
        reportBis = Date.Now

        reportPhName = True
        reportPhDate = True
        reportMSName = True
        reportMSDate = True
        reportVLinien = False
        reportLegend = False
        reportOnePage = True

        ' für BHTC immer true
        reportExtendedMode = True
        reportShowHorizontals = False
        ' für BHTC immer false
        reportAmpeln = False
        reportAllIfOne = False
        reportFullyContained = False
        reportSortedDauer = False
        reportProjectline = False
        reportUseOriginalNames = False
        reportKwInMilestone = False

    End Sub

    ''' <summary>
    ''' legt ein neues ReportProfil an unter Angabe der bekannten Filter Collections
    ''' Eingabe Parameter kann auch Nothing sein 
    ''' </summary>
    ''' <param name="kennung">Name des Reports</param>
    ''' <param name="rPhase">report Phase</param>
    ''' <param name="rMilestone">report Meilenstein</param>
    ''' <param name="rBU">report BU</param>
    ''' <param name="rTyp">report Typ</param>
    ''' <param name="rRolle">report Rolle</param>
    ''' <param name="rCost">report Cost</param>
    ''' <remarks></remarks>
    Sub New(ByVal kennung As String, _
            ByVal rPhase As SortedList(Of String, String), ByVal rMilestone As SortedList(Of String, String), _
                ByVal rBU As SortedList(Of String, String), ByVal rTyp As SortedList(Of String, String), _
                               ByVal rRolle As SortedList(Of String, String), ByVal rCost As SortedList(Of String, String))

        reportPhase = New SortedList(Of String, String)
        reportPhase = copyList(rPhase)

        reportMilestone = New SortedList(Of String, String)
        reportMilestone = copyList(rMilestone)

        reportRolle = New SortedList(Of String, String)
        reportRolle = copyList(rRolle)

        reportCost = New SortedList(Of String, String)
        reportCost = copyList(rCost)

        reportBU = New SortedList(Of String, String)
        reportBU = copyList(rBU)

        reportTyp = New SortedList(Of String, String)
        reportTyp = copyList(rTyp)

        reportCalendarVon = StartofCalendar
        reportCalendarBis = Date.Now
        reportVon = StartofCalendar
        reportBis = Date.Now
        reportPPTTemplate = ""

        name = kennung

    End Sub
End Class
