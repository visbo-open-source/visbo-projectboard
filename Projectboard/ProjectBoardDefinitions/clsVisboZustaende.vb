''' <summary>
''' enthält bestimme Zustands-Variablen der Projekt-Tafel 
''' </summary>
''' <remarks></remarks>
Public Class clsVisboZustaende

    ' enthält den zuletzt eingegebenen Rollen-Kosten Namen 
    Private _oldValue As String

    Public Property showTimeZoneBalken As Boolean
    Public Property projectBoardMode As ptModus


    Public Property meMaxZeile As Integer

    ' nimmt im Massen-Edit Ressourcen die Spalten-Nummer für Ressource-/Kostenauf 
    Public Property meColRC As Integer
    ' nimmt  im Massen-Edit Ressourcen die Spalten-Nummer für den Projekt-Namen auf , im Massen Edit Termine den Elem-Name 
    Public Property meColpName As Integer = 2
    ' nimmt  im Massen-Edit Ressourcen die Spalten-Nummer für StartData auf  , im MassenEdit Termine Startdate
    Public Property meColSD As Integer
    ' nimmt  im Massen-Edit Ressourcen die Spalten-Nummer für EndData  , im MassenEdit Termine Ende-date
    Public Property meColED As Integer
    ' nimmt das letzte Projekt auf, zu dem zuletzt Informationen angezeigt/aktualisiert wurden ...
    Public Property currentProject As clsProjekt
    ' hat den letzten Stand in der Datenbank zu dem Projekt, das zuletzt angezeigt wurde 
    Public Property currentProjectinSession As clsProjekt

    ' wird in MassEdit Termine verwendet ... 
    Private _currentElemID As String
    Public Property currentElemID As String
        Get
            currentElemID = _currentElemID
        End Get
        Set(value As String)
            If Not IsNothing(value) Then
                _currentElemID = value
            Else
                _currentElemID = ""
            End If
        End Set
    End Property

    Public ReadOnly Property currentZeileIsMilestone As Boolean
        Get
            Dim tmpResult As Boolean = Nothing
            If _currentElemID <> "" Then
                tmpResult = elemIDIstMeilenstein(_currentElemID)
            End If
            currentZeileIsMilestone = tmpResult
        End Get
    End Property


    Public ReadOnly Property getcurrentPhase() As clsPhase
        Get
            Dim cPhase As clsPhase = Nothing

            If currentElemID <> "" Then
                If Not elemIDIstMeilenstein(currentElemID) Then
                    If Not IsNothing(currentProject) Then
                        cPhase = currentProject.getPhaseByID(currentElemID)
                    End If
                End If
            End If

            getcurrentPhase = cPhase
        End Get
    End Property

    Public ReadOnly Property getcurrentMilestone() As clsMeilenstein
        Get
            Dim cMilestone As clsMeilenstein = Nothing

            If currentElemID <> "" Then
                If elemIDIstMeilenstein(currentElemID) Then
                    If Not IsNothing(currentProject) Then
                        cMilestone = currentProject.getMilestoneByID(currentElemID)
                    End If
                End If
            End If

            getcurrentMilestone = cMilestone
        End Get
    End Property


    ''' <summary>
    ''' gibt den zuletzt eingegebenen Wert zurück 
    ''' </summary>
    ''' <returns></returns>
    Public Property oldValue As String
        Get
            oldValue = _oldValue
        End Get
        Set(value As String)
            If IsNothing(value) Then
                _oldValue = ""
            Else
                _oldValue = value
            End If
        End Set
    End Property

    Private _oldRow As Integer
    ''' <summary>
    ''' nimmt die letzte Zeile im massEdit auf 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property oldRow As Integer
        Get
            oldRow = _oldRow
        End Get
        Set(value As Integer)
            If IsNothing(value) Then
                _oldRow = 0
            ElseIf value > 0 Then
                _oldRow = value
            Else
                _oldRow = 0
            End If
        End Set
    End Property

    Sub New()
        _showTimeZoneBalken = False
        _projectBoardMode = ptModus.graficboard
        _meMaxZeile = 0
        _oldValue = ""
        _oldRow = 0
        _currentProject = Nothing
        _currentProjectinSession = Nothing
        _currentElemID = ""

    End Sub
End Class
