''' <summary>
''' enthält bestimme Zustands-Variablen der Projekt-Tafel 
''' </summary>
''' <remarks></remarks>
Public Class clsVisboZustaende

    Private _auslastungsArray(,) As Double
    Private _oldValue As String

    Public Property showTimeZoneBalken As Boolean
    Public Property projectBoardMode As Integer

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

    ' wird jetzt von getUpdatedAuslastungsArray übernommen ...
    ''Public ReadOnly Property getAuslastungsArray(ByVal von As Integer, ByVal bis As Integer, _
    ''                                             ByVal percentValues As Boolean) As Double(,)
    ''    Get
    ''        If IsNothing(_auslastungsArray) Then
    ''            Try
    ''                _auslastungsArray = ShowProjekte.getAuslastungsArray(von, bis, percentValues)
    ''            Catch ex As Exception
    ''                ReDim _auslastungsArray(RoleDefinitions.Count - 1, bis - von + 1)
    ''            End Try
    ''        Else
    ''            If _auslastungsArray.Length = (RoleDefinitions.Count - 1) * (bis - von + 1) Then
    ''                ' alles gut 
    ''            Else
    ''                Try
    ''                    _auslastungsArray = ShowProjekte.getAuslastungsArray(von, bis, percentValues)
    ''                Catch ex As Exception
    ''                    ReDim _auslastungsArray(RoleDefinitions.Count - 1, bis - von + 1)
    ''                End Try
    ''            End If
    ''        End If

    ''        getAuslastungsArray = _auslastungsArray

    ''    End Get
    ''End Property

    ''' <summary>
    ''' aktualisiert den Auslastungs-Array und gibt ihn zurück
    ''' </summary>
    ''' <param name="roleNames"></param>
    ''' <param name="von"></param>
    ''' <param name="bis"></param>
    ''' <param name="percentValues"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getUpDatedAuslastungsArray(ByVal roleNames As Collection,
                                                            ByVal von As Integer, ByVal bis As Integer,
                                                            ByVal percentValues As Boolean) As Double(,)
        Get
            Dim resultValues() As Double = Nothing
            Dim createArray As Boolean = False

            If Not IsNothing(roleNames) Then
                If roleNames.Count = 0 Then
                    createArray = True
                End If
            Else
                createArray = True
            End If

            If IsNothing(_auslastungsArray) Then
                createArray = True
            ElseIf _auslastungsArray.Length <> RoleDefinitions.Count * (bis - von + 2) Then
                createArray = True
            End If

            If createArray Then
                Try
                    _auslastungsArray = ShowProjekte.getAuslastungsArray(von, bis, percentValues)
                Catch ex As Exception
                    ReDim _auslastungsArray(RoleDefinitions.Count - 1, bis - von + 1)
                End Try
            End If


            If Not IsNothing(roleNames) Then

                If roleNames.Count > 0 Then
                    For ax As Integer = 1 To roleNames.Count

                        Try
                            Dim roleName As String = CStr(roleNames.Item(ax))
                            Dim roleID As Integer = RoleDefinitions.getRoledef(roleName).UID
                            resultValues = ShowProjekte.getAuslastungsArrayOfRole(roleID, von, bis, percentValues)
                            ' hier muss nun der _auslastungsArray aktualisiert werden 
                            For ix As Integer = 0 To bis - von + 1
                                _auslastungsArray(roleID - 1, ix) = resultValues(ix)
                            Next
                        Catch ex As Exception

                        End Try

                    Next
                End If


            End If

            getUpDatedAuslastungsArray = _auslastungsArray
        End Get
    End Property

    ''' <summary>
    ''' Speicher für den Auslastungs-Array wieder freigeben 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub clearAuslastungsArray()
        _auslastungsArray = Nothing
    End Sub

    Sub New()
        _showTimeZoneBalken = False
        _projectBoardMode = ptModus.graficboard
        _meMaxZeile = 0
        _oldValue = ""
        _oldRow = 0
        _currentProject = Nothing
        _currentProjectinSession = Nothing
        _currentElemID = ""
        _auslastungsArray = Nothing
    End Sub
End Class
