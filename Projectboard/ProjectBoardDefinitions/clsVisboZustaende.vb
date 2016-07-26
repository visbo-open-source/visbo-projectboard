''' <summary>
''' enthält bestimme Zustands-Variablen der Projekt-Tafel 
''' </summary>
''' <remarks></remarks>
Public Class clsVisboZustaende

    Private _auslastungsArray(,) As Double
    Public Property showTimeZoneBalken As Boolean
    Public Property projectBoardMode As Integer

    Public Property meMaxZeile As Integer

    ' nimmt im Massen-Edit Ressourcen die Spalten-Nummer für Ressource-/Kostenauf 
    Public Property meColRC As Integer
    ' nimmt  im Massen-Edit Ressourcen die Spalten-Nummer für StartData auf  
    Public Property meColSD As Integer
    ' nimmt  im Massen-Edit Ressourcen die Spalten-Nummer für EndData  
    Public Property meColED As Integer

    Public Property oldValue As String

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
    Public ReadOnly Property getUpDatedAuslastungsArray(ByVal roleNames As Collection, _
                                                            ByVal von As Integer, ByVal bis As Integer, _
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
        _auslastungsArray = Nothing
    End Sub
End Class
