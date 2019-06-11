''' <summary>
''' enthält alle gültigen Organisationen und Tagessätze
''' </summary>
Public Class clsOrganisations

    ' Der integer key entspricht der Monatsspalte ab StartofCalendar; damit lässt sich die entsprechend gültige Org schnell bestimmen
    Private _validOrganisations As SortedList(Of Integer, clsOrganisation)

    Public ReadOnly Property count As Integer
        Get
            count = _validOrganisations.Count
        End Get
    End Property
    Public ReadOnly Property getOrganisationValidAt(ByVal datum As Date) As clsOrganisation
        Get
            Dim searchkey As Integer = getColumnOfDate(datum)
            Dim tmpOrga As clsOrganisation = Nothing

            If searchkey < getColumnOfDate(StartofCalendar) Then
                ' nichts tun - es wird nach der Orga gefragt, die vor dem StartofCalendar gültig war: Nonsense 
            Else
                Dim found As Boolean = False
                Dim ix As Integer = _validOrganisations.Count - 1

                Do While ix >= 0 And Not found
                    If searchkey >= _validOrganisations.ElementAt(ix).Key Then
                        ' das erste Auftreten searchkey > Orga-validFrom heisst : das ist die gesuchte Orga 
                        found = True
                    ElseIf ix = _validOrganisations.Count - 1 And searchkey < _validOrganisations.ElementAt(ix).Key Then
                        found = True
                    Else
                        ix = ix - 1
                    End If
                Loop

                If found Then
                    tmpOrga = _validOrganisations.ElementAt(ix).Value
                End If

            End If

            getOrganisationValidAt = tmpOrga
        End Get
    End Property
    ''' <summary>
    ''' fügt in die Organisationsliste eine neue Orga ein 
    ''' und baut dabei die OrgaTeamChilds für die jeweiligen Rollen mit auf
    ''' </summary>
    ''' <param name="orga"></param>
    Public Sub addOrga(ByVal orga As clsOrganisation)

        Dim key As Integer = getColumnOfDate(orga.validFrom)
        If _validOrganisations.ContainsKey(key) Then
            _validOrganisations.Remove(key)
        End If

        ' ur:2019-05-29: wird nun in retrieveOrganisationfromDB erledigt
        'Try
        '    ' jetzt wird in der Orga die virtuelle Orga-/Team Struktur aufgebaut
        '    ' die ist dafür notwendig, dass die entsprechende Orga-Einheit den Personal-aufwand ausweist, der durch Multi-Orga-Teams entsteht. 
        '    ' Die Orga-Einheit, in der der Team Aufwand ausgewiesen wird, ist die jüngste (Groß-)Elternteil der allen Team-Mitgliedern gemeinsam ist. 
        '    Call orga.allRoles.buildOrgaTeamChilds()
        'Catch ex As Exception
        '    Call MsgBox("Fehler in Organisations Strukur: Zuordnung Teams als virtuelle Kinder fehlgeschlagen!" & vbLf & "Bitte kontaktieren Sie ihren System-Admin")
        'End Try


        ' jetzt kann fehlerfrei eingetragen werden 
        _validOrganisations.Add(key, orga)

    End Sub
    Sub New()
        _validOrganisations = New SortedList(Of Integer, clsOrganisation)

    End Sub

End Class
