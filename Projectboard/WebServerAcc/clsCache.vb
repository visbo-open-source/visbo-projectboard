Public Class clsCache

    Public Property VPsN As SortedList(Of String, clsVP)
    Public Property VPsId As SortedList(Of String, clsVP)

    Public Property VPvs As SortedList(Of String, SortedList(Of String, clsVarTs))
    Public Property updateDelay As Long = 60
    'Public Property varTsListe As SortedList(Of String, clsVarTs)
    Public Sub New()
        _VPsN = New SortedList(Of String, clsVP)
        _VPsId = New SortedList(Of String, clsVP)
        _VPvs = New SortedList(Of String, SortedList(Of String, clsVarTs))
        _updateDelay = 5
    End Sub


    '''' <summary>
    '''' gets or sets the sortedlist of (string, sortedList(of string, clsvarts)
    '''' </summary>
    '''' <value></value>
    '''' <returns></returns>
    '''' <remarks></remarks>
    'Public Property liste() As SortedList(Of String, SortedList(Of String, clsVarTs))
    '    Get
    '        liste = _cacheVPv
    '    End Get

    '    Set(value As SortedList(Of String, SortedList(Of String, clsVarTs)))

    '        If Not IsNothing(value) Then
    '            _cacheVPv = value
    '        End If

    '    End Set

    'End Property

    '''' <summary>
    '''' gibt die Anzahl Listenelemente der Sorted Liste zurück 
    '''' </summary>
    '''' <value></value>
    '''' <returns></returns>
    '''' <remarks></remarks>
    'Public ReadOnly Property Count() As Integer
    '    Get
    '        Count = _cacheVPv.Count
    '    End Get
    'End Property
    '''' <summary>
    '''' true, wenn die SortedList ein Element mit angegebenem Key enthält
    '''' false, sonst
    '''' </summary>
    '''' <param name="key"></param>
    '''' <value></value>
    '''' <returns></returns>
    '''' <remarks></remarks>
    'Public ReadOnly Property Containskey(ByVal key As String) As Boolean
    '    Get
    '        Containskey = _cacheVPv.ContainsKey(key)
    '    End Get
    'End Property



    '''' <summary>
    '''' fügt der Sorted List eine Liste mit Timestamps zu einer Variante des vpid hinzu
    '''' </summary>
    '''' <param name="vpid"></param>
    '''' <param name="varts"></param>
    'Public Sub Add(ByVal vpid As String, ByVal varts As SortedList(Of String, clsVarTs))

    '    If _cacheVPv.ContainsKey(vpid) Then
    '        _cacheVPv.Remove(vpid)
    '    End If

    '    _cacheVPv.Add(vpid, varts)

    'End Sub


    ' existiert es bereits ? 
    ' wenn ja, dann löschen ...


    ''' <summary>
    ''' Cache füllen mit ProjektShortVersions zum Zeitpunkt timeCached
    ''' </summary>
    ''' <param name="result">Liste von KurzProjektVersionen</param>
    ''' <param name="timeCached">Zeitpunkt, zu dem der Cache gefüllt wurde</param>
    Public Sub createVPvShort(ByVal result As List(Of clsProjektWebShort), ByVal timeCached As Date)

        Dim vpid As String = ""
        Dim hvpv As SortedList(Of String, clsVarTs)
        Dim hVarTS As New clsVarTs

        If result.Count > 0 Then
            vpid = result.ElementAt(0).vpid
        End If

        If _VPvs.ContainsKey(vpid) Then
            hvpv = _VPvs(vpid)
        Else
            hvpv = New SortedList(Of String, clsVarTs)
        End If

        For Each vpv As clsProjektWebShort In result

            If Not hvpv.ContainsKey(vpv.variantName) Then
                hVarTS = New clsVarTs
            Else
                hVarTS = hvpv(vpv.variantName)
            End If

            ' vpv nur eintragen, wenn der timestamp nicht bereits vorhanden
            If Not hVarTS.tsShort.ContainsKey(vpv.timestamp) Then
                hVarTS.vname = vpv.variantName
                hVarTS.timeCached = timeCached
                hVarTS.tsShort.Add(vpv.timestamp, vpv)
            End If


            If Not hvpv.ContainsKey(vpv.variantName) Then
                hvpv.Add(vpv.variantName, hVarTS)
            End If
        Next

        If _VPvs.ContainsKey(vpid) Then
            _VPvs(vpid) = hvpv
        Else
            _VPvs.Add(vpid, hvpv)
        End If

    End Sub

    ''' <summary>
    ''' Cache füllen mit ProjektShortVersions zum Zeitpunkt timeCached
    ''' </summary>
    ''' <param name="result">Liste von KurzProjektVersionen</param>
    ''' <param name="timeCached">Zeitpunkt, zu dem der Cache gefüllt wurde</param>
    Public Sub createVPvLong(ByVal result As List(Of clsProjektWebLong), ByVal timeCached As Date)

        Dim vpid As String = ""
        Dim hvpv As New SortedList(Of String, clsVarTs)
        Dim hVarTS As New clsVarTs

        If result.Count > 0 Then
            vpid = result.ElementAt(0).vpid
        End If

        If _VPvs.ContainsKey(vpid) Then
            hvpv = _VPvs(vpid)
        Else
            hvpv = New SortedList(Of String, clsVarTs)
        End If

        For Each vpv As clsProjektWebLong In result

            If Not hvpv.ContainsKey(vpv.variantName) Then
                hVarTS = New clsVarTs
            Else
                hVarTS = hvpv(vpv.variantName)
            End If

            If Not hVarTS.tsLong.ContainsKey(vpv.timestamp) Then
                hVarTS.vname = vpv.variantName
                hVarTS.timeCached = timeCached
                hVarTS.tsLong.Add(vpv.timestamp, vpv)
            End If

            If Not hvpv.ContainsKey(vpv.variantName) Then
                hvpv.Add(vpv.variantName, hVarTS)
            End If

        Next

        If _VPvs.ContainsKey(vpid) Then
            _VPvs(vpid) = hvpv
        Else
            _VPvs.Add(vpid, hvpv)
        End If

    End Sub

    Public Function existsInCache(ByVal vpid As String,
                                  ByVal vName As String,
                                  Optional ByVal vpvid As String = "",
                                  Optional ByVal longVersion As Boolean = False) As Boolean

        Dim nothingToDo As Boolean = False

        If vpid <> "" Then
            If _VPvs.ContainsKey(vpid) Then

                If vpvid <> "" Then
                    For vNamelist As Integer = 0 To _VPvs(vpid).Count - 1
                        Dim hvname As String = _VPvs(vpid).ElementAt(vNamelist).Value.vname
                        For Each kvp As KeyValuePair(Of Date, clsProjektWebLong) In _VPvs(vpid)(hvname).tsLong
                            If kvp.Value._id = vpvid Then
                                nothingToDo = True
                                Exit For
                            End If
                        Next
                        If nothingToDo Then
                            Exit For
                        End If
                    Next
                Else

                    If vName <> "" Then

                        If _VPvs(vpid).ContainsKey(vName) Then

                            ' nachsehen, ob im Cache für Projekt vpid die Variante variantName und ihre Timestamps gespeichert sind, 
                            ' wenn ja, dann result-liste aufbauen
                            If Not longVersion Then
                                If _VPvs(vpid)(vName).tsShort.Count > 0 And
                               DateDiff(DateInterval.Minute, _VPvs(vpid)(vName).timeCached, Date.Now) <= updateDelay Then

                                    nothingToDo = True

                                Else
                                    nothingToDo = False
                                End If
                            Else
                                If _VPvs(vpid)(vName).tsLong.Count > 0 And
                               DateDiff(DateInterval.Minute, _VPvs(vpid)(vName).timeCached, Date.Now) <= updateDelay Then

                                    nothingToDo = True

                                Else
                                    nothingToDo = False
                                End If
                            End If
                        Else
                            nothingToDo = False

                        End If


                    Else  ' von if vname <> ""

                        ' nachsehen, ob im Cache für Projekt vpid alle Variante und Timestamps gespeichert sind, 
                        ' wenn ja, dann result-liste aufbauen

                        Dim vp As clsVP = _VPsId(vpid)

                        ' VisboProjekt Standard, keine Variante (Variante = "")
                        If _VPvs(vpid).ContainsKey(vName) Then
                            If Not longVersion Then
                                If _VPvs(vpid)(vName).tsShort.Count > 0 And
                               DateDiff(DateInterval.Minute, _VPvs(vpid)(vName).timeCached, Date.Now) <= updateDelay Then

                                    nothingToDo = True
                                Else

                                    nothingToDo = False

                                End If
                            Else
                                If _VPvs(vpid)(vName).tsLong.Count > 0 And
                               DateDiff(DateInterval.Minute, _VPvs(vpid)(vName).timeCached, Date.Now) <= updateDelay Then

                                    nothingToDo = True
                                Else

                                    nothingToDo = False

                                End If
                            End If
                        End If

                        If nothingToDo Then

                            For Each vpvar As clsVPvariant In vp.Variant
                                Try
                                    If _VPvs(vpid).ContainsKey(vpvar.variantName) Then
                                        If Not longVersion Then
                                            If _VPvs(vpid)(vpvar.variantName).tsShort.Count > 0 And
                                               DateDiff(DateInterval.Minute, _VPvs(vpid)(vpvar.variantName).timeCached, Date.Now) <= updateDelay Then

                                                nothingToDo = nothingToDo And True
                                            Else

                                                nothingToDo = nothingToDo And False
                                                Exit For

                                            End If
                                        Else
                                            If _VPvs(vpid)(vpvar.variantName).tsLong.Count > 0 And
                                               DateDiff(DateInterval.Minute, _VPvs(vpid)(vpvar.variantName).timeCached, Date.Now) <= updateDelay Then

                                                nothingToDo = nothingToDo And True
                                            Else

                                                nothingToDo = nothingToDo And False
                                                Exit For
                                            End If
                                        End If


                                    End If
                                Catch ex As Exception

                                End Try

                            Next

                        End If  ' end if von it nothingToDo = true

                    End If    ' end if von vName <> ""

                End If   ' end if von if vpvid <> ""

            Else
                nothingToDo = nothingToDo And False

            End If
        Else
            nothingToDo = nothingToDo And False

        End If

        existsInCache = nothingToDo

    End Function
End Class

