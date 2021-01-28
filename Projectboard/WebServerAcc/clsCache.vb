
Imports ProjectBoardDefinitions

Public Class clsCache
    ' alle VP sortiert nach Name
    Public Property VPsN As SortedList(Of String, clsVP)
    ' alle VP sortiert nach ID
    Public Property VPsId As SortedList(Of String, clsVP)
    ' VPversions sortiert nach vpvid
    Public Property VPvs As SortedList(Of String, SortedList(Of String, clsVarTs))
    Public Property VCrole As SortedList(Of String, clsVCrole)
    Public Property VCcost As SortedList(Of String, clsVCcost)
    Public Property updateDelay As Long
    'Public Property varTsListe As SortedList(Of String, clsVarTs)
    Public Sub New()
        _VPsN = New SortedList(Of String, clsVP)
        _VPsId = New SortedList(Of String, clsVP)
        _VPvs = New SortedList(Of String, SortedList(Of String, clsVarTs))
        _VCrole = New SortedList(Of String, clsVCrole)
        _VCcost = New SortedList(Of String, clsVCcost)
        _updateDelay = cacheUpdateDelay
    End Sub

    Public Sub Clear()
        _VPsN.Clear()
        _VPsId.Clear()
        _VPvs.Clear()
        _VCrole.Clear()
        _VCcost.Clear()
        _updateDelay = cacheUpdateDelay
    End Sub


    ''' <summary>
    ''' Cache füllen mit ProjektShortVersions zum Zeitpunkt timeCached
    ''' </summary>
    ''' <param name="result">Liste von KurzProjektVersionen</param>
    ''' <param name="timeCached">Zeitpunkt, zu dem der Cache gefüllt wurde</param>
    Public Sub createVPvShort(ByVal result As List(Of clsProjektWebShort), ByVal timeCached As Date)

        Dim vpid As String = ""
        Dim hvpv As SortedList(Of String, clsVarTs)
        Dim hVarTS As New clsVarTs
        Dim vp As New clsVP
        Try

            For Each vpv As clsProjektWebShort In result

                vp = VPsId(vpv.vpid)
                vpid = vpv.vpid


                If _VPvs.ContainsKey(vpid) Then
                    hvpv = _VPvs(vpid)
                Else
                    hvpv = New SortedList(Of String, clsVarTs)
                End If

                If Not hvpv.ContainsKey(vpv.variantName) Then
                    hVarTS = New clsVarTs
                Else
                    hVarTS = hvpv(vpv.variantName)
                End If

                ' vpv nur eintragen, wenn der timestamp nicht bereits vorhanden
                If Not hVarTS.tsShort.ContainsKey(vpv.timestamp) Then
                    hVarTS.vname = vpv.variantName
                    ' Anzahl Versionen dieser Variante merken in hVarTS
                    If vpv.variantName = "" Then
                        hVarTS.vpvCount = vp.vpvCount
                    Else
                        For Each vpvariant In vp.Variant
                            If vpvariant.variantName = vpv.variantName Then
                                hVarTS.vpvCount = vpvariant.vpvCount
                            End If
                        Next
                    End If
                    hVarTS.timeCShort = timeCached
                    hVarTS.tsShort.Add(vpv.timestamp, vpv)
                End If

                If Not hvpv.ContainsKey(vpv.variantName) Then
                    hvpv.Add(vpv.variantName, hVarTS)
                Else
                    hvpv.Remove(vpv.variantName)
                    hvpv.Add(vpv.variantName, hVarTS)
                End If

                If _VPvs.ContainsKey(vpid) Then
                    _VPvs(vpid) = hvpv
                Else
                    _VPvs.Add(vpid, hvpv)
                End If

            Next

        Catch ex As Exception
            Throw New ArgumentException("Fehler im Caching: createVPvShort")
        End Try
    End Sub

    ''' <summary>
    ''' Cache füllen mit ProjektLongVersions und ShortVersions zum Zeitpunkt timeCached
    ''' </summary>
    ''' <param name="result">Liste von KurzProjektVersionen</param>
    ''' <param name="timeCached">Zeitpunkt, zu dem der Cache gefüllt wurde</param>
    Public Sub createVPvLong(ByVal result As List(Of clsProjektWebLong), Optional ByVal timeCached As Date = Nothing)

        Dim vpid As String = ""
        Dim hvpv As SortedList(Of String, clsVarTs)
        Dim hVarTS As New clsVarTs
        Dim vp As New clsVP

        Try
            For Each vpv As clsProjektWebLong In result

                Dim vpvshort As clsProjektWebShort = vpvLong2vpvshort(vpv)
                vp = VPsId(vpv.vpid)
                vpid = vpv.vpid


                If _VPvs.ContainsKey(vpid) Then
                    hvpv = _VPvs(vpid)
                Else
                    hvpv = New SortedList(Of String, clsVarTs)
                End If


                If Not hvpv.ContainsKey(vpv.variantName) Then
                    hVarTS = New clsVarTs
                Else
                    hVarTS = hvpv(vpv.variantName)
                End If

                ' longVersion in den Cache
                If Not hVarTS.tsLong.ContainsKey(vpv.timestamp) Then
                    hVarTS.vname = vpv.variantName
                    ' Anzahl Versionen dieser Variante merken in hVarTS
                    If vpv.variantName = "" Then
                        hVarTS.vpvCount = vp.vpvCount
                    Else
                        For Each vpvariant In vp.Variant
                            If vpvariant.variantName = vpv.variantName Then
                                vpvariant.vpvCount = vpvariant.vpvCount + 1
                                hVarTS.vpvCount = vpvariant.vpvCount
                            End If
                        Next
                    End If
                    ' timeCached soll nicht aktualisiert werden, da Timestamps nicht vollständig sind, sondern nur einzelne dazukamen
                    If timeCached > Date.MinValue Then
                        hVarTS.timeCLong = timeCached
                    End If
                    hVarTS.tsLong.Add(vpv.timestamp, vpv)
                End If

                ' gleichzeitig auch die shortVersion cachen 
                If Not hVarTS.tsShort.ContainsKey(vpvshort.timestamp) Then
                    hVarTS.vname = vpvshort.variantName
                    ' timeCached soll nicht aktualisiert werden, da Timestamps nicht vollständig sind, sondern nur einzelne dazukamen
                    If timeCached > Date.MinValue Then
                        hVarTS.timeCShort = timeCached
                    End If
                    hVarTS.tsShort.Add(vpvshort.timestamp, vpvshort)
                End If

                If Not hvpv.ContainsKey(vpv.variantName) Then
                    hvpv.Add(vpv.variantName, hVarTS)
                End If

                If _VPvs.ContainsKey(vpid) Then
                    _VPvs(vpid) = hvpv
                Else
                    _VPvs.Add(vpid, hvpv)
                End If
            Next


        Catch ex As Exception
            Throw New ArgumentException("Fehler im Caching: createVPvLong")
        End Try

    End Sub
    Public Sub deleteVPv(ByVal vpvid As String)

        Dim vpid As String
        Dim vname As String
        Dim found As Boolean = False

        If _VPvs.Count > 0 Then

            While Not found

                For Each kvp As KeyValuePair(Of String, SortedList(Of String, clsVarTs)) In _VPvs

                    vpid = kvp.Key
                    Dim VPvs_value As SortedList(Of String, clsVarTs) = _VPvs(vpid)

                    If VPvs_value.Count <> 0 Then

                        Dim varTS As SortedList(Of String, clsVarTs) = _VPvs(vpid)
                        For Each kvp1 As KeyValuePair(Of String, clsVarTs) In varTS
                            vname = kvp1.Key
                            Dim vpvlongListe As SortedList(Of Date, clsProjektWebLong) = kvp1.Value.tsLong
                            For Each vpvlong As KeyValuePair(Of Date, clsProjektWebLong) In vpvlongListe
                                If vpvlong.Value._id = vpvid Then
                                    vpvlongListe.Remove(vpvlong.Key)
                                    found = True
                                    Exit For
                                End If
                            Next
                            Dim vpvshortListe As SortedList(Of Date, clsProjektWebShort) = kvp1.Value.tsShort
                            For Each vpvshort As KeyValuePair(Of Date, clsProjektWebShort) In vpvshortListe
                                If vpvshort.Value._id = vpvid Then
                                    vpvshortListe.Remove(vpvshort.Key)
                                    found = True
                                    Exit For
                                End If
                            Next
                            If found Then
                                Exit For
                            End If
                        Next

                    End If

                    If found Then
                        Exit For
                    End If
                Next

            End While

        End If
    End Sub

    Public Function existsInCache(ByVal vpid As String,
                                  ByVal vName As String,
                                  Optional ByVal vpvid As String = "",
                                  Optional ByVal longVersion As Boolean = False,
                                  Optional ByVal refDate As Date = Nothing) As Boolean

        Dim nothingToDo As Boolean = False
        Dim timeDiff As Long = 0

        Try

            If vpid <> "" Then

                If _VPvs.ContainsKey(vpid) Then

                    ' Anzahl Variante > 0
                    If _VPvs(vpid).Count > 0 Then

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

                            If vName <> "" And vName <> noVariantName Then

                                If _VPvs(vpid).ContainsKey(vName) Then

                                    ' nachsehen, ob im Cache für Projekt vpid die Variante variantName und ihre Timestamps gespeichert sind, 
                                    ' wenn ja, dann result-liste aufbauen


                                    If Not longVersion Then

                                        timeDiff = DateDiff(DateInterval.Minute, _VPvs(vpid)(vName).timeCShort, Date.Now.ToUniversalTime)
                                        If _VPvs(vpid)(vName).tsShort.Count = _VPvs(vpid)(vName).vpvCount And
                                            timeDiff <= updateDelay Then

                                            nothingToDo = True

                                        Else
                                            nothingToDo = False
                                        End If
                                    Else

                                        timeDiff = DateDiff(DateInterval.Minute, _VPvs(vpid)(vName).timeCLong, Date.Now.ToUniversalTime)
                                        If _VPvs(vpid)(vName).tsLong.Count = _VPvs(vpid)(vName).vpvCount And
                                           timeDiff <= updateDelay Then

                                            nothingToDo = True

                                        Else
                                            nothingToDo = False
                                        End If
                                    End If
                                Else
                                    nothingToDo = False

                                End If


                            Else  ' von if vname <> "" and vname <> novariantname

                                ' nachsehen, ob im Cache für Projekt vpid alle Variante und Timestamps gespeichert sind, 
                                ' wenn ja, dann result-liste aufbauen
                                ' tk 5.5. soll / kann hier eine if _VPsID.containskey(vpid) rein ,
                                ' wieso kann das überhaupt sein, wo doch VPvs.(vpid).count > 0 ? 
                                Dim vp As clsVP = _VPsId(vpid)

                                ' VisboProjekt Standard, keine Variante (Variante = "")
                                If vName <> noVariantName Then

                                    If _VPvs(vpid).ContainsKey(vName) Then

                                        If Not longVersion Then

                                            timeDiff = DateDiff(DateInterval.Minute, _VPvs(vpid)(vName).timeCShort, Date.Now.ToUniversalTime)

                                            If (_VPvs(vpid)(vName).tsShort.Count = _VPvs(vpid)(vName).vpvCount) And
                                            (_VPvs(vpid)(vName).tsShort.Count >= _VPvs(vpid)(vName).tsLong.Count) And
                                            timeDiff <= updateDelay Then

                                                nothingToDo = True
                                            Else
                                                nothingToDo = False

                                            End If
                                        Else

                                            timeDiff = DateDiff(DateInterval.Minute, _VPvs(vpid)(vName).timeCLong, Date.Now.ToUniversalTime)

                                            If (_VPvs(vpid)(vName).tsLong.Count = _VPvs(vpid)(vName).vpvCount) And
                                            (_VPvs(vpid)(vName).tsLong.Count = _VPvs(vpid)(vName).tsShort.Count) And
                                            timeDiff <= updateDelay Then

                                                If refDate <= Date.MinValue Then
                                                    ' kein refDate angegeben
                                                    nothingToDo = True
                                                Else
                                                    nothingToDo = False
                                                End If

                                            Else

                                                nothingToDo = False

                                            End If
                                        End If
                                    End If

                                Else   ' vname = noVariantname, alle Varianten sind relevant


                                    For Each vpvar As clsVPvariant In vp.Variant
                                        Try
                                            If _VPvs(vpid).ContainsKey(vpvar.variantName) Then

                                                If Not longVersion Then
                                                    timeDiff = DateDiff(DateInterval.Minute, _VPvs(vpid)(vpvar.variantName).timeCShort, Date.Now.ToUniversalTime)
                                                    If (_VPvs(vpid)(vpvar.variantName).tsShort.Count = _VPvs(vpid)(vpvar.variantName).vpvCount) And
                                                        (_VPvs(vpid)(vpvar.variantName).tsShort.Count >= _VPvs(vpid)(vpvar.variantName).tsLong.Count) And
                                                         timeDiff <= updateDelay Then

                                                    Else

                                                        nothingToDo = nothingToDo And False
                                                        Exit For

                                                    End If
                                                Else

                                                    timeDiff = DateDiff(DateInterval.Minute, _VPvs(vpid)(vName).timeCLong, Date.Now.ToUniversalTime)
                                                    If (_VPvs(vpid)(vpvar.variantName).tsLong.Count = _VPvs(vpid)(vpvar.variantName).vpvCount) And
                                                        (_VPvs(vpid)(vpvar.variantName).tsLong.Count = _VPvs(vpid)(vpvar.variantName).tsShort.Count) And
                                                        timeDiff <= updateDelay Then

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


                                End If ' end if von vName <> noVariantName



                            End If    ' end if von vName <> ""

                        End If   ' end if von if vps_id


                    Else
                        nothingToDo = nothingToDo And False

                    End If  ' end if von if _vpvs(vpid).count > 0


                Else
                    nothingToDo = nothingToDo And False

                End If

            Else   ' hier ist vpid = ""

                Dim ok As Boolean = True
                If _VPvs.Count > 0 Then

                    For Each kvp As KeyValuePair(Of String, SortedList(Of String, clsVarTs)) In _VPvs

                        vpid = kvp.Key

                        If _VPvs(vpid).Count > 0 Then

                            Dim vp As clsVP = _VPsId(vpid)

                            If vName <> noVariantName Then

                                If _VPvs(vpid).ContainsKey(vName) Then

                                    If Not longVersion Then
                                        timeDiff = DateDiff(DateInterval.Minute, _VPvs(vpid)(vName).timeCShort, Date.Now.ToUniversalTime)
                                        If (_VPvs(vpid)(vName).tsShort.Count > 0) And
                                            (_VPvs(vpid)(vName).tsShort.Count >= _VPvs(vpid)(vName).tsLong.Count) And
                                            timeDiff <= updateDelay Then

                                            ok = True
                                        Else

                                            ok = False


                                        End If
                                    Else
                                        timeDiff = DateDiff(DateInterval.Minute, _VPvs(vpid)(vName).timeCLong, Date.Now.ToUniversalTime)
                                        If (_VPvs(vpid)(vName).tsLong.Count > 0) And
                                            (_VPvs(vpid)(vName).tsLong.Count = _VPvs(vpid)(vName).tsShort.Count) And
                                            timeDiff <= updateDelay Then

                                            ok = True
                                        Else
                                            ok = False
                                        End If
                                    End If
                                Else
                                    ok = False
                                End If

                            Else   ' vname = noVariantname, alle Varianten sind relevant


                                For Each vpvar As clsVPvariant In vp.Variant
                                    Try
                                        If _VPvs(vpid).ContainsKey(vpvar.variantName) Then

                                            If Not longVersion Then
                                                timeDiff = DateDiff(DateInterval.Minute, _VPvs(vpid)(vpvar.variantName).timeCShort, Date.Now.ToUniversalTime)
                                                If (_VPvs(vpid)(vpvar.variantName).tsShort.Count > 0) And
                                                (_VPvs(vpid)(vpvar.variantName).tsShort.Count >= _VPvs(vpid)(vpvar.variantName).tsLong.Count) And
                                                 timeDiff <= updateDelay Then

                                                    ok = ok And True
                                                Else

                                                    ok = False
                                                    Exit For

                                                End If
                                            Else

                                                timeDiff = DateDiff(DateInterval.Minute, _VPvs(vpid)(vName).timeCLong, Date.Now.ToUniversalTime)
                                                If (_VPvs(vpid)(vpvar.variantName).tsLong.Count > 0) And
                                                (_VPvs(vpid)(vpvar.variantName).tsLong.Count = _VPvs(vpid)(vpvar.variantName).tsShort.Count) And
                                                timeDiff <= updateDelay Then

                                                    ok = ok And True
                                                Else

                                                    ok = False
                                                    Exit For
                                                End If
                                            End If

                                        Else
                                            ok = False
                                            Exit For

                                        End If
                                    Catch ex As Exception

                                    End Try

                                Next

                            End If ' end if von vName <> noVariantName

                        Else

                            ok = False

                        End If

                        If ok = False Then

                            Exit For

                        End If

                    Next

                Else
                    ok = False
                End If

                nothingToDo = ok

            End If

        Catch ex As Exception
            nothingToDo = False
        End Try

        existsInCache = nothingToDo

    End Function
    ''' <summary>
    ''' wandelt ein Projekt der Longversion in eines der shortversion
    ''' </summary>
    ''' <param name="vpvL"></param>
    ''' <returns></returns>
    Private Function vpvLong2vpvshort(ByVal vpvL As clsProjektWebLong) As clsProjektWebShort

        Dim vpvshort As New clsProjektWebShort
        Try

            vpvshort._id = vpvL._id
            vpvshort.name = vpvL.name
            vpvshort.vpid = vpvL.vpid
            vpvshort.timestamp = vpvL.timestamp
            vpvshort.Erloes = vpvL.Erloes
            vpvshort.startDate = vpvL.startDate
            vpvshort.endDate = vpvL.endDate
            vpvshort.status = vpvL.status
            vpvshort.variantName = vpvL.variantName
            vpvshort.ampelStatus = vpvL.ampelStatus

        Catch ex As Exception
            vpvLong2vpvshort = Nothing
        End Try

        vpvLong2vpvshort = vpvshort

    End Function
End Class

