Imports ProjectBoardDefinitions
Imports MongoDbAccess
Imports WebServerAcc

''' <summary>
''' diese Klasse wird als eine Instanz-Variable angesprochen 
''' sie gilt für alle geöffneten Powerpoint Dateien - bzw. je eine Instanz-Variable für ein VISBO Center 
''' </summary>
Public Class clsPPTTimeMachine

    Private _visboCenter As String

    ' die Liste der Projekte mit ihren versionierten Ständen 
    Private _projectTimeStamps As SortedList(Of String, clsProjektHistorie)

    ' enthält aktuell immer nur zwei Datums - nämlich das kleineste auftretende und das größe auftretende 
    Private _minmaxTimeStamps(1) As Date


    ''' <summary>
    ''' true, wenn das Projekt bereits in der Liste enthalten ist 
    ''' false, sonst
    ''' </summary>
    ''' <param name="pVName"></param>
    ''' <returns></returns>
    Public ReadOnly Property containsProject(ByVal pVName As String) As Boolean
        Get
            containsProject = _projectTimeStamps.ContainsKey(pVName)
        End Get
    End Property

    Public ReadOnly Property count() As Integer
        Get
            count = _projectTimeStamps.Count
        End Get
    End Property

    ''' <summary>
    ''' passt den kleinsten und größten vorkommenden TS Wert so an wie es den auf der Seite vorhandenen Projekten entspricht
    ''' die sind in der smartSlideListe enthalten 
    ''' </summary>

    Private Sub adjustMinMaxToCurrentSlide()

        _minmaxTimeStamps(0) = Date.Now
        _minmaxTimeStamps(1) = Date.MinValue

        Dim defaultSettingNecessary As Boolean = True


        If smartSlideLists.countProjects > 0 Then
            ' es gibt Projekte , also anpassen 


            For i As Integer = 1 To smartSlideLists.countProjects
                Dim pvName As String = smartSlideLists.getPVName(i)
                If _projectTimeStamps.ContainsKey(pvName) Then
                    Dim phistory As clsProjektHistorie = _projectTimeStamps.Item(pvName)

                    If Not IsNothing(phistory) Then
                        If phistory.Count > 0 Then

                            Dim minTS As Date = phistory.First.timeStamp
                            Dim maxTs As Date = phistory.Last.timeStamp

                            If _minmaxTimeStamps(0) > minTS Then
                                _minmaxTimeStamps(0) = minTS
                            End If

                            If _minmaxTimeStamps(1) < maxTs Then
                                _minmaxTimeStamps(1) = maxTs
                            End If

                            defaultSettingNecessary = False

                        End If

                    End If
                End If

            Next
        Else
            ' es muss nichts weiter getan werden, minmax Werte sind bereits zurückgesetzt 
        End If


        If smartSlideLists.countPortfolios > 0 Then

            ' es gibt Portfolios, also anpassen 

            Dim err As New clsErrorCodeMsg
            Dim minTS As Date = Date.Now

            For i As Integer = 1 To smartSlideLists.countPortfolios


                Dim pfName As String = smartSlideLists.getPfName(i)

                Dim portfolio As clsConstellation = CType(databaseAcc, DBAccLayer.Request).retrieveFirstVersionOfOneConstellationFromDB(pfName,
                                                                                                                                        minTS, err,
                                                                                                                                        Date.MinValue.AddDays(1))

                If _minmaxTimeStamps(0) > minTS Then
                    _minmaxTimeStamps(0) = minTS.AddSeconds(1)
                End If


                defaultSettingNecessary = False

                '        End If

                '    End If
                'End If

            Next
        Else
            ' es muss nichts weiter getan werden, minmax Werte sind bereits zurückgesetzt 
        End If


        If defaultSettingNecessary Then
            _minmaxTimeStamps(0) = Date.Now.Date
            _minmaxTimeStamps(1) = _minmaxTimeStamps(0).AddHours(23).AddMinutes(59)
        End If


    End Sub



    ''' <summary>
    ''' fügt ein Projekt incl Nothing-ProjektHistorie als Platzhalter hinzu 
    ''' </summary>
    ''' <param name="pvName"></param>
    Public Sub addProject(ByVal pvName As String)

        If Not _projectTimeStamps.ContainsKey(pvName) Then
            _projectTimeStamps.Add(pvName, Nothing)
        Else
            ' nichts tun , ist schon drin ... 
        End If

    End Sub

    Public ReadOnly Property historyExists(ByVal pvName As String) As Boolean
        Get
            Dim tmpResult As Boolean
            If Not _projectTimeStamps.ContainsKey(pvName) Then
                tmpResult = False
            Else
                tmpResult = Not IsNothing(_projectTimeStamps.Item(pvName))
            End If
            historyExists = tmpResult
        End Get
    End Property

    Public ReadOnly Property getProjectVersion(ByVal pvName As String, ByVal refDate As Date) As clsProjekt
        Get
            Dim result As clsProjekt = Nothing
            Dim err As New clsErrorCodeMsg

            If _projectTimeStamps.ContainsKey(pvName) Then
                Dim myHistory As clsProjektHistorie = _projectTimeStamps.Item(pvName)

                If IsNothing(myHistory) Then

                    ' ProjektHistorie von Cache oder Datenbank holen 
                    Dim pName As String = getPnameFromKey(pvName)
                    Dim vName As String = getVariantnameFromKey(pvName)

                    _projectTimeStamps.Item(pvName) = CType(databaseAcc, DBAccLayer.Request).retrieveProjectHistoryFromDB(pName, vName, Date.MinValue, Date.Now, err)
                    myHistory = _projectTimeStamps.Item(pvName)

                End If

                ' jetzt sollte spätestens die ProjektHistorie gesetzt sein 
                If Not IsNothing(myHistory) Then
                    result = myHistory.ElementAtorBefore(refDate)
                End If
            End If

            getProjectVersion = result

        End Get
    End Property

    Public ReadOnly Property getFirstContractedVersion(ByVal pvname As String) As clsProjekt
        Get
            Dim result As clsProjekt = Nothing
            Dim err As New clsErrorCodeMsg

            If _projectTimeStamps.ContainsKey(pvname) Then
                Dim myHistory As clsProjektHistorie = _projectTimeStamps.Item(pvname)

                If IsNothing(myHistory) Then
                    ' von Cache oder Datenbank holen 
                    Dim pName As String = getPnameFromKey(pvname)
                    Dim vName As String = getVariantnameFromKey(pvname)

                    _projectTimeStamps.Item(pvname) = CType(databaseAcc, DBAccLayer.Request).retrieveProjectHistoryFromDB(pName, vName, Date.MinValue, Date.Now, err)
                    myHistory = _projectTimeStamps.Item(pvname)

                End If

                ' jetzt sollte spätestens die ProjektHistorie gesetzt sein 
                If Not IsNothing(myHistory) Then
                    result = myHistory.beauftragung
                End If
            End If

            getFirstContractedVersion = result
        End Get
    End Property

    Public ReadOnly Property getLastContractedVersion(ByVal pvName As String, ByVal refdate As Date) As clsProjekt
        Get
            Dim result As clsProjekt = Nothing
            Dim err As New clsErrorCodeMsg

            If _projectTimeStamps.ContainsKey(pvName) Then
                Dim myHistory As clsProjektHistorie = _projectTimeStamps.Item(pvName)

                If IsNothing(myHistory) Then
                    ' von Cache oder Datenbank holen 
                    Dim pName As String = getPnameFromKey(pvName)
                    Dim vName As String = getVariantnameFromKey(pvName)

                    _projectTimeStamps.Item(pvName) = CType(databaseAcc, DBAccLayer.Request).retrieveProjectHistoryFromDB(pName, vName, Date.MinValue, Date.Now, err)
                    myHistory = _projectTimeStamps.Item(pvName)

                End If

                ' jetzt sollte spätestens die ProjektHistorie gesetzt sein 
                If Not IsNothing(myHistory) Then
                    result = myHistory.lastBeauftragung(refdate)
                End If
            End If

            getLastContractedVersion = result
        End Get
    End Property

    ''' <summary>
    ''' gibt das Datum zurück, das eingestellt wird, wenn der TimeMachine Button gedrückt wird ... 
    ''' 
    ''' </summary>
    ''' <param name="kennung"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getNextNavigationDate(ByVal kennung As Integer,
                                          ByVal specDate As Date,
                                          Optional ByVal justForInformation As Boolean = False
                                          ) As Date

        'Dim key As String = CType(currentSlide.Parent, PowerPoint.Presentation).Name
        Dim anzahlShapesOnSlide As Integer = currentSlide.Shapes.Count
        Dim err As New clsErrorCodeMsg
        Dim tmpDate As Date = Date.Now

        Dim jetzt As Date = Date.Now

        If smartSlideLists.countProjects > 0 Then
            'Dim tmpTM As clsPPTTimeMachine = Nothing
            If _projectTimeStamps.Count > 0 Then

                ' jetzt prüfen, ob es wenigstens schon eine erste Festlegung der minmax-Werte gegeben hat
                ' der Hinweis, dass noch keine Zuordnung stattgefunden hat, ist wenn minmaxTimestamps(1) = Date.mimvalue ist 
                If _minmaxTimeStamps(1) = Date.MinValue Then
                    ' jetzt für alle Projekte, wo clsProjektHistorie noch Nothing ist die Historie holen 
                    For i As Integer = 0 To _projectTimeStamps.Count - 1

                        Dim pHistory As clsProjektHistorie = _projectTimeStamps.ElementAt(i).Value
                        Dim key As String = _projectTimeStamps.ElementAt(i).Key
                        Dim pName As String = getPnameFromKey(key)
                        Dim vName As String = getVariantnameFromKey(key)

                        If IsNothing(pHistory) Then

                            _projectTimeStamps.Item(key) = CType(databaseAcc, DBAccLayer.Request).retrieveProjectHistoryFromDB(pName, vName, Date.MinValue, Date.Now, err)

                        ElseIf pHistory.Count = 0 Then

                            _projectTimeStamps.Item(key) = CType(databaseAcc, DBAccLayer.Request).retrieveProjectHistoryFromDB(pName, vName, Date.MinValue, Date.Now, err)

                        End If

                    Next
                End If

            End If

        End If


        ' jetzt müssen minmax-Werte an die aktuelle Slide angepasst werden 
        ' jetzt noch minmax-Timestamps anpassen 
        Call adjustMinMaxToCurrentSlide()

        Select Case kennung
                Case ptNavigationButtons.nachher


                    If currentTimestamp.AddMonths(1) <= jetzt Then
                        tmpDate = currentTimestamp.AddMonths(1)
                    Else
                        tmpDate = jetzt
                    End If

                    tmpDate = tmpDate.Date.AddHours(23).AddMinutes(59)


                Case ptNavigationButtons.vorher

                    If currentTimestamp.AddMonths(-1) > _minmaxTimeStamps(0) Then
                        tmpDate = currentTimestamp.AddMonths(-1)
                    Else
                        tmpDate = _minmaxTimeStamps(0)
                    End If

                    tmpDate = tmpDate.Date.AddHours(23).AddMinutes(59)


                Case ptNavigationButtons.erster

                    tmpDate = _minmaxTimeStamps(0)


                Case ptNavigationButtons.letzter

                    'ur: 20190513: letzter Stand ist gleich Stand jetzt
                    'tmpDate = _minmaxTimeStamps(1)
                    tmpDate = jetzt
                    tmpDate = tmpDate.Date.AddHours(23).AddMinutes(59)

                Case ptNavigationButtons.update

                    tmpDate = jetzt
                    tmpDate = tmpDate.Date.AddHours(23).AddMinutes(59)

                Case ptNavigationButtons.individual

                    If specDate > _minmaxTimeStamps(0) And specDate < jetzt Then
                        tmpDate = specDate
                    Else
                        If specDate > jetzt Then
                            tmpDate = jetzt
                        ElseIf specDate < _minmaxTimeStamps(0) Then
                            tmpDate = _minmaxTimeStamps(0)
                        End If

                        tmpDate = tmpDate.Date.AddHours(23).AddMinutes(59)
                    End If



                Case ptNavigationButtons.previous

                    If smartSlideLists.prevDate >= _minmaxTimeStamps(0) Then
                        tmpDate = smartSlideLists.prevDate
                    Else
                        tmpDate = _minmaxTimeStamps(0)
                        tmpDate = tmpDate.Date.AddHours(23).AddMinutes(59)
                    End If



            End Select

        'If Not justForInformation Then
        '    tmpTM.timeStampsIndex = tmpIndex
        'End If

        'Else
        '    ' nichts tun ...
        '    tmpDate = jetzt
        'End If



        getNextNavigationDate = tmpDate
    End Function


    Public Sub New()
        _minmaxTimeStamps(0) = Date.Now
        _minmaxTimeStamps(1) = Date.MinValue
        _projectTimeStamps = New SortedList(Of String, clsProjektHistorie)
        _visboCenter = ""
    End Sub

    '''' <summary>
    '''' Initialisieren der Time-Machine
    '''' </summary>
    '''' <remarks></remarks>
    'Public Sub initPPTTimeMachine(Optional ByVal showMessage As Boolean = True)

    '    Dim err As New clsErrorCodeMsg

    '    Dim msg As String = ""
    '    Dim key As String = CType(currentSlide.Parent, PowerPoint.Presentation).Name

    '    Dim tsCollection As New Collection

    '    If userIsEntitled(msg) Then
    '        ' prüfen, ob es eine Smart Slide ist und ob die Projekt-Historien bereits geladen sind ...
    '        If smartSlideLists.countProjects > 0 Then

    '            ' muss noch eingeloggt werden ? 
    '            ' das wird ja schon im userISEntitled gemacht 
    '            'If noDBAccessInPPT Then

    '            '    noDBAccessInPPT = Not logInToMongoDB(True)

    '            '    If noDBAccessInPPT Then
    '            '        If englishLanguage Then
    '            '            msg = "no database access ... "
    '            '        Else
    '            '            msg = "kein Datenbank Zugriff ... "
    '            '        End If
    '            '        Call MsgBox(msg)
    '            '    Else
    '            '        ' hier müssen jetzt die Role- & Cost-Definitions gelesen werden 

    '            '        RoleDefinitions = CType(databaseAcc, DBAccLayer.Request).retrieveRolesFromDB(Date.Now)
    '            '        CostDefinitions = CType(databaseAcc, DBAccLayer.Request).retrieveCostsFromDB(Date.Now)

    '            '    End If

    '            'End If

    '            If Not noDBAccessInPPT Then

    '                If Not smartSlideLists.historiesExist Then


    '                    Dim anzahlProjekte As Integer = smartSlideLists.countProjects
    '                    ' größter kleinster Wert 
    '                    Dim gkw As Date = Date.MinValue

    '                    For i As Integer = 1 To anzahlProjekte
    '                        Dim tmpName As String = smartSlideLists.getPVName(i)
    '                        Dim pName As String = getPnameFromKey(tmpName)
    '                        Dim vName As String = getVariantnameFromKey(tmpName)
    '                        Dim pvName As String = calcProjektKeyDB(pName, vName)

    '                        tsCollection = CType(databaseAcc, DBAccLayer.Request).retrieveZeitstempelFirstLastFromDB(pvName, err)
    '                        ' ermitteln des größten kleinstern Wertes ...
    '                        ' stellt sicher, dass , wenn mehrere Projekte dargesteltl sind, nur TimeStamps abgerufen werden, die jedes Projekt hat ... 

    '                        ' tk 28.10.18 das ist kontraproduktiv ...weil damit Timestamps rausfliegen, die die nicht in jedem Projekt liegen 
    '                        'Dim kleinsterWert As Date = Date.Now
    '                        'If Not IsNothing(tsCollection) Then
    '                        '    If tsCollection.Count > 0 Then
    '                        '        ' tsCollection ist absteigend sortiert ... 
    '                        '        kleinsterWert = tsCollection.Item(tsCollection.Count)
    '                        '    End If
    '                        'End If
    '                        'If kleinsterWert > gkw Then
    '                        '    gkw = kleinsterWert
    '                        'End If

    '                        smartSlideLists.addToListOfTS(tsCollection)
    '                    Next

    '                    ' tk 28.10.18 keine Reduzierung mehr .. 
    '                    'If anzahlProjekte > 1 Then
    '                    '    ' jetzt werden aus der TimeStampListe alle TimeStamps rausgeworfen, die kleiner als der gkw sind ... 
    '                    '    smartSlideLists.adjustListOfTS(gkw)
    '                    'End If

    '                End If

    '                ' jetzt wird die varPPTTM aufgebaut bzw. erweitert - sie darf nicht gelöscht werden

    '                If IsNothing(tmpTM) Then
    '                    tmpTM = New clsPPTTimeMachine
    '                    tmpTM.timeStamps = smartSlideLists.getListOfTS
    '                Else
    '                    ' sie so übernehmen wie sie ist ... 
    '                    If tsCollection.Count > 0 Then
    '                        tmpTM.addNewList(tsCollection)
    '                    End If
    '                End If

    '                ' tk 28.10.18 nicht mehr nötig ..
    '                ' -------------------------------------------------------------------------------------------------------------------------
    '                ' ab hier war es der Load des Formulars
    '                ' -------------------------------------------------------------------------------------------------------------------------


    '                ' '' '' jetzt wird das Formular TimeStamps aufgerufen ...
    '                '' ''Dim tmFormular As New frmPPTTimeMachine
    '                '' ''Dim dgRes As Windows.Forms.DialogResult = tmFormular.ShowDialog
    '                ' '' ''tmFormular.Show()



    '                'Dim currentDate As Date = Date.Now

    '                '    ' die MArker, falls welche sichtbar sind , wegmachen ... 
    '                '    Call deleteMarkerShapes()

    '                '    'currentTSIndex = -1
    '                '    ' gibt es ein Creation Date ?
    '                '    If smartSlideLists.creationDate > Date.MinValue Then
    '                '        currentDate = currentTimestamp
    '                '    Else
    '                '        currentDate = Date.MinValue
    '                '    End If

    '                '    If noDBAccessInPPT Then
    '                '        Call MsgBox("no Database Access  ... action cancelled ...")
    '                '        'MyBase.Close()-Do nothing    
    '                '    Else
    '                '        ' gibt es überhaupt TimeStamps ? 
    '                '        Try
    '                '            If tsCollection.Count > 0 Then
    '                '                varPPTTM.addNewList(tsCollection)
    '                '            End If
    '                '        Catch ex As Exception

    '                '        End Try



    '                '    If Not IsNothing(varPPTTM.timeStamps) Then
    '                '            If varPPTTM.timeStamps.Count >= 1 Then

    '                '                ' bestimme hier aufgrund des Datums den timestampsIndex
    '                '                If varPPTTM.timeStamps.Count > 0 Then
    '                '                    If smartSlideLists.countProjects = 1 Then
    '                '                        ' nimm das Datum, das in der sortierten Liste unmittelbar davor liegt 
    '                '                        Dim ix As Integer = varPPTTM.timeStamps.Count - 1
    '                '                        Dim found As Boolean = False
    '                '                        Do While ix >= 0 And Not found
    '                '                            If currentTimestamp >= varPPTTM.timeStamps.ElementAt(ix).Key Then
    '                '                                found = True
    '                '                            Else
    '                '                                ix = ix - 1
    '                '                            End If
    '                '                        Loop

    '                '                        If found Then
    '                '                            varPPTTM.timeStampsIndex = ix
    '                '                        End If
    '                '                    Else
    '                '                        ' ist ja schon gesetzt 
    '                '                    End If
    '                '                End If


    '                '                'lblMessage.Text = ""
    '                '                'Me.Text = "Time-Machine: " & timeStamps.First.Key.ToShortDateString & " - " & _
    '                '                '    timeStamps.Last.Key.ToShortDateString & " (" & timeStamps.Count.ToString & ")"

    '                '            Else

    '                '                currentDate = Date.MinValue

    '                '                'lblMessage.Text = "keine Einträge in der Datenbank vorhanden !"
    '                '                'Me.Text = "Time-Machine: "
    '                '            End If
    '                '        End If

    '                '        '' die beiden Buttons Home und ChangedPosition invisible setzen ..
    '                '        'Call setBtnEnablements()

    '            End If


    '        Else

    '            If showMessage Then
    '                Call MsgBox("es gibt auf dieser Seite keine Datenbank-relevanten Informationen ...")
    '            End If

    '        End If
    '    Else
    '        Call MsgBox(msg)
    '    End If



    'End Sub
End Class
