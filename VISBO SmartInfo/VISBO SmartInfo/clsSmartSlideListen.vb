Imports ProjectBoardDefinitions
Imports MongoDbAccess
''' <summary>
''' baut die SmartListen für die betreffende Slide auf
''' dazu gehören classifiedName, OriginalNames, ShortNames, FullBreadCrumbs, ampelColr, 
''' Deliverables, movedElements und Project TimeStamps 
''' die Project TimeStamps werden erstmal für jedes Projekt erst mal nur mit Nothing angelegt, 
''' erst wenn TimeMachine aktiviert wird werden sie nach Bedarf geholt ...
''' </summary>
''' <remarks></remarks>
Public Class clsSmartSlideListen

    ' um zu verhindern, dass der Speicherbedarf wegen sortierter String Listen sehr groß wird, 
    ' wird eine Hilfsliste eingeführt, die für jeden auftretenden Shape-Namen (eindeutig !) eine eindeutige lfdNr zuweist 
    Private _planShapeIDs As SortedList(Of String, Integer)
    Private _IDplanShapes As SortedList(Of Integer, String)

    Private _cNList As SortedList(Of String, SortedList(Of Integer, Boolean))
    ' enthält die Liste der Original Namen 
    Private _oNList As SortedList(Of String, SortedList(Of Integer, Boolean))
    ' enthält die Liste der ShortNames
    Private _sNList As SortedList(Of String, SortedList(Of Integer, Boolean))
    ' enthält die Liste der full BreadCrumbs 
    Private _bCList As SortedList(Of String, SortedList(Of Integer, Boolean))
    ' enthält die Liste der Elemente, die keine, eine grüne, gelbe, rote Bewertung haben 
    Private _aCList As SortedList(Of Integer, SortedList(Of Integer, Boolean))
    ' enthält die Liste der Lieferumfänge; ein Lieferumfang kann ggf in mehreren Elementen vorkommen 
    Private _LUList As SortedList(Of String, SortedList(Of Integer, Boolean))
    ' enthält die Liste der Elemente, die manuell verschoben wurden ... 
    Private _mVList As SortedList(Of Integer, Boolean)
    ' enthält die Liste an Projekt-Historien 
    Private _projectTimeStamps As SortedList(Of String, clsProjektHistorie)
    ' enthält die Liste an TimeStamps, die in der Time-Machine betrachtet werden können 
    ' der bool'sche Wert kann später dafür sorgen, dass ein Eintrag berücksichtigt / nicht berücksichtigt werden soll 
    Private _listOfTimeStamps As SortedList(Of Date, Boolean)

    Private _creationDate As Date

    Private _slideDBUrl As String
    Private _slideDBName As String


    ''' <summary>
    ''' entfernt die Moved Information aus 
    ''' </summary>
    ''' <param name="shpName"></param>
    ''' <remarks></remarks>
    Public Sub removeSMLmvInfo(ByVal shpName As String)

        Dim uid As Integer = _planShapeIDs.Item(shpName)
        If _mVList.ContainsKey(uid) Then
            _mVList.Remove(uid)
        End If

    End Sub
    ''' <summary>
    ''' liest bzw. setzt das Creation Date der Slide 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property creationDate As Date
        Get
            creationDate = _creationDate
        End Get
        Set(value As Date)
            If Not IsNothing(value) Then
                _creationDate = value
            Else
                _creationDate = Date.MinValue
            End If

        End Set
    End Property

    Public Property slideDBUrl As String
        Get
            slideDBUrl = _slideDBUrl
        End Get
        Set(value As String)
            If Not IsNothing(value) Then
                _slideDBUrl = value
            Else
                _slideDBUrl = ""
            End If
        End Set
    End Property

    Public Property slideDBName As String
        Get
            slideDBName = _slideDBName
        End Get
        Set(value As String)
            If Not IsNothing(value) Then
                _slideDBName = value
            Else
                _slideDBName = ""
            End If
        End Set
    End Property

    ''' <summary>
    ''' liefert das TimeStamp Projekt, entweder aus der  smartslideListe oder aus der Datenbank; 
    ''' wenn es aus der DB geholt wird, dann wird es in smartSlideList gechacht
    ''' wenn es auch in de rDatenbank nicht existiert, so wid Nothing zurückgegeben 
    ''' </summary>
    ''' <param name="pvName"></param>
    ''' <param name="tsDate"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getTSProject(ByVal pvName As String, ByVal tsDate As Date) As clsProjekt
        Get
            Dim tmpProject As clsProjekt = Nothing

            Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
            Dim pName As String = getPnameFromKey(pvName)
            Dim vName As String = getVariantnameFromKey(pvName)


            If _projectTimeStamps.ContainsKey(pvName) Then
                Dim timeStamps As clsProjektHistorie = _projectTimeStamps.Item(pvName)
                If Not IsNothing(timeStamps) Then

                    tmpProject = timeStamps.ElementAtorBefore(tsDate)
                    If IsNothing(tmpProject) Then
                        ' aus Datenbank holen 
                        tmpProject = request.retrieveOneProjectfromDB(pName, vName, tsDate)

                        If Not IsNothing(tmpProject) Then
                            timeStamps.Add(tsDate, tmpProject)
                        End If


                    End If
                Else
                    timeStamps = New clsProjektHistorie
                    ' jetzt aus Datenbank holen 
                    tmpProject = request.retrieveOneProjectfromDB(pName, vName, tsDate)

                    If Not IsNothing(tmpProject) Then
                        timeStamps.Add(tsDate, tmpProject)
                    End If

                    _projectTimeStamps.Item(pvName) = timeStamps

                End If

            End If

            getTSProject = tmpProject

        End Get
    End Property
    ''' <summary>
    ''' fügt der Liste an TimeStamps alle Daten, die in einer Collection übergeben werden, hinzu  
    ''' </summary>
    ''' <param name="tsCollection"></param>
    ''' <remarks></remarks>
    Public Sub addToListOfTS(ByVal tsCollection As Collection)

        If Not IsNothing(tsCollection) Then

            Try

                For Each tmpDate As Date In tsCollection
                    If Not _listOfTimeStamps.ContainsKey(tmpDate) Then
                        ' bool'scher Wert hat aktuell keine Bedeutung 
                        ' könnte später bestimmt werden, ob der TimeStamp bereits aus der DB geholt wurde oder nicht .. 
                        _listOfTimeStamps.Add(tmpDate, False)
                    End If
                Next

            Catch ex As Exception
                Exit Sub
            End Try

        End If

    End Sub

    Public ReadOnly Property getListOfTS As SortedList(Of Date, Boolean)
        Get
            getListOfTS = _listOfTimeStamps
        End Get
    End Property

    ''' <summary>
    ''' gibt die Gesamt-Liste aller TimeStamps für das Time-Machine Formular zurück 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getArrayOfTS As Date()
        Get
            Dim tmpArray() As Date = Nothing

            If Not IsNothing(_listOfTimeStamps) Then
                If _listOfTimeStamps.Count > 0 Then
                    tmpArray = _listOfTimeStamps.Keys.ToArray()
                End If
            End If


            'If listOfTimeStamps.Count > 0 Then
            '    ReDim tmpArray(listOfTimeStamps.Count - 1)
            '    Dim index As Integer = 0
            '    For Each kvp As KeyValuePair(Of Date, Boolean) In listOfTimeStamps
            '        tmpArray(index) = kvp.Key
            '    Next
            'End If

            getArrayOfTS = tmpArray

        End Get
    End Property


    ''' <summary>
    ''' liefert true, wenn das Projekt mit projectVariantName = pName#vName in der Liste der Projekte enthalten ist 
    ''' </summary>
    ''' <param name="pvName"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property containsProject(ByVal pvName As String) As Boolean
        Get
            containsProject = _projectTimeStamps.ContainsKey(pvName)
        End Get
    End Property

    ''' <summary>
    ''' gibt den Projekt-Varianten-Namen des i.ten-Elements zurück
    ''' i läuft von 1.. count 
    ''' der Name hat folgenden Aufbau: pName#vName 
    ''' Aufruf mit unzulässigem Index gibt Nothing zurück 
    ''' </summary>
    ''' <param name="index"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getPVName(ByVal index As Integer) As String
        Get

            If index >= 1 And index <= _projectTimeStamps.Count Then
                getPVName = _projectTimeStamps.ElementAt(index - 1).Key
            Else
                getPVName = Nothing
            End If

        End Get
    End Property

    ''' <summary>
    ''' liefert die Anzahl an Projekten, die mit oder ohne TimeStamps aufgeführt sind 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property countProjects() As Integer
        Get
            countProjects = _projectTimeStamps.Count
        End Get
    End Property

    ''' <summary>
    ''' gibt für das angegebene Projekte die Liste der Time-Stamps zurück
    ''' Nothing, wenn sie noch nicht aus der Datenbank geladen wurde  
    ''' </summary>
    ''' <param name="pvName"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getTimeStampListe(ByVal pvName As String) As clsProjektHistorie
        Get
            If _projectTimeStamps.ContainsKey(pvName) Then
                getTimeStampListe = _projectTimeStamps.Item(pvName)
            Else
                getTimeStampListe = Nothing
            End If
        End Get
    End Property

    ''' <summary>
    ''' fügt der Projektliste ein neues Element hinzu; 
    ''' die Project TimeStampListe kann Nothing sein ... 
    ''' </summary>
    ''' <param name="pvName"></param>
    ''' <param name="pHistory"></param>
    ''' <remarks></remarks>
    Public Sub addProject(ByVal pvName As String, Optional ByVal pHistory As clsProjektHistorie = Nothing)

        If _projectTimeStamps.ContainsKey(pvName) Then
            _projectTimeStamps.Remove(pvName)
        End If

        _projectTimeStamps.Add(pvName, pHistory)

    End Sub

    Public ReadOnly Property historiesExist() As Boolean
        Get
            Dim tmpResult As Boolean = True

            Dim i As Integer = 0
            Do While i <= _projectTimeStamps.Count - 1 And tmpResult
                If IsNothing(_projectTimeStamps.ElementAt(i).Value) Then
                    tmpResult = False
                Else
                    i = i + 1
                End If
            Loop

            historiesExist = tmpResult

        End Get
    End Property
    Public ReadOnly Property getUID(ByVal shapeName As String) As Integer
        Get
            Dim uid As Integer
            If _planShapeIDs.ContainsKey(shapeName) Then
                uid = _planShapeIDs.Item(shapeName)
            Else
                uid = _planShapeIDs.Count + 1
                _planShapeIDs.Add(shapeName, uid)
                _IDplanShapes.Add(uid, shapeName)
            End If

            getUID = uid

        End Get
    End Property

    ''' <summary>
    ''' gibt den ShapeName zurück, der zur UID gehört; 
    ''' </summary>
    ''' <param name="UID"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private ReadOnly Property getShapeNameOfUid(ByVal uid As Integer) As String
        Get
            Dim tmpStr As String = ""
            Dim tmpStrTest As String = ""
            Dim found As Boolean = False
            Dim index As Integer = 0

            tmpStr = _IDplanShapes.Item(uid)

            '' für Testzwecke 
            'Do While index <= planShapeIDs.Count - 1 And Not found
            '    If planShapeIDs.ElementAt(index).Value = uid Then
            '        found = True
            '        tmpStrTest = planShapeIDs.ElementAt(index).Key
            '    Else
            '        index = index + 1
            '    End If

            'Loop

            'If tmpStr <> tmpStrTest Then
            '    Dim a As Integer = 0
            'End If

            getShapeNameOfUid = tmpStr

        End Get
    End Property
    ''' <summary>
    ''' fügt der Liste an classified Names einen weiteren Namen hinzu
    ''' wenn der schon existiert, wird die Liste an shapeNames ergänzt; statt ShapeName wird dessen uid geschrieben  
    ''' </summary>
    ''' <param name="cName"></param>
    ''' <param name="shapeName"></param>
    ''' <remarks></remarks>
    Public Sub addCN(ByVal cName As String, shapeName As String)

        Dim uid As Integer = Me.getUID(shapeName)

        Dim listOfShapeNames As SortedList(Of Integer, Boolean)

        If _cNList.ContainsKey(cName) Then
            listOfShapeNames = _cNList.Item(cName)
            If listOfShapeNames.ContainsKey(uid) Then
                ' nichts tun , ist schon drin ...
            Else
                ' aufnehmen ; der bool'sche Value hat aktuell keine Bedeutung 
                listOfShapeNames.Add(uid, True)
            End If
        Else
            ' dann muss das erste aufgenommen werden 
            listOfShapeNames = New SortedList(Of Integer, Boolean)
            listOfShapeNames.Add(uid, True)
            _cNList.Add(cName, listOfShapeNames)
        End If

    End Sub

    ''' <summary>
    ''' fügt der Liste an original Names einen weiteren Namen hinzu
    ''' wenn der schon existiert, wird die Liste an shapeNames ergänzt; statt ShapeName wird dessen uid geschrieben  
    ''' </summary>
    ''' <param name="oName">original Name</param>
    ''' <param name="shapeName"></param>
    ''' <remarks></remarks>
    Public Sub addON(ByVal oName As String, shapeName As String)

        Dim uid As Integer = Me.getUID(shapeName)

        Dim listOfShapeNames As SortedList(Of Integer, Boolean)

        If _oNList.ContainsKey(oName) Then
            listOfShapeNames = _oNList.Item(oName)
            If listOfShapeNames.ContainsKey(uid) Then
                ' nichts tun , ist schon drin ...
            Else
                ' aufnehmen ; der bool'sche Value hat aktuell keine Bedeutung 
                listOfShapeNames.Add(uid, True)
            End If
        Else
            ' dann muss das erste aufgenommen werden 
            listOfShapeNames = New SortedList(Of Integer, Boolean)
            listOfShapeNames.Add(uid, True)
            _oNList.Add(oName, listOfShapeNames)
        End If

    End Sub

    ''' <summary>
    ''' fügt der Liste an Short Names einen weiteren Namen hinzu
    ''' wenn der leer ist, wird stattdessen die uid genommen 
    ''' wenn der schon existiert, wird die Liste an shapeNames ergänzt; statt ShapeName wird dessen uid geschrieben
    ''' </summary>
    ''' <param name="sName"></param>
    ''' <param name="shapeName"></param>
    ''' <remarks></remarks>
    Public Sub addSN(ByVal sName As String, shapeName As String)


        Dim uid As Integer = Me.getUID(shapeName)
        Dim listOfShapeNames As SortedList(Of Integer, Boolean)

        If IsNothing(sName) Then
            sName = uid.ToString
        ElseIf sName.Trim.Length = 0 Then
            sName = uid.ToString
        End If

        If _sNList.ContainsKey(sName) Then
            listOfShapeNames = _sNList.Item(sName)
            If listOfShapeNames.ContainsKey(uid) Then
                ' nichts tun , ist schon drin ...
            Else
                ' aufnehmen ; der bool'sche Value hat aktuell keine Bedeutung 
                listOfShapeNames.Add(uid, True)
            End If
        Else
            ' dann muss das erste aufgenommen werden 
            listOfShapeNames = New SortedList(Of Integer, Boolean)
            listOfShapeNames.Add(uid, True)
            _sNList.Add(sName, listOfShapeNames)
        End If

    End Sub

    ''' <summary>
    ''' fügt der Liste an BreadCrumbs Names einen weiteren bc hinzu
    ''' wenn der schon existiert, wird die Liste an shapeNames ergänzt; statt ShapeName wird dessen uid geschrieben  
    ''' </summary>
    ''' <param name="bCrumb"></param>
    ''' <param name="shapeName"></param>
    ''' <remarks></remarks>
    Public Sub addBC(ByVal bCrumb As String, shapeName As String)

        Dim uid As Integer = Me.getUID(shapeName)

        Dim fullbCrumb As String = "(" & getPnameFromShpName(shapeName) & ")" & _
            bCrumb.Replace("#", " - ") & getElemNameFromShpName(shapeName)


        Dim listOfShapeNames As SortedList(Of Integer, Boolean)

        If _bCList.ContainsKey(fullbCrumb) Then
            listOfShapeNames = _bCList.Item(fullbCrumb)
            If listOfShapeNames.ContainsKey(uid) Then
                ' nichts tun , ist schon drin ...
            Else
                ' aufnehmen ; der bool'sche Value hat aktuell keine Bedeutung 
                listOfShapeNames.Add(uid, True)
            End If
        Else
            ' dann muss das erste aufgenommen werden 
            listOfShapeNames = New SortedList(Of Integer, Boolean)
            listOfShapeNames.Add(uid, True)
            _bCList.Add(fullbCrumb, listOfShapeNames)
        End If

    End Sub

    ''' <summary>
    ''' fügt der Liste an Lieferumfängen weitere hinzu ;
    ''' übergeben wird der komplette String mit Lieferumfängen, einzelne sind duch # voneinander getrennt 
    ''' </summary>
    ''' <param name="lieferumfaenge"></param>
    ''' <param name="shapeName"></param>
    ''' <remarks></remarks>
    Public Sub addLU(ByVal lieferumfaenge As String, shapeName As String)

        Dim uid As Integer = Me.getUID(shapeName)
        Dim lieferumfang As String
        Dim trennzeichen As String = "#"

        Dim tmpStr() As String = lieferumfaenge.Split(New Char() {CType(trennzeichen, Char)})

        For i As Integer = 1 To tmpStr.Length
            lieferumfang = tmpStr(i - 1)

            Dim listOfShapeIDs As SortedList(Of Integer, Boolean)
            If _LUList.ContainsKey(lieferumfang) Then
                listOfShapeIDs = _LUList.Item(lieferumfang)
                If listOfShapeIDs.ContainsKey(uid) Then
                    ' nichts tun , ist schon drin ...
                Else
                    ' aufnehmen ; der bool'sche Value hat aktuell keine Bedeutung 
                    listOfShapeIDs.Add(uid, True)
                End If
            Else
                ' dann muss das erste aufgenommen werden 
                listOfShapeIDs = New SortedList(Of Integer, Boolean)
                listOfShapeIDs.Add(uid, True)
                _LUList.Add(lieferumfang, listOfShapeIDs)
            End If
        Next

    End Sub

    ''' <summary>
    ''' fügt der Liste an "verschobenen Elementen" ein weiteres hinzu ...
    ''' </summary>
    ''' <param name="shapeName"></param>
    ''' <remarks></remarks>
    Public Sub addMV(ByVal shapeName As String)
        Dim uid As Integer = Me.getUID(shapeName)
        If _mVList.ContainsKey(uid) Then
            ' nichts tun , ist schon drin
        Else
            _mVList.Add(uid, True)
        End If
    End Sub

    ''' <summary>
    ''' fügt der Liste an Ampelfarben eine weitere (0,1,2,3) hinzu
    ''' wenn die schon existiert, wird die Liste an shapeNames ergänzt; statt ShapeName wird dessen uid geschrieben  
    ''' </summary>
    ''' <param name="ampelColor"></param>
    ''' <param name="shapeName"></param>
    ''' <remarks></remarks>
    Public Sub addAC(ByVal ampelColor As Integer, shapeName As String)

        Dim uid As Integer = Me.getUID(shapeName)

        Dim listOfShapeNames As SortedList(Of Integer, Boolean)

        ' konsistent machen ... wenn die Farbe nicht erkannt werden kann, wird sie wie <nicht gesetzt> behandelt 
        If ampelColor < 0 Or ampelColor > 3 Then
            ' nichts tun ... 
        Else
            If _aCList.ContainsKey(ampelColor) Then
                listOfShapeNames = _aCList.Item(ampelColor)
                If listOfShapeNames.ContainsKey(uid) Then
                    ' nichts tun , ist schon drin ...
                Else
                    ' aufnehmen ; der bool'sche Value hat aktuell keine Bedeutung 
                    listOfShapeNames.Add(uid, True)
                End If
            Else
                ' dann muss das erste aufgenommen werden 
                listOfShapeNames = New SortedList(Of Integer, Boolean)
                listOfShapeNames.Add(uid, True)
                _aCList.Add(ampelColor, listOfShapeNames)
            End If
        End If



    End Sub

    ''' <summary>
    ''' gibt für die angegebene Ampelfarbe die Namen alle Shapes zurück, die diese Ampelfarbe haben 
    ''' leere Collection, wenn es keine Shapes dieser Farbe gibt
    ''' </summary>
    ''' <param name="ampelColor"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getShapeNamesWithColor(ByVal ampelColor As Integer) As Collection
        Get
            Dim tmpCollection As New Collection

            Try
                If Not IsNothing(_aCList) Then
                    Dim uidsWithColor As SortedList(Of Integer, Boolean) = _aCList.Item(ampelColor)

                    If Not IsNothing(uidsWithColor) Then
                        ' jetzt sind in der uidList alle ShapeUIDs aufgeführt - die müssen jetzt durch ihre ShapeNames ersetzt werden 
                        For Each kvp As KeyValuePair(Of Integer, Boolean) In uidsWithColor

                            Dim shpName As String = Me.getShapeNameOfUid(kvp.Key)

                            If shpName.Trim.Length > 0 Then
                                If Not tmpCollection.Contains(shpName) Then
                                    tmpCollection.Add(shpName, shpName)
                                End If
                            End If

                        Next
                    End If

                End If
            Catch ex As Exception

            End Try


            getShapeNamesWithColor = tmpCollection

        End Get
    End Property

    ''' <summary>
    ''' liefert eine nicht-sortierte Collection an Namen; das sind alle auftretenden cNames von Phasen und Meilensteinen  
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getElementNamen() As List(Of String)
        Get
            Dim tmpCollection As New List(Of String)

            For Each kvp As KeyValuePair(Of String, SortedList(Of Integer, Boolean)) In _cNList
                If Not tmpCollection.Contains(kvp.Key) Then
                    tmpCollection.Add(kvp.Key)
                End If
            Next

            getElementNamen = tmpCollection
        End Get
    End Property



    ''' <summary>
    ''' bekommt als Input eine Menge von selektierten Namen , classified, Short, Original, etc. 
    ''' gibt als Output die korrespondierenden Shape-Namen
    ''' Achtung: Anzahl Input Elemente muss nicht Anzahl Output Elemente sein;  
    ''' </summary>
    ''' <param name="nameArray"></param>
    ''' <param name="type"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getShapesNames(ByVal nameArray() As String, _
                                                ByVal type As Integer, colorCode As Integer) As Collection
        Get
            Dim tmpCollection As New Collection
            Dim tmpCC As Integer = colorCode
            Dim NList As SortedList(Of String, SortedList(Of Integer, Boolean))
            Dim alleUIDs As New SortedList(Of Integer, Boolean)
            Dim anzahlNames As Integer = nameArray.Length

            Dim alleUIDsWithCertainColor As New SortedList(Of Integer, Boolean)
            Dim resultingUIDs As New SortedList(Of Integer, Boolean)

            If tmpCC >= 8 Then
                Dim redUIDs As SortedList(Of Integer, Boolean) = _aCList.Item(3)
                For i As Integer = 1 To redUIDs.Count
                    If Not alleUIDsWithCertainColor.ContainsKey(redUIDs.ElementAt(i - 1).Key) Then
                        alleUIDsWithCertainColor.Add(redUIDs.ElementAt(i - 1).Key, redUIDs.ElementAt(i - 1).Value)
                    End If
                Next
                tmpCC = tmpCC - 8
            End If

            If tmpCC >= 4 Then
                Dim yellowUIDs As SortedList(Of Integer, Boolean) = _aCList.Item(2)
                For i As Integer = 1 To yellowUIDs.Count
                    If Not alleUIDsWithCertainColor.ContainsKey(yellowUIDs.ElementAt(i - 1).Key) Then
                        alleUIDsWithCertainColor.Add(yellowUIDs.ElementAt(i - 1).Key, yellowUIDs.ElementAt(i - 1).Value)
                    End If
                Next
                tmpCC = tmpCC - 4
            End If

            If tmpCC >= 2 Then
                Dim greenUIDs As SortedList(Of Integer, Boolean) = _aCList.Item(1)
                For i As Integer = 1 To greenUIDs.Count
                    If Not alleUIDsWithCertainColor.ContainsKey(greenUIDs.ElementAt(i - 1).Key) Then
                        alleUIDsWithCertainColor.Add(greenUIDs.ElementAt(i - 1).Key, greenUIDs.ElementAt(i - 1).Value)
                    End If
                Next
                tmpCC = tmpCC - 2
            End If

            If tmpCC >= 1 Then
                Dim noColorUIDs As SortedList(Of Integer, Boolean) = _aCList.Item(0)
                For i As Integer = 1 To noColorUIDs.Count
                    If Not alleUIDsWithCertainColor.ContainsKey(noColorUIDs.ElementAt(i - 1).Key) Then
                        alleUIDsWithCertainColor.Add(noColorUIDs.ElementAt(i - 1).Key, noColorUIDs.ElementAt(i - 1).Value)
                    End If
                Next
                tmpCC = tmpCC - 1
            End If


            Select Case type
                Case pptInfoType.cName
                    NList = _cNList
                Case pptInfoType.oName
                    NList = _oNList
                Case pptInfoType.sName
                    NList = _sNList
                Case pptInfoType.bCrumb
                    NList = _bCList
                Case pptInfoType.lUmfang
                    NList = _LUList
                Case pptInfoType.mvElement
                    NList = _cNList
                Case Else
                    NList = _cNList
            End Select

            For i As Integer = 0 To anzahlNames - 1

                Dim uidList As SortedList(Of Integer, Boolean) = NList.Item(nameArray(i))

                If ((colorCode = 0) Or (colorCode = 15)) Then
                    ' ohne Berücksichtigung von Farben aufnehmen 
                    For Each kvp As KeyValuePair(Of Integer, Boolean) In uidList
                        If Not alleUIDs.ContainsKey(kvp.Key) Then
                            alleUIDs.Add(kvp.Key, kvp.Value)
                        End If
                    Next
                Else
                    ' hat das Element auch eine der gesuchten Farben ? 
                    For Each kvp As KeyValuePair(Of Integer, Boolean) In uidList
                        If alleUIDsWithCertainColor.ContainsKey(kvp.Key) Then
                            If Not alleUIDs.ContainsKey(kvp.Key) Then
                                alleUIDs.Add(kvp.Key, kvp.Value)
                            End If
                        End If
                    Next
                End If


            Next

            ' jetzt muss geprüft werden, ob es sich um mVList handelt - dann muss nochmal ausgedünnt werden ... 
            If type = pptInfoType.mvElement Then
                Dim realUIDs As New SortedList(Of Integer, Boolean)
                For Each kvp As KeyValuePair(Of Integer, Boolean) In alleUIDs
                    If _mVList.ContainsKey(kvp.Key) Then
                        realUIDs.Add(kvp.Key, kvp.Value)
                    End If
                Next
                alleUIDs = realUIDs
            End If

            ' jetzt sind in der uidList alle ShapeUIDs aufgeführt - die müssen jetzt durch ihre ShapeNames ersetzt werden 
            For Each kvp As KeyValuePair(Of Integer, Boolean) In alleUIDs

                Dim shpName As String = Me.getShapeNameOfUid(kvp.Key)

                If shpName.Trim.Length > 0 Then
                    If Not tmpCollection.Contains(shpName) Then
                        tmpCollection.Add(shpName, shpName)
                    End If
                End If

            Next

            getShapesNames = tmpCollection

        End Get
    End Property


    ''' <summary>
    ''' gibt eine Liste zurück an Element-Namen, die den Suchstr enthalten und ausserdem die übergebene Farben-Kennung haben
    ''' leere Liste, wenn es keine Entsprechung gibt  
    ''' </summary>
    ''' <param name="colorCode"></param>
    ''' <param name="suchStr"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getNCollection(ByVal colorCode As Integer, _
                                         ByVal suchStr As String, _
                                         ByVal type As Integer) As Collection
        Get
            Dim NList As SortedList(Of String, SortedList(Of Integer, Boolean))

            Select Case type
                Case pptInfoType.cName
                    NList = _cNList
                Case pptInfoType.oName
                    NList = _oNList
                Case pptInfoType.sName
                    NList = _sNList
                Case pptInfoType.bCrumb
                    NList = _bCList
                Case pptInfoType.lUmfang
                    NList = _LUList
                Case pptInfoType.mvElement
                    NList = _cNList
                Case Else
                    NList = _cNList
            End Select

            Dim tmpCollection As New Collection
            Dim alleUIDsMitgesuchterFarbe As SortedList(Of Integer, Boolean)

            Dim txtRestriction As Boolean = False
            Dim colRestriction As Boolean = False

            ' gibt es eine Text Restriction, also muss der Name irgendwas enthalten ...
            If IsNothing(suchStr) Then
            ElseIf suchStr.Trim.Length = 0 Then
            Else
                txtRestriction = True
            End If

            ' gibt es eine Color Restriction, also sollen nur bestimmte Farben angezeigt werden 
            If colorCode < 1 Or colorCode >= 15 Then
                colRestriction = False
            Else
                colRestriction = True
            End If

            Dim uidList As SortedList(Of Integer, Boolean)

            ' erst wird die Liste an uids ermittelt, die den entsprechenden Farb-Code aufweisen
            ' dann wird untersucht, welche dieser uids ggf noch dem Suchstring entsprechen ... 

            alleUIDsMitgesuchterFarbe = New SortedList(Of Integer, Boolean)


            If colRestriction Then

                ' jetzt muss eine Schleife gemacht werden
                Dim singleFlag As Integer

                Do While colorCode > 0
                    If colorCode >= 8 Then
                        ' red Flag 
                        singleFlag = 3
                        If _aCList.ContainsKey(singleFlag) Then
                            For Each kvp As KeyValuePair(Of Integer, Boolean) In _aCList.Item(singleFlag)
                                alleUIDsMitgesuchterFarbe.Add(kvp.Key, kvp.Value)
                            Next
                        End If
                        colorCode = colorCode - 8

                    ElseIf colorCode >= 4 Then
                        ' yellow flag 
                        singleFlag = 2
                        If _aCList.ContainsKey(singleFlag) Then
                            For Each kvp As KeyValuePair(Of Integer, Boolean) In _aCList.Item(singleFlag)
                                alleUIDsMitgesuchterFarbe.Add(kvp.Key, kvp.Value)
                            Next
                        End If
                        colorCode = colorCode - 4

                    ElseIf colorCode >= 2 Then
                        ' green flag
                        singleFlag = 1
                        If _aCList.ContainsKey(singleFlag) Then
                            For Each kvp As KeyValuePair(Of Integer, Boolean) In _aCList.Item(singleFlag)
                                alleUIDsMitgesuchterFarbe.Add(kvp.Key, kvp.Value)
                            Next
                        End If
                        colorCode = colorCode - 2

                    ElseIf colorCode >= 1 Then
                        ' nicht bewertet 
                        singleFlag = 0
                        If _aCList.ContainsKey(singleFlag) Then
                            For Each kvp As KeyValuePair(Of Integer, Boolean) In _aCList.Item(singleFlag)
                                alleUIDsMitgesuchterFarbe.Add(kvp.Key, kvp.Value)
                            Next
                        End If
                        colorCode = colorCode - 1
                    End If
                Loop


                If alleUIDsMitgesuchterFarbe.Count > 0 Then
                    ' es gibt Shapes - jetzt prüfen, ob es TextRestriktion gibt 
                    If txtRestriction Then
                        ' ermittle die UIDS, die den gesuchten Text enthalten , prüfe gleichzeitig, 
                        ' ob sie bereits in alleUIDSMitgesuchterFarbe sind ... 
                        ' trage die in ErgebnisListe ein 

                        ' Nlsit enthält die Namen, Original-NAmen, etc; jeweils mit einer Liste an UIDS, welche Elemente alle diesen 
                        ' einen Namen enthalten ; ggf kann aj z.B Montage mehrfach vorkommen - und die eine Montage UID hat die gesuchte Farbe, die andere nicht ... 
                        For Each listElem As KeyValuePair(Of String, SortedList(Of Integer, Boolean)) In NList

                            If listElem.Key.Contains(suchStr) Then
                                uidList = listElem.Value
                                For Each chkUID As KeyValuePair(Of Integer, Boolean) In uidList
                                    If alleUIDsMitgesuchterFarbe.ContainsKey(chkUID.Key) Then
                                        ' diese UID ist jetzt eine Ergebnis-UID , die sowhl die richtige Farbe als auch den richtigen Text-String hat 
                                        ' in listElem.key steht der gesuchte String .. 
                                        If Not tmpCollection.Contains(listElem.Key) Then
                                            tmpCollection.Add(listElem.Key, listElem.Key)
                                        End If

                                    End If
                                Next
                            End If

                        Next
                    Else
                        ' ermittle jetzt die Namen, Original-Namen für die Farb-UIDs
                        ' keine Text Restriktion
                        For Each listElem As KeyValuePair(Of String, SortedList(Of Integer, Boolean)) In NList

                            uidList = listElem.Value
                            For Each chkUID As KeyValuePair(Of Integer, Boolean) In uidList
                                If alleUIDsMitgesuchterFarbe.ContainsKey(chkUID.Key) Then
                                    ' diese UID ist jetzt eine Ergebnis-UID , die sowhl die richtige Farbe als auch den richtigen Text-String hat 
                                    ' in listElem.key steht der gesuchte String .. 
                                    If Not tmpCollection.Contains(listElem.Key) Then
                                        tmpCollection.Add(listElem.Key, listElem.Key)
                                    End If
                                End If
                            Next

                        Next
                    End If



                Else
                    ' nichts tun - alleUIDsMitgesuchterFarbe ist leer ...  
                End If

            Else
                ' keine Farb-Einschränkung - also einfach mal die cNList durchgehen 
                For Each listElem As KeyValuePair(Of String, SortedList(Of Integer, Boolean)) In NList

                    If txtRestriction Then
                        If listElem.Key.Contains(suchStr) Then
                            If Not tmpCollection.Contains(listElem.Key) Then
                                tmpCollection.Add(listElem.Key, listElem.Key)
                            End If
                        End If
                    Else
                        If Not tmpCollection.Contains(listElem.Key) Then
                            tmpCollection.Add(listElem.Key, listElem.Key)
                        End If
                    End If


                Next

            End If

            ' jetzt muss im Fall mvList noch geprüft werden, welche Elemente denn verschoben wurden ...
            If type = pptInfoType.mvElement Then
                Dim newCollection As New Collection
                For Each tmpElem As String In tmpCollection
                    Dim tmpUids As SortedList(Of Integer, Boolean) = NList.Item(tmpElem)
                    Dim found As Boolean = False
                    Dim lx As Integer = 0

                    Do While lx <= tmpUids.Count - 1 And Not found
                        If _mVList.ContainsKey(tmpUids.ElementAt(lx).Key) Then
                            If colRestriction Then
                                If alleUIDsMitgesuchterFarbe.ContainsKey(tmpUids.ElementAt(lx).Key) Then
                                    found = True
                                Else
                                    lx = lx + 1
                                End If
                            Else
                                found = True
                            End If

                        Else
                            lx = lx + 1
                        End If
                        If found And Not newCollection.Contains(tmpElem) Then
                            newCollection.Add(tmpElem, tmpElem)
                        End If
                    Loop
                Next
                tmpCollection = newCollection
            End If

            getNCollection = tmpCollection

        End Get
    End Property

    ''' <summary>
    ''' gibt eine Liste zurück an Element-Namen, die den Suchstr enthalten und ausserdem die übergebene Farben-Kennung haben
    ''' leere Liste, wenn es keine Entsprechung gibt  
    ''' </summary>
    ''' <param name="colorCode"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getTNCollection(ByVal colorCode As Integer, _
                                             ByVal nameCollection As Collection) As Collection
        Get
            Dim NList As SortedList(Of String, SortedList(Of Integer, Boolean)) = _cNList


            Dim tmpCollection As New Collection
            Dim alleUIDsMitgesuchterFarbe As SortedList(Of Integer, Boolean)

            Dim colRestriction As Boolean = False


            ' gibt es eine Color Restriction, also sollen nur bestimmte Farben angezeigt werden 
            If colorCode < 1 Or colorCode >= 15 Then
                colRestriction = False
                tmpCollection = nameCollection
            Else
                colRestriction = True

                ' erst wird die Liste an uids ermittelt, die den entsprechenden Farb-Code aufweisen
                ' dann wird untersucht, welche dieser uids ggf noch dem Suchstring entsprechen ... 

                alleUIDsMitgesuchterFarbe = New SortedList(Of Integer, Boolean)

                ' jetzt muss eine Schleife gemacht werden
                Dim singleFlag As Integer

                Do While colorCode > 0
                    If colorCode >= 8 Then
                        ' red Flag 
                        singleFlag = 3
                        If _aCList.ContainsKey(singleFlag) Then
                            For Each kvp As KeyValuePair(Of Integer, Boolean) In _aCList.Item(singleFlag)
                                alleUIDsMitgesuchterFarbe.Add(kvp.Key, kvp.Value)
                            Next
                        End If
                        colorCode = colorCode - 8

                    ElseIf colorCode >= 4 Then
                        ' yellow flag 
                        singleFlag = 2
                        If _aCList.ContainsKey(singleFlag) Then
                            For Each kvp As KeyValuePair(Of Integer, Boolean) In _aCList.Item(singleFlag)
                                alleUIDsMitgesuchterFarbe.Add(kvp.Key, kvp.Value)
                            Next
                        End If
                        colorCode = colorCode - 4

                    ElseIf colorCode >= 2 Then
                        ' green flag
                        singleFlag = 1
                        If _aCList.ContainsKey(singleFlag) Then
                            For Each kvp As KeyValuePair(Of Integer, Boolean) In _aCList.Item(singleFlag)
                                alleUIDsMitgesuchterFarbe.Add(kvp.Key, kvp.Value)
                            Next
                        End If
                        colorCode = colorCode - 2

                    ElseIf colorCode >= 1 Then
                        ' nicht bewertet 
                        singleFlag = 0
                        If _aCList.ContainsKey(singleFlag) Then
                            For Each kvp As KeyValuePair(Of Integer, Boolean) In _aCList.Item(singleFlag)
                                alleUIDsMitgesuchterFarbe.Add(kvp.Key, kvp.Value)
                            Next
                        End If
                        colorCode = colorCode - 1
                    End If
                Loop

                If alleUIDsMitgesuchterFarbe.Count > 0 Then
                    ' es gibt Shapes - jetzt prüfen, ob eines dazu zu den Namen aus nameCollection gehören  

                    ' ermittle jetzt die Namen, Original-Namen für die Farb-UIDs
                    ' keine Text Restriktion
                    For Each kvp As KeyValuePair(Of Integer, Boolean) In alleUIDsMitgesuchterFarbe

                        Dim shpName As String = smartSlideLists.getShapeNameOfUid(kvp.Key)

                        Try
                            Dim tmpShape As PowerPoint.Shape = currentSlide.Shapes(shpName)
                            Dim pruefName As String = tmpShape.Tags.Item("CN")
                            If Not IsNothing(pruefName) Then
                                If pruefName.Length > 0 Then
                                    If nameCollection.Contains(pruefName) Then
                                        If Not tmpCollection.Contains(pruefName) Then
                                            tmpCollection.Add(pruefName, pruefName)
                                        End If
                                    End If
                                End If
                            End If
                        Catch ex As Exception

                        End Try

                    Next


                Else
                    ' nichts tun - alleUIDsMitgesuchterFarbe ist leer ...  
                End If

            End If


            getTNCollection = tmpCollection

        End Get
    End Property




    Public Sub New()
        _planShapeIDs = New SortedList(Of String, Integer)
        _IDplanShapes = New SortedList(Of Integer, String)
        _cNList = New SortedList(Of String, SortedList(Of Integer, Boolean))
        _oNList = New SortedList(Of String, SortedList(Of Integer, Boolean))
        _sNList = New SortedList(Of String, SortedList(Of Integer, Boolean))
        _bCList = New SortedList(Of String, SortedList(Of Integer, Boolean))
        _aCList = New SortedList(Of Integer, SortedList(Of Integer, Boolean))
        _LUList = New SortedList(Of String, SortedList(Of Integer, Boolean))
        _mVList = New SortedList(Of Integer, Boolean)
        _projectTimeStamps = New SortedList(Of String, clsProjektHistorie)
        _listOfTimeStamps = New SortedList(Of Date, Boolean)
        _creationDate = Date.MinValue
        _slideDBUrl = ""
        _slideDBName = ""
    End Sub

End Class
