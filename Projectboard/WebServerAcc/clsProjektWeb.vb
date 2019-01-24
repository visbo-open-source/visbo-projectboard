
Imports ProjectBoardDefinitions



''' <summary>
''' ''' Vorsicht !!! 
''' bei allen Änderungen in clsProjektWeb und in clsPhaseWeb, da für den MongoDB-Zugriff separate Klassen existieren, die aber fast gleich sind.
''' 
''' Klasse, in der alle Definitionen enthalten sind, die die clsProjektDB für den MongoDB auch enthält, nur passend für ReST-Server-Zugriff
''' </summary>
Public Class clsProjektWeb

    Inherits clsProjektWebShort

    '' ur:2018.07.05: folgende doppelt auskommentiert Definitionen sind in clsProjektWebShort enthalten
    ''Public _id As Object
    ''Public name As String

    '' Änderung ur: vpid wird für VisualBoard als Web-Anwendung benötigt. 
    ''              Im vc sind VisboProjekte enthalten, die über vpid eindeutig vorhandene Projekte referenziert sind.
    ''Public vpid As Object
    ''Public timestamp As Date
    ''Public Erloes As Double
    ''Public startDate As Date
    ''Public endDate As Date
    ''Public status As String

    ''Public variantName As String
    ''Public ampelStatus As Integer


    ' Änderung ur: vor WebServer war dies die ID in der MongoDB (Projektname#Variantename#timestamp)
    Public origId As Object

    Public variantDescription As String
    Public Risiko As Double
    Public StrategicFit As Double

    ' Änderung tk: die CustomFields ergänzt ...
    Public customDblFields As List(Of clsStringDouble)
    Public customStringFields As List(Of clsStringString)
    Public customBoolFields As List(Of clsStringBoolean)


    Public leadPerson As String
    Public tfSpalte As Integer
    Public tfZeile As Integer

    Public earliestStart As Integer
    Public earliestStartDate As Date
    Public latestStart As Integer
    Public latestStartDate As Date

    Public ampelErlaeuterung As String
    Public farbe As Integer
    Public Schrift As Integer
    Public Schriftfarbe As Object
    Public VorlagenName As String
    Public Dauer As Integer
    Public AllPhases As List(Of clsPhaseWeb)
    Public hierarchy As clsHierarchyWeb


    ' ergänzt am 16.11.13
    Public volumen As Double
    Public complexity As Double
    Public description As String
    Public businessUnit As String

    ' ergänzt am 9.6.18 
    Public actualDataUntil As Date

    ' ur: 04.12.2018, da benötigt beim StoreProjecttoDB um sicherzustellen, 
    '                 dass nicht eine inzwischengespeicherte Projektversion einfach überschrieben wird.
    ' verschoben von clsProjektWebLong
    Public updatedAt As String


    ''' <summary>
    ''' kopiert den Inhalt des Projektes (clsProjekt) in clsProjektWeb
    ''' </summary>
    ''' <param name="projekt"></param>
    Public Sub copyfrom(ByVal projekt As clsProjekt)
        Dim i As Integer


        'Me.timestamp = Date.Now
        'Me.Id = 0

        With projekt
            ' damit alle Projekte die gleiche Timestamp für das Datenbank Speichern haben wird das in der 
            ' aufrufenden Sequenz erledigt Me.timestamp = Date.UtcNow
            If Not IsNothing(.timeStamp) Then
                Me.timestamp = .timeStamp.ToUniversalTime
            Else
                Me.timestamp = Date.UtcNow
            End If
            ' ur: 28.05.2018: mit Server wurde umgestellt: id wird von Mongo vergeben
            'If Not IsNothing(.Id) Then
            '    Me.Id = .Id
            'End If

            ' wenn es einen Varianten-Namen gibt, wird als Datenbank Name 
            ' .name = calcprojektkey(projekt) abgespeichert; das macht das Auslesen später effizienter 

            Me.name = .name
            ' ur: 28.05.2018: für RestServer ist Projektname immer ohne Variante
            ' Me.name = calcProjektKeyDB(projekt.name, projekt.variantName)

            Me.variantName = .variantName
            Me.variantDescription = .variantDescription
            ' 6.11.2018: ur: wieder herausgenommen, nun in clsVP
            ''If Not IsNothing(.kundenNummer) Then
            ''    Me.kundennummer = .kundenNummer
            ''Else
            ''    Me.kundennummer = ""
            ''End If

            ' 6.11.2018: ur: hinzugefügt, das in clsProjekt am 7.10.2018 eingeführt
            Me.actualDataUntil = .actualDataUntil.ToUniversalTime


            Me.Risiko = .Risiko
            Me.StrategicFit = .StrategicFit
            Me.Erloes = .Erloes
            Me.leadPerson = .leadPerson
            Me.tfSpalte = .tfspalte
            Me.tfZeile = .tfZeile
            Me.startDate = .startDate.ToUniversalTime
            Me.endDate = .endeDate.ToUniversalTime
            Me.earliestStartDate = .earliestStartDate.ToUniversalTime
            Me.latestStartDate = .latestStartDate.ToUniversalTime
            Me.earliestStart = .earliestStart
            Me.latestStart = .latestStart
            Me.status = .Status
            Me.ampelStatus = .ampelStatus
            Me.ampelErlaeuterung = .ampelErlaeuterung
            Me.farbe = .farbe
            Me.Schrift = .Schrift
            Me.Schriftfarbe = .Schriftfarbe
            Me.VorlagenName = .VorlagenName
            Me.Dauer = .anzahlRasterElemente
            ' ergänzt am 16.11.13
            Me.volumen = .volume
            Me.complexity = .complexity
            Me.description = .description
            Me.businessUnit = .businessUnit

            'ergänzt an 04.12.2018 wird nur zu interne Projektstruktur durchgereicht
            '                      und wieder zurück
            Me.updatedAt = .updatedAt

            Me.hierarchy.copyFrom(projekt.hierarchy)

            For i = 1 To .CountPhases
                Dim newPhase As New clsPhaseWeb
                newPhase.copyFrom(.getPhase(i), .farbe)
                AllPhases.Add(newPhase)
            Next

            ' jetzt werden die CustomFields rausgeschrieben, so fern es welche gibt ... 
            For Each kvp As KeyValuePair(Of Integer, String) In projekt.customStringFields

                If IsNothing(kvp.Value) Or kvp.Value = "" Then
                    Dim hvar As New clsStringString(CStr(kvp.Key), CStr(" "))
                    Me.customStringFields.Add(hvar)
                Else
                    Dim hvar As New clsStringString(CStr(kvp.Key), CStr(kvp.Value))
                    Me.customStringFields.Add(hvar)
                End If

            Next

            For Each kvp As KeyValuePair(Of Integer, Double) In projekt.customDblFields
                Dim hvar As New clsStringDouble(CStr(kvp.Key), CDbl(kvp.Value))
                Me.customDblFields.Add(hvar)
            Next

            For Each kvp As KeyValuePair(Of Integer, Boolean) In projekt.customBoolFields
                Dim hvar As New clsStringBoolean(CStr(kvp.Key), CBool(kvp.Value))
                Me.customBoolFields.Add(hvar)
            Next


        End With

    End Sub

    ''' <summary>
    ''' kopiert den Inhalt eines Projektes (clsProjektWeb) und Teile von clsVP in clsProjekt
    ''' </summary>
    ''' <param name="projekt"></param>
    Public Sub copyto(ByRef projekt As clsProjekt, ByVal vp As clsVP)
        Dim i As Integer
        Dim tmpstr(5) As String


        With projekt
            .timeStamp = Me.timestamp.ToLocalTime

            ' jetzt muss der Datenbank Name aufgesplittet werden in name und variant-Name
            If Me.variantName <> "" And Me.variantName.Trim.Length > 0 Then
                tmpstr = Me.name.Split(New Char() {CChar("#")}, 3)
                If tmpstr.Length > 1 Then
                    If tmpstr(1) = Me.variantName Then
                        .name = tmpstr(0)
                    Else
                        .name = Me.name
                    End If
                Else
                    .name = Me.name
                End If
            Else
                .name = Me.name
            End If

            .variantName = Me.variantName

            ' ergänzt am 17.10.18
            ' 6.11.2018: ur: wieder herausgenommen: ist nun in clsVP
            'If IsNothing(Me.kundennummer) Then
            '    .kundenNummer = ""
            'Else
            '    .kundenNummer = Me.kundennummer
            'End If


            If IsNothing(Me.variantDescription) Then
                .variantDescription = ""
            Else
                .variantDescription = Me.variantDescription
            End If

            .Risiko = Me.Risiko
            .StrategicFit = Me.StrategicFit
            .Erloes = Me.Erloes
            .leadPerson = Me.leadPerson
            ' es gibt kein Attribut tfspalte mehr - es ist ein Readonly Attribut, wo _Start ausgelesen wird 
            '.tfSpalte = Me.tfSpalte
            ' tfzeile wird jetzt ausschließlich durch die Konstellation bestimmt; 
            ' es darf hier nicht mehr gesetzt werden, weil tfzeile die currentConstellation updated ...
            '.tfZeile = Me.tfZeile
            .startDate = Me.startDate.ToLocalTime
            .earliestStartDate = Me.earliestStartDate.ToLocalTime
            .latestStartDate = Me.latestStartDate.ToLocalTime
            .earliestStart = Me.earliestStart
            .latestStart = Me.latestStart
            .Status = Me.status


            .farbe = Me.farbe
            .Schrift = Me.Schrift

            .volume = Me.volumen
            .complexity = Me.complexity
            .description = Me.description
            .businessUnit = Me.businessUnit



            ' Änderung notwendig, weil mal in der Datenbank Schrift mit -10 stand
            If .Schrift < 0 Then
                .Schrift = -1 * .Schrift
            End If
            .Schriftfarbe = Me.Schriftfarbe
            .VorlagenName = Me.VorlagenName

            ' Änderung 18.5.2014: jetzt prüfen, ob diese Vorlage existiert: 
            ' wenn ja, dann übernehmen Farbe, Schrift und Schriftfarbe
            Try
                If Projektvorlagen.Contains(.VorlagenName) Then
                    Dim pvorlage As clsProjektvorlage = Projektvorlagen.getProject(.VorlagenName)
                    .Schrift = pvorlage.Schrift
                    .Schriftfarbe = pvorlage.Schriftfarbe
                    .farbe = pvorlage.farbe
                End If
            Catch ex As Exception
                Call MsgBox(ex.Message & ": im Catch")
            End Try

            Me.hierarchy.copyTo(projekt.hierarchy)

            '.Dauer = Me.Dauer

            For i = 1 To Me.AllPhases.Count
                Dim newPhase As New clsPhase(projekt)
                AllPhases.Item(i - 1).copyto(newPhase, i)
                .AddPhase(newPhase)
            Next

            ' jetzt werden Ampel Status und Beschreibung gesetzt 
            ' da das jetzt in der Phase(1) abgespeichert ist, darf das erst gemacht werden, wenn die Phasen alle kopiert sind ... 
            .ampelStatus = Me.ampelStatus
            .ampelErlaeuterung = Me.ampelErlaeuterung

            ' jetzt werden die CustomFields rausgeschrieben, so fern es welche gibt ... 

            If Not IsNothing(Me.customStringFields) Then
                For Each hvar As clsStringString In Me.customStringFields
                    projekt.customStringFields.Add(CInt(hvar.strkey), hvar.strvalue)
                Next
            End If
            If Not IsNothing(Me.customDblFields) Then
                For Each hvar As clsStringDouble In Me.customDblFields
                    projekt.customDblFields.Add(CInt(hvar.str), hvar.dbl)
                Next
            End If
            If Not IsNothing(Me.customBoolFields) Then
                For Each hvar As clsStringBoolean In Me.customBoolFields
                    projekt.customBoolFields.Add(CInt(hvar.str), hvar.bool)
                Next
            End If


            ' ergänzt am 17.10.18: 
            ' ur: 04.01.2019: muss am Ende stehen, da beim Setzen von actualDataUntil die DauerinDays und damit die Phasen
            ' des aktuellen Projektes und nicht der Vorlage benötigt werden.
            If awinSettings.autoSetActualDataDate Then

                If Me.timestamp.AddMonths(-1) > Me.startDate Then
                    .actualDataUntil = Me.timestamp.AddMonths(-1)
                End If

            Else
                If IsNothing(Me.actualDataUntil) Then
                    .actualDataUntil = Date.MinValue
                Else
                    .actualDataUntil = Me.actualDataUntil.ToLocalTime
                End If
            End If

            'ur:24.01.2019: Infos aus clsVP in clsProjekt benötigt

            If Not IsNothing(vp) Then
                .projectType = vp.vpType
                .kundenNummer = vp.kundennummer
            End If

            ' ur:04.12.2018: ergänzt
            .updatedAt = Me.updatedAt

        End With

    End Sub


    Public Sub New()

        AllPhases = New List(Of clsPhaseWeb)
        hierarchy = New clsHierarchyWeb

        customDblFields = New List(Of clsStringDouble)
        customStringFields = New List(Of clsStringString)
        customBoolFields = New List(Of clsStringBoolean)

    End Sub
    '''''<Serializable()>
    '''''Public Class clsProjektWeb
    '''''    Private _name As String
    '''''    Private _vpid As Object
    '''''    Private _origId As String
    '''''    Private _variantName As String
    '''''    Private _variantDescription As String
    '''''    Private _Risiko As Double
    '''''    Private _StrategicFit As Double
    '''''    Private _customDblFields As SortedList(Of String, Double)
    '''''    Private _customStringFields As SortedList(Of String, String)
    '''''    Private _customBoolFields As SortedList(Of String, Boolean)
    '''''    Private _Erloes As Double
    '''''    Private _leadPerson As String
    '''''    Private _tfSpalte As Integer
    '''''    Private _tfZeile As Integer
    '''''    Private _startDate As String
    '''''    Private _endDate As String
    '''''    Private _earliestStart As Integer
    '''''    Private _earliestStartDate As String
    '''''    Private _latestStart As Integer
    '''''    Private _latestStartDate As String
    '''''    Private _status As String
    '''''    Private _ampelStatus As Integer
    '''''    Private _ampelErlaeuterung As String
    '''''    Private _farbe As Integer
    '''''    Private _Schrift As Integer
    '''''    Private _Schriftfarbe As Object
    '''''    Private _VorlagenName As String
    '''''    Private _Dauer As Integer
    '''''    Private _AllPhases As List(Of clsPhaseDB)
    '''''    Private _hierarchy As clsHierarchyDB
    '''''    Private _Id As Object
    '''''    Private _timestamp As String
    '''''    Private _volumen As Double
    '''''    Private _complexity As Double
    '''''    Private _description As String
    '''''    Private _businessUnit As String

    '''''    Public Property name As String
    '''''        Get
    '''''            name = _name
    '''''        End Get
    '''''        Set(value As String)
    '''''            _name = value
    '''''        End Set
    '''''    End Property
    '''''    ' Änderung ur: vpid wird für VisualBoard als Web-Anwendung benötigt. 
    '''''    '              Im vc sind VisboProjekte enthalten, die über vpid eindeutig vorhandene Projekte referenziert sind.
    '''''    Public Property vpid As Object
    '''''        Get
    '''''            vpid = _vpid
    '''''        End Get
    '''''        Set(value As Object)
    '''''            _vpid = value
    '''''        End Set
    '''''    End Property
    '''''    ' hier ist die ursprüngliche ID in der Form: projName#varName#timestamp
    '''''    ' enthalten
    '''''    Public Property origId As String
    '''''        Get
    '''''            origId = _origId
    '''''        End Get
    '''''        Set(value As String)
    '''''            _origId = value
    '''''        End Set
    '''''    End Property
    '''''    Public Property variantName As String
    '''''        Get
    '''''            variantName = _variantName
    '''''        End Get
    '''''        Set(value As String)
    '''''            _variantName = value
    '''''        End Set
    '''''    End Property
    '''''    Public Property variantDescription As String
    '''''        Get
    '''''            variantDescription = _variantDescription
    '''''        End Get
    '''''        Set(value As String)
    '''''            _variantDescription = value
    '''''        End Set
    '''''    End Property
    '''''    Public Property Risiko As Double
    '''''        Get
    '''''            Risiko = _Risiko
    '''''        End Get
    '''''        Set(value As Double)
    '''''            _Risiko = value
    '''''        End Set
    '''''    End Property
    '''''    Public Property StrategicFit As Double
    '''''        Get
    '''''            StrategicFit = _StrategicFit
    '''''        End Get
    '''''        Set(value As Double)
    '''''            _StrategicFit = value
    '''''        End Set
    '''''    End Property

    '''''    ' Änderung tk: die CustomFields ergänzt ...
    '''''    'Public customDblFields As Object
    '''''    Public Property customDblFields As SortedList(Of String, Double)
    '''''        Get
    '''''            customDblFields = _customDblFields
    '''''        End Get
    '''''        Set(value As SortedList(Of String, Double))
    '''''            If Not IsNothing(value) Then
    '''''                _customDblFields = value
    '''''            End If
    '''''        End Set
    '''''    End Property

    '''''    'Public customStringFields As Object
    '''''    Public Property customStringFields As SortedList(Of String, String)
    '''''        Get
    '''''            customStringFields = _customStringFields
    '''''        End Get
    '''''        Set(value As SortedList(Of String, String))
    '''''            If Not IsNothing(value) Then
    '''''                _customStringFields = value
    '''''            End If
    '''''        End Set
    '''''    End Property

    '''''    'Public customBoolFields As Object
    '''''    Public Property customBoolFields As SortedList(Of String, Boolean)
    '''''        Get
    '''''            customBoolFields = _customBoolFields
    '''''        End Get
    '''''        Set(value As SortedList(Of String, Boolean))
    '''''            If Not IsNothing(value) Then
    '''''                _customBoolFields = value
    '''''            End If
    '''''        End Set
    '''''    End Property

    '''''    Public Property Erloes As Double
    '''''        Get
    '''''            Erloes = _Erloes
    '''''        End Get
    '''''        Set(value As Double)
    '''''            _Erloes = value
    '''''        End Set
    '''''    End Property
    '''''    Public Property leadPerson As String
    '''''        Get
    '''''            leadPerson = _leadPerson
    '''''        End Get
    '''''        Set(value As String)
    '''''            _leadPerson = value
    '''''        End Set
    '''''    End Property
    '''''    Public Property tfSpalte As Integer
    '''''        Get
    '''''            tfSpalte = _tfSpalte
    '''''        End Get
    '''''        Set(value As Integer)
    '''''            _tfSpalte = value
    '''''        End Set
    '''''    End Property
    '''''    Public Property tfZeile As Integer
    '''''        Get
    '''''            tfZeile = _tfZeile
    '''''        End Get
    '''''        Set(value As Integer)
    '''''            _tfZeile = value
    '''''        End Set
    '''''    End Property
    '''''    'Public startDate As date
    '''''    Public Property startDate As String
    '''''        Get
    '''''            startDate = _startDate
    '''''        End Get
    '''''        Set(value As String)
    '''''            _startDate = value
    '''''        End Set
    '''''    End Property
    '''''    'Public endDate As Date
    '''''    Public Property endDate As String
    '''''        Get
    '''''            endDate = _endDate
    '''''        End Get
    '''''        Set(value As String)
    '''''            _endDate = value
    '''''        End Set
    '''''    End Property
    '''''    Public Property earliestStart As Integer
    '''''        Get
    '''''            earliestStart = _earliestStart
    '''''        End Get
    '''''        Set(value As Integer)
    '''''            _earliestStart = value
    '''''        End Set
    '''''    End Property
    '''''    'Public earliestStartDate As Date
    '''''    Public Property earliestStartDate As String
    '''''        Get
    '''''            earliestStartDate = _earliestStartDate
    '''''        End Get
    '''''        Set(value As String)
    '''''            _earliestStartDate = value
    '''''        End Set
    '''''    End Property
    '''''    Public Property latestStart As Integer
    '''''        Get
    '''''            latestStart = _latestStart
    '''''        End Get
    '''''        Set(value As Integer)
    '''''            _latestStart = value
    '''''        End Set
    '''''    End Property
    '''''    'Public latestStartDate As Date
    '''''    Public Property latestStartDate As String
    '''''        Get
    '''''            latestStartDate = _latestStartDate
    '''''        End Get
    '''''        Set(value As String)
    '''''            _latestStartDate = value
    '''''        End Set
    '''''    End Property
    '''''    Public Property status As String
    '''''        Get
    '''''            status = _status
    '''''        End Get
    '''''        Set(value As String)
    '''''            _status = value
    '''''        End Set
    '''''    End Property
    '''''    Public Property ampelStatus As Integer
    '''''        Get
    '''''            ampelStatus = _ampelStatus
    '''''        End Get
    '''''        Set(value As Integer)
    '''''            _ampelStatus = value
    '''''        End Set
    '''''    End Property
    '''''    Public Property ampelErlaeuterung As String
    '''''        Get
    '''''            ampelErlaeuterung = _ampelErlaeuterung
    '''''        End Get
    '''''        Set(value As String)
    '''''            _ampelErlaeuterung = value
    '''''        End Set
    '''''    End Property
    '''''    Public Property farbe As Integer
    '''''        Get
    '''''            farbe = _farbe
    '''''        End Get
    '''''        Set(value As Integer)
    '''''            _farbe = value
    '''''        End Set
    '''''    End Property
    '''''    Public Property Schrift As Integer
    '''''        Get
    '''''            Schrift = _Schrift
    '''''        End Get
    '''''        Set(value As Integer)
    '''''            _Schrift = value
    '''''        End Set
    '''''    End Property
    '''''    Public Property Schriftfarbe As Object
    '''''        Get
    '''''            Schriftfarbe = _Schriftfarbe
    '''''        End Get
    '''''        Set(value As Object)
    '''''            _Schriftfarbe = value
    '''''        End Set
    '''''    End Property
    '''''    Public Property VorlagenName As String
    '''''        Get
    '''''            VorlagenName = _VorlagenName
    '''''        End Get
    '''''        Set(value As String)
    '''''            _VorlagenName = value
    '''''        End Set
    '''''    End Property
    '''''    Public Property Dauer As Integer
    '''''        Get
    '''''            Dauer = _Dauer
    '''''        End Get
    '''''        Set(value As Integer)
    '''''            _Dauer = value
    '''''        End Set
    '''''    End Property
    '''''    Public Property AllPhases As List(Of clsPhaseDB)
    '''''        Get
    '''''            AllPhases = _AllPhases
    '''''        End Get
    '''''        Set(value As List(Of clsPhaseDB))
    '''''            If Not IsNothing(value) Then
    '''''                _AllPhases = value
    '''''            End If
    '''''        End Set
    '''''    End Property
    '''''    'Public hierarchy As clsHierarchyDB
    '''''    Public Property hierarchy As clsHierarchyDB
    '''''        Get
    '''''            hierarchy = _hierarchy
    '''''        End Get
    '''''        Set(value As clsHierarchyDB)
    '''''            _hierarchy = value
    '''''        End Set
    '''''    End Property
    '''''    'wird im ServerUmfeld als normale DB-Id verwendet nicht: ProjName#varName#Timestamp
    '''''    Public Property Id As Object
    '''''        Get
    '''''            Id = _Id
    '''''        End Get
    '''''        Set(value As Object)
    '''''            _Id = value
    '''''        End Set
    '''''    End Property
    '''''    'Public timestamp As Date
    '''''    Public Property timestamp As String
    '''''        Get
    '''''            timestamp = _timestamp
    '''''        End Get
    '''''        Set(value As String)
    '''''            _timestamp = value
    '''''        End Set
    '''''    End Property
    '''''    ' ergänzt am 16.11.13
    '''''    Public Property volumen As Double
    '''''        Get
    '''''            volumen = _volumen
    '''''        End Get
    '''''        Set(value As Double)
    '''''            _volumen = value
    '''''        End Set
    '''''    End Property
    '''''    Public Property complexity As Double
    '''''        Get
    '''''            complexity = _complexity
    '''''        End Get
    '''''        Set(value As Double)
    '''''            _complexity = value
    '''''        End Set
    '''''    End Property
    '''''    Public Property description As String
    '''''        Get
    '''''            description = _description
    '''''        End Get
    '''''        Set(value As String)
    '''''            _description = value
    '''''        End Set
    '''''    End Property
    '''''    Public Property businessUnit As String
    '''''        Get
    '''''            businessUnit = _businessUnit
    '''''        End Get
    '''''        Set(value As String)
    '''''            _businessUnit = value
    '''''        End Set
    '''''    End Property

    '''''    Public Sub copyfrom(ByVal projekt As clsProjekt)
    '''''        Dim i As Integer


    '''''        'Me.timestamp = Date.Now
    '''''        'Me.Id = 0

    '''''        With projekt
    '''''            ' damit alle Projekte die gleiche Timestamp für das Datenbank Speichern haben wird das in der 
    '''''            ' aufrufenden Sequenz erledigt Me.timestamp = Date.UtcNow
    '''''            If Not IsNothing(.timeStamp) Then
    '''''                Me.timestamp = .timeStamp.ToUniversalTime
    '''''            Else
    '''''                Me.timestamp = Date.UtcNow
    '''''            End If

    '''''            If Not IsNothing(.Id) Then
    '''''                Me.Id = .Id
    '''''            End If

    '''''            ' wenn es einen Varianten-Namen gibt, wird als Datenbank Name 
    '''''            ' .name = calcprojektkey(projekt) abgespeichert; das macht das Auslesen später effizienter 

    '''''            Me.name = calcProjektKeyDB(projekt.name, projekt.variantName)

    '''''            Me.variantName = .variantName
    '''''            Me.variantDescription = .variantDescription

    '''''            Me.Risiko = .Risiko
    '''''            Me.StrategicFit = .StrategicFit
    '''''            Me.Erloes = .Erloes
    '''''            Me.leadPerson = .leadPerson
    '''''            Me.tfSpalte = .tfspalte
    '''''            Me.tfZeile = .tfZeile
    '''''            Me.startDate = .startDate.ToUniversalTime
    '''''            Me.endDate = .endeDate.ToUniversalTime
    '''''            Me.earliestStartDate = .earliestStartDate.ToUniversalTime
    '''''            Me.latestStartDate = .latestStartDate.ToUniversalTime
    '''''            Me.earliestStart = .earliestStart
    '''''            Me.latestStart = .latestStart
    '''''            Me.status = .Status
    '''''            Me.ampelStatus = .ampelStatus
    '''''            Me.ampelErlaeuterung = .ampelErlaeuterung
    '''''            Me.farbe = .farbe
    '''''            Me.Schrift = .Schrift
    '''''            Me.Schriftfarbe = .Schriftfarbe
    '''''            Me.VorlagenName = .VorlagenName
    '''''            Me.Dauer = .anzahlRasterElemente
    '''''            ' ergänzt am 16.11.13
    '''''            Me.volumen = .volume
    '''''            Me.complexity = .complexity
    '''''            Me.description = .description
    '''''            'Me.businessUnit = .businessUnit

    '''''            Me.hierarchy.copyFrom(projekt.hierarchy)

    '''''            For i = 1 To .CountPhases
    '''''                Dim newPhase As New clsPhaseDB
    '''''                newPhase.copyFrom(.getPhase(i), .farbe)
    '''''                AllPhases.Add(newPhase)
    '''''            Next

    '''''            ' jetzt werden die CustomFields rausgeschrieben, so fern es welche gibt ... 
    '''''            For Each kvp As KeyValuePair(Of Integer, String) In projekt.customStringFields
    '''''                Me.customStringFields.Add(CStr(kvp.Key), kvp.Value)
    '''''            Next

    '''''            For Each kvp As KeyValuePair(Of Integer, Double) In projekt.customDblFields
    '''''                Me.customDblFields.Add(CStr(kvp.Key), kvp.Value)
    '''''            Next

    '''''            For Each kvp As KeyValuePair(Of Integer, Boolean) In projekt.customBoolFields
    '''''                Me.customBoolFields.Add(CStr(kvp.Key), kvp.Value)
    '''''            Next


    '''''        End With

    '''''    End Sub

    '''''    Public Sub copyto(ByRef projekt As clsProjekt)
    '''''        Dim i As Integer
    '''''        Dim tmpstr(5) As String
    '''''        Dim provider As CultureInfo = CultureInfo.CurrentCulture

    '''''        With projekt
    '''''            'Dim ok As Boolean = DateTime.TryParseExact(Me.timestamp, "yyyy-MM-ddThh:mm:ss.fffZ",
    '''''            '                                               provider, style:=DateTimeStyles.AssumeUniversal, result:= .timeStamp)
    '''''            'projekt.timeStamp = projekt.timeStamp.ToLocalTime
    '''''            .timeStamp = DateTime.ParseExact(Me.timestamp, "yyyy-MM-ddThh:mm:ss.fffZ",
    '''''                                             provider, style:=DateTimeStyles.AssumeUniversal)
    '''''            .Id = Me.Id

    '''''            ' jetzt muss der Datenbank Name aufgesplittet werden in name und variant-Name
    '''''            If Me.variantName <> "" And Me.variantName.Trim.Length > 0 Then
    '''''                tmpstr = Me.name.Split(New Char() {CChar("#")}, 3)
    '''''                If tmpstr.Length > 1 Then
    '''''                    If tmpstr(1) = Me.variantName Then
    '''''                        .name = tmpstr(0)
    '''''                    Else
    '''''                        .name = Me.name
    '''''                    End If
    '''''                Else
    '''''                    .name = Me.name
    '''''                End If
    '''''            Else
    '''''                .name = Me.name
    '''''            End If

    '''''            .variantName = Me.variantName

    '''''            If IsNothing(Me.variantDescription) Then
    '''''                .variantDescription = ""
    '''''            Else
    '''''                .variantDescription = Me.variantDescription
    '''''            End If

    '''''            .Risiko = Me.Risiko
    '''''            .StrategicFit = Me.StrategicFit
    '''''            .Erloes = Me.Erloes
    '''''            .leadPerson = Me.leadPerson
    '''''            ' es gibt kein Attribut tfspalte mehr - es ist ein Readonly Attribut, wo _Start ausgelesen wird 
    '''''            '.tfSpalte = Me.tfSpalte
    '''''            ' tfzeile wird jetzt ausschließlich durch die Konstellation bestimmt; 
    '''''            ' es darf hier nicht mehr gesetzt werden, weil tfzeile die currentConstellation updated ...
    '''''            '.tfZeile = Me.tfZeile
    '''''            .startDate = DateTime.ParseExact(Me.timestamp, "yyyy-MM-ddThh:mm:ss.fffZ",
    '''''                                                           provider, style:=DateTimeStyles.AssumeUniversal)
    '''''            .earliestStartDate = DateTime.ParseExact(Me.timestamp, "yyyy-MM-ddThh:mm:ss.fffZ",
    '''''                                                           provider, style:=DateTimeStyles.AssumeUniversal)
    '''''            .latestStartDate = DateTime.ParseExact(Me.timestamp, "yyyy-MM-ddThh:mm:ss.fffZ",
    '''''                                                           provider, style:=DateTimeStyles.AssumeUniversal)
    '''''            .earliestStart = Me.earliestStart
    '''''            .latestStart = Me.latestStart
    '''''            .Status = Me.status

    '''''            .farbe = Me.farbe
    '''''            .Schrift = Me.Schrift

    '''''            .volume = Me.volumen
    '''''            .complexity = Me.complexity
    '''''            .description = Me.description
    '''''            '.businessUnit = Me.businessUnit

    '''''            ' Änderung notwendig, weil mal in der Datenbank Schrift mit -10 stand
    '''''            If .Schrift < 0 Then
    '''''                .Schrift = -1 * .Schrift
    '''''            End If
    '''''            .Schriftfarbe = Me.Schriftfarbe
    '''''            .VorlagenName = Me.VorlagenName

    '''''            ' Änderung 18.5.2014: jetzt prüfen, ob diese Vorlage existiert: 
    '''''            ' wenn ja, dann übernehmen Farbe, Schrift und Schriftfarbe
    '''''            Try
    '''''                If Projektvorlagen.Contains(.VorlagenName) Then
    '''''                    Dim pvorlage As clsProjektvorlage = Projektvorlagen.getProject(.VorlagenName)
    '''''                    .Schrift = pvorlage.Schrift
    '''''                    .Schriftfarbe = pvorlage.Schriftfarbe
    '''''                    .farbe = pvorlage.farbe
    '''''                End If
    '''''            Catch ex As Exception

    '''''            End Try

    '''''            Me.hierarchy.copyTo(projekt.hierarchy)

    '''''            '.Dauer = Me.Dauer
    '''''            For i = 1 To Me.AllPhases.Count
    '''''                Dim newPhase As New clsPhase(projekt)
    '''''                AllPhases.Item(i - 1).copyto(newPhase, i)
    '''''                .AddPhase(newPhase)
    '''''            Next

    '''''            ' jetzt werden Ampel Status und Beschreibung gesetzt 
    '''''            ' da das jetzt in der Phase(1) abgespeichert ist, darf das erst gemacht werden, wenn die Phasen alle kopiert sind ... 
    '''''            .ampelStatus = Me.ampelStatus
    '''''            .ampelErlaeuterung = Me.ampelErlaeuterung

    '''''            ' jetzt werden die CustomFields rausgeschrieben, so fern es welche gibt ... 

    '''''            If Not IsNothing(Me.customStringFields) Then
    '''''                For Each kvp As KeyValuePair(Of String, String) In Me.customStringFields
    '''''                    projekt.customStringFields.Add(CInt(kvp.Key), kvp.Value)
    '''''                Next
    '''''            End If
    '''''            If Not IsNothing(Me.customDblFields) Then
    '''''                For Each kvp As KeyValuePair(Of String, Double) In Me.customDblFields
    '''''                    projekt.customDblFields.Add(CInt(kvp.Key), kvp.Value)
    '''''                Next
    '''''            End If
    '''''            If Not IsNothing(Me.customBoolFields) Then
    '''''                For Each kvp As KeyValuePair(Of String, Boolean) In Me.customBoolFields
    '''''                    projekt.customBoolFields.Add(CInt(kvp.Key), kvp.Value)
    '''''                Next
    '''''            End If


    '''''        End With

    '''''    End Sub


    '''''    Public Sub New()

    '''''        _AllPhases = New List(Of clsPhaseDB)
    '''''        _hierarchy = New clsHierarchyDB()
    '''''        _Id = ""

    '''''        _name = ""
    '''''        ' Änderung ur: vpid wird für VisualBoard als Web-Anwendung benötigt. 
    '''''        '              Im vc sind VisboProjekte enthalten, die über vpid eindeutig vorhandene Projekte referenziert sind.
    '''''        _vpid = ""
    '''''        ' hier ist die ursprüngliche ID in der Form: projName#varName#timestamp
    '''''        ' enthalten
    '''''        _origId = ""
    '''''        _variantName = ""
    '''''        _variantDescription = ""
    '''''        _Risiko = 0.0
    '''''        _StrategicFit = 0.0
    '''''        ' Änderung tk: die CustomFields ergänzt ...
    '''''        _customDblFields = New SortedList(Of String, Double)
    '''''        'Public customDblFields As Object
    '''''        _customStringFields = New SortedList(Of String, String)
    '''''        'Public customStringFields As Object
    '''''        _customBoolFields = New SortedList(Of String, Boolean)
    '''''        'Public customBoolFields As Object

    '''''        _Erloes = 0.0
    '''''        _leadPerson = "noname"
    '''''        _tfSpalte = 0
    '''''        _tfZeile = 0
    '''''        'Public startDate As date
    '''''        _startDate = Date.MinValue.ToString
    '''''        'Public endDate As Date
    '''''        _endDate = Date.MaxValue.ToString
    '''''        _earliestStart = 0
    '''''        'Public earliestStartDate As Date
    '''''        _earliestStartDate = ""
    '''''        _latestStart = 0
    '''''        'Public latestStartDate As Date
    '''''        _latestStartDate = ""
    '''''        _status = ""
    '''''        _ampelStatus = 0
    '''''        _ampelErlaeuterung = ""
    '''''        _farbe = 0
    '''''        _Schrift = 0
    '''''        _Schriftfarbe = Nothing
    '''''        _VorlagenName = ""
    '''''        _Dauer = 0
    '''''        'Public timestamp As Date
    '''''        _timestamp = Date.Now.ToString
    '''''        ' ergänzt am 16.11.13
    '''''        _volumen = 0.0
    '''''        _complexity = 0.0
    '''''        _description = ""
    '''''        _businessUnit = ""

    '''''    End Sub

    '''''End Class
End Class
