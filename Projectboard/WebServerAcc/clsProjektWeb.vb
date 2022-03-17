
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
    ''Public vpStatus as String

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

    'ergänzt am 14.10.2019 für HTML5-Viewer in UI 
    Public keyMetrics As clsKeyMetrics


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

            ' ur:20210426: sollte nun automatisch vom Server aus den VP-Properties geholt werden
            ' diese Werte werden von der ServerVersion ab Mai 2021 nicht mehr gespeichert, sonder die von der zugehörigen VP
            Me.Risiko = .Risiko
            Me.StrategicFit = .StrategicFit
            Me.businessUnit = .businessUnit

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
            'ur: 210203: Me.status = .Status        ' wird nicht mehr an die DB weitergegeben
            ' ur: 20210915 neues Property übernommen aus VP kann in Projectboard nicht geändert werden
            Me.vpStatus = .vpStatus
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

            ' 20.04.30: ur: es wird am Client keine keyMetrics mehr angelegt
            ' jetzt werden die keyMetrics übertragen, sofern am Client definiert

            If (Not IsNothing(projekt.keyMetrics)) Then
                Me.keyMetrics.costCurrentActual = projekt.keyMetrics.costCurrentActual
                Me.keyMetrics.costCurrentTotal = projekt.keyMetrics.costCurrentTotal
                Me.keyMetrics.costBaseLastActual = projekt.keyMetrics.costBaseLastActual
                Me.keyMetrics.costBaseLastTotal = projekt.keyMetrics.costBaseLastTotal

                Me.keyMetrics.timeCompletionCurrentActual = projekt.keyMetrics.timeCompletionCurrentActual
                Me.keyMetrics.timeCompletionBaseLastActual = projekt.keyMetrics.timeCompletionBaseLastActual
                Me.keyMetrics.timeCompletionCurrentTotal = projekt.keyMetrics.timeCompletionCurrentTotal
                Me.keyMetrics.timeCompletionBaseLastTotal = projekt.keyMetrics.timeCompletionBaseLastTotal
                Me.keyMetrics.endDateCurrent = projekt.keyMetrics.endDateCurrent.ToUniversalTime
                Me.keyMetrics.endDateBaseLast = projekt.keyMetrics.endDateBaseLast.ToUniversalTime

                Me.keyMetrics.deliverableCompletionCurrentActual = projekt.keyMetrics.deliverableCompletionCurrentActual
                Me.keyMetrics.deliverableCompletionCurrentTotal = projekt.keyMetrics.deliverableCompletionCurrentTotal

                Me.keyMetrics.deliverableCompletionBaseLastActual = projekt.keyMetrics.deliverableCompletionBaseLastActual
                Me.keyMetrics.deliverableCompletionBaseLastTotal = projekt.keyMetrics.deliverableCompletionBaseLastTotal

                Me.keyMetrics.timeDelayCurrentActual = projekt.keyMetrics.timeDelayCurrentActual
                Me.keyMetrics.timeDelayCurrentTotal = projekt.keyMetrics.timeDelayCurrentTotal
                Me.keyMetrics.deliverableDelayCurrentActual = projekt.keyMetrics.deliverableDelayCurrentActual
                Me.keyMetrics.deliverableDelayCurrentTotal = projekt.keyMetrics.deliverableDelayCurrentTotal
            Else
                Me.keyMetrics = Nothing
            End If


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

            ' tk 11.5.19 , Me.vpid hat den glkeichen Inhalt wie vp._id
            .vpID = Me.vpid
            ' ur: 07.06.2020: Me._id ist die vpv._id dieser ProjektVersion
            .Id = Me._id

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

            ' ur: 20210426: neue vp-Properties nun aus VP in VPV kopieren(siehe unten)
            .Risiko = Me.Risiko
            .StrategicFit = Me.StrategicFit
            .businessUnit = Me.businessUnit

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
            '.Status = Me.status

            '.farbe = Me.farbe
            .Schrift = Me.Schrift

            .volume = Me.volumen
            .complexity = Me.complexity
            .description = Me.description

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
                    '.farbe = pvorlage.farbe
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
                    ' ur: 22.02.19: im copyfrom wird dieses Leerzeichen eingefügt, da sonst nicht gespeichert wird.
                    ' hier muss das wieder herausgefiltert werden.
                    If hvar.strvalue = " " Then
                        hvar.strvalue = ""
                    End If
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

            ' ergänzt 14.10.2019: keyMetrics für HTML5 in vpv-long gespeichert
            ' jetzt werden die keyMetrics übertragen

            ' nur anlegen, wenn Me.keyMetrics Werte enthält
            If (Not IsNothing(Me.keyMetrics)) Then

                .keyMetrics = New clsKeyMetrics
                .keyMetrics.costCurrentActual = Me.keyMetrics.costCurrentActual
                .keyMetrics.costCurrentTotal = Me.keyMetrics.costCurrentTotal
                .keyMetrics.costBaseLastActual = Me.keyMetrics.costBaseLastActual
                .keyMetrics.costBaseLastTotal = Me.keyMetrics.costBaseLastTotal
                .keyMetrics.timeCompletionCurrentActual = Me.keyMetrics.timeCompletionCurrentActual
                .keyMetrics.timeCompletionBaseLastActual = Me.keyMetrics.timeCompletionBaseLastActual
                .keyMetrics.timeCompletionCurrentTotal = Me.keyMetrics.timeCompletionCurrentTotal
                .keyMetrics.timeCompletionBaseLastTotal = Me.keyMetrics.timeCompletionBaseLastTotal

                If IsNothing(Me.keyMetrics.endDateCurrent) Then
                    .keyMetrics.endDateCurrent = Date.MinValue
                Else
                    .keyMetrics.endDateCurrent = Me.keyMetrics.endDateCurrent.ToLocalTime
                End If
                If IsNothing(Me.keyMetrics.endDateBaseLast) Then
                    .keyMetrics.endDateBaseLast = Date.MinValue
                Else
                    .keyMetrics.endDateBaseLast = Me.keyMetrics.endDateBaseLast.ToLocalTime
                End If

                .keyMetrics.deliverableCompletionCurrentActual = Me.keyMetrics.deliverableCompletionCurrentActual
                .keyMetrics.deliverableCompletionCurrentTotal = Me.keyMetrics.deliverableCompletionCurrentTotal
                .keyMetrics.deliverableCompletionBaseLastActual = Me.keyMetrics.deliverableCompletionBaseLastActual
                .keyMetrics.deliverableCompletionBaseLastTotal = Me.keyMetrics.deliverableCompletionBaseLastTotal

                .keyMetrics.timeDelayCurrentActual = Me.keyMetrics.timeDelayCurrentActual
                .keyMetrics.timeDelayCurrentTotal = Me.keyMetrics.timeDelayCurrentTotal
                .keyMetrics.deliverableDelayCurrentActual = Me.keyMetrics.deliverableDelayCurrentActual
                .keyMetrics.deliverableDelayCurrentTotal = Me.keyMetrics.deliverableDelayCurrentTotal
            Else
                .keyMetrics = Nothing
            End If

            If IsNothing(Me.actualDataUntil) Then
                .actualDataUntil = Date.MinValue
            Else
                .actualDataUntil = Me.actualDataUntil.ToLocalTime
            End If



            ''ur:24.01.2019: Infos aus clsVP in clsProjekt benötigt
            If Not IsNothing(vp) Then
                .projectType = vp.vpType
                .kundenNummer = vp.kundennummer

                If Not IsNothing(vp.managerID) Then
                    'read the userName an put it into leadPerson

                End If

                ' ur: 20210426: neue vp-Properties nun aus VP in VPV kopieren
                If Not IsNothing(vp.customFieldDouble) Then
                    For Each item As clsCustomFieldDbl In vp.customFieldDouble
                        If item.name = vp_strategicFit And item.type = "System" Then
                            .StrategicFit = item.value
                        End If
                        If item.name = vp_risk And item.type = "System" Then
                            .Risiko = item.value
                        End If
                    Next
                End If

                If Not IsNothing(vp.customFieldString) Then
                    For Each item As clsCustomFieldStr In vp.customFieldString
                        If item.name = vp_businessUnit And item.type = "System" Then
                            .businessUnit = item.value
                        End If
                    Next
                End If
                If Not IsNothing(vp.customFieldDate) Then
                    ' ur: 20210616: prepared for Monitoring process quality
                    'For Each item As clsCustomFieldDate In vp.customFieldDate
                    '    If item.name = vp_pmCommit And item.type = "System" Then
                    '        .pmCommit = item.value
                    '    End If
                    'Next
                End If

                ' ur: 20210915 neue Property aus vp übernommen, kann in projectboard nicht geändert werden
                .vpStatus = vp.vpStatus

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

        ' 20.04.30: ur: keyMetrics nicht mehr mit anlegen
        keyMetrics = New clsKeyMetrics
    End Sub

End Class
