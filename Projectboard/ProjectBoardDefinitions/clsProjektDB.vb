''' <summary>
''' Klassen-Definition für ein Projekt-Dokument in MongoDB
''' benötigt Klassen-Definitionen clsPhaseDB, clsRolleDB, clsKostenartDB, clsHierarchyDB, clsHierarchyNodeDB, clsResultDB
''' </summary>
''' <remarks></remarks>
Public Class clsProjektDB

    Public name As String
    Public variantName As String
    Public variantDescription As String
    Public Risiko As Double
    Public StrategicFit As Double

    ' Änderung tk: die CustomFields ergänzt ...
    Public customDblFields As SortedList(Of String, Double)
    Public customStringFields As SortedList(Of String, String)
    Public customBoolFields As SortedList(Of String, Boolean)

    Public Erloes As Double
    Public leadPerson As String
    Public tfSpalte As Integer
    Public tfZeile As Integer
    Public startDate As Date
    Public endDate As Date
    Public earliestStart As Integer
    Public earliestStartDate As Date
    Public latestStart As Integer
    Public latestStartDate As Date
    Public status As String
    Public ampelStatus As Integer
    Public ampelErlaeuterung As String
    Public farbe As Integer
    Public Schrift As Integer
    Public Schriftfarbe As Object
    Public VorlagenName As String
    Public Dauer As Integer
    Public AllPhases As List(Of clsPhaseDB)
    Public hierarchy As clsHierarchyDB
    Public Id As String
    Public timestamp As Date
    ' ergänzt am 16.11.13
    Public volumen As Double
    Public complexity As Double
    Public description As String
    Public businessUnit As String

    ' ergänzt am 23.5.18 
    Public projectType As Integer = ptPRPFType.project

    ' ergänzt am 9.6.18 
    Public actualDataUntil As Date = Date.MinValue

    ' ergänzt am 12.6.18 
    Public kundenNummer As String = ""

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

            If Not IsNothing(.Id) Then
                Me.Id = .Id
            End If

            If Not IsNothing(.projectType) Then
                Me.projectType = .projectType
            Else
                Me.projectType = ptPRPFType.project
            End If

            If Not IsNothing(.kundenNummer) Then
                Me.kundenNummer = .kundenNummer
            Else
                Me.kundenNummer = ""
            End If


            ' wenn es einen Varianten-Namen gibt, wird als Datenbank Name 
            ' .name = calcprojektkey(projekt) abgespeichert; das macht das Auslesen später effizienter 

            ' ist es ein Summary Projekt ? 


            Me.name = calcProjektKeyDB(projekt.name, projekt.variantName)

            Me.variantName = .variantName
            Me.variantDescription = .variantDescription

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

            Me.hierarchy.copyFrom(projekt.hierarchy)

            For i = 1 To .CountPhases
                Dim newPhase As New clsPhaseDB
                newPhase.copyFrom(.getPhase(i), .farbe)
                AllPhases.Add(newPhase)
            Next

            ' jetzt werden die CustomFields rausgeschrieben, so fern es welche gibt ... 
            For Each kvp As KeyValuePair(Of Integer, String) In projekt.customStringFields
                Me.customStringFields.Add(CStr(kvp.Key), kvp.Value)
            Next

            For Each kvp As KeyValuePair(Of Integer, Double) In projekt.customDblFields
                Me.customDblFields.Add(CStr(kvp.Key), kvp.Value)
            Next

            For Each kvp As KeyValuePair(Of Integer, Boolean) In projekt.customBoolFields
                Me.customBoolFields.Add(CStr(kvp.Key), kvp.Value)
            Next


        End With

    End Sub

    Public Sub copyto(ByRef projekt As clsProjekt)
        Dim i As Integer
        Dim tmpstr(5) As String


        With projekt
            .timeStamp = Me.timestamp.ToLocalTime
            .Id = Me.Id

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

            ' tk 24.5.18 , wenn Nothing wird das in der Setting Property abgefangen 
            .projectType = Me.projectType

            If IsNothing(Me.kundenNummer) Then
                .kundenNummer = ""
            Else
                .kundenNummer = Me.kundenNummer
            End If

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
                For Each kvp As KeyValuePair(Of String, String) In Me.customStringFields
                    projekt.customStringFields.Add(CInt(kvp.Key), kvp.Value)
                Next
            End If
            If Not IsNothing(Me.customDblFields) Then
                For Each kvp As KeyValuePair(Of String, Double) In Me.customDblFields
                    projekt.customDblFields.Add(CInt(kvp.Key), kvp.Value)
                Next
            End If
            If Not IsNothing(Me.customBoolFields) Then
                For Each kvp As KeyValuePair(Of String, Boolean) In Me.customBoolFields
                    projekt.customBoolFields.Add(CInt(kvp.Key), kvp.Value)
                Next
            End If


        End With

    End Sub


    Public Sub New()

        AllPhases = New List(Of clsPhaseDB)
        hierarchy = New clsHierarchyDB

        customDblFields = New SortedList(Of String, Double)
        customStringFields = New SortedList(Of String, String)
        customBoolFields = New SortedList(Of String, Boolean)

    End Sub

End Class
