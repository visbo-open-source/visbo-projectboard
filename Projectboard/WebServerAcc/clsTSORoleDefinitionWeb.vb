Imports ProjectBoardDefinitions
Public Class clsTSORoleDefinitionWeb

    ' bei subRoleIDs eigentlich integer, string), muss wegen Mongo auf String geändert werden 
    ' tk 29.5.18 in den SubroleID values steht jetzt im String nicht mehr der Name, der ist ohnehin redundant zur UID, sondern der Prozentsatz, wieviel die Rolle zur Kapa der Sammelrolle beiträgt 
    ' wenn ein nicht als double interpretierbarer Wert drinsteht (=alte Speicherungen, dann wird der Wert auf 1.0 gesetzt 
    Public subRoleIDs As List(Of clsSubRoleID)


    Public teamIDs As List(Of clsSubRoleID)


    Public isExternRole As Boolean
        Public isTeam As Boolean


    Public aliases As String()
    Public employeeNr As String
    Public defaultDayCapa As Double
    Public entryDate As Date
    Public exitDate As Date

    Public uid As Integer

    Public name As String
    Public farbe As Long
    Public defaultKapa As Double
    Public tagessatzIntern As Double
    '9.11.20 ur wird nun doppelt geführt und dann später wird tagessatzintern ausgeführt
    Public tagessatz As Double

    Public kapazitaet() As Double

    '08.07.2021  neu hinzugekommen mit VS-943 Re-Structure Orga
    Public isAggregationRole As Boolean
    Public isSummaryRole As Boolean
    'Public isActDataRelevant As Boolean

    '25.02.2022: ur: changes with view on timeStampedOrga (all properties in english)
    Public type As Integer
    Public dailyRate As Double
    Public defCapaMonth As Double
    Public defCapaDay As Double
    Public capaPerMonth() As Double

    ' startOfCal ist wichtig, damit die korrekte Zuordnung der Kapa-Werte zu den Monaten gemacht werden kann 
    Public startOfCal As Date

    ' war vorher : Public Sub copyTo(ByRef roleDef As clsRollenDefinition, ByVal orgaStartOfCalendar As Date)
    Public Sub copyTo(ByRef roleDef As clsRollenDefinition)

        Dim entryColOfDate As Integer = 0
        Dim exitColOfDate As Integer = 0

        If subRoleIDs.Count >= 1 Then
            ' wegen Mongo müssen die Keys in String Format sein ... 

            For Each sr As clsSubRoleID In Me.subRoleIDs
                Dim tmpValue As Double = 1.0
                If IsNumeric(sr.value) Then
                    tmpValue = CDbl(sr.value)
                    If tmpValue >= 0 And tmpValue <= 1.0 Then
                        ' alles ok
                    Else
                        tmpValue = 1.0
                    End If
                Else
                    tmpValue = 1.0
                End If

                Try
                    roleDef.addSubRole(CInt(sr.key), tmpValue)
                Catch ex As Exception
                    Call MsgBox("1119765: not allowed to have both team-Membership and Childs ..")
                End Try

            Next
        End If

        ' tk 23.11.18 dazugekommen 
        If teamIDs.Count >= 1 Then
            ' wegen Mongo müssen die Keys in String Format sein ... 

            For Each sr As clsSubRoleID In Me.teamIDs
                Dim tmpValue As Double = 1.0
                If IsNumeric(sr.value) Then
                    tmpValue = CDbl(sr.value)
                    If tmpValue >= 0 And tmpValue <= 1.0 Then
                        ' alles ok
                    Else
                        tmpValue = 1.0
                    End If
                Else
                    tmpValue = 1.0
                End If
                Try
                    roleDef.addSkill(CInt(sr.key), tmpValue)
                Catch ex As Exception
                    Call MsgBox("1119765: not allowed to to have team-Membership and Childs ..")
                End Try

            Next
        End If

        roleDef.UID = Me.uid
        roleDef.name = Me.name
        ' tk 20.7.21 keine individuelle Farbe mehr für Rollen 
        'roleDef.farbe = Me.farbe

        ' ur: 25022022 renamed this property in the Server
        'roleDef.defaultKapa = Me.defaultKapa
        roleDef.defaultKapa = Me.defCapaMonth

        ' tk 8.1.20
        roleDef.aliases = Me.aliases
        ' ur: 25022022 renamed this property in the Server
        'roleDef.defaultDayCapa = Me.defaultDayCapa
        roleDef.defaultDayCapa = Me.defCapaDay
        roleDef.employeeNr = Me.employeeNr

        roleDef.entryDate = Me.entryDate.ToLocalTime
        If roleDef.entryDate <= StartofCalendar Then
            entryColOfDate = 1
        Else
            entryColOfDate = getColumnOfDate(roleDef.entryDate)
        End If

        roleDef.exitDate = Me.exitDate.ToLocalTime
        If roleDef.exitDate >= DateAndTime.DateSerial(2200, 12, 31) Then
            exitColOfDate = 241
        Else
            exitColOfDate = Math.Min(roleDef.kapazitaet.Length, getColumnOfDate(roleDef.exitDate))
        End If


        ' tk 23.11.18 
        roleDef.isExternRole = Me.isExternRole

        ' ur: 25022022: isTeam replaced by type
        'roleDef.isSkill = Me.isTeam
        If (Me.isSummaryRole) And (Me.type = 2) Then
            roleDef.isSkill = True
        Else
            roleDef.isSkill = False
        End If


        ' 9.11.20 ur for a smart change to tagessatz
        If Not IsNothing(Me.dailyRate) Then
            If Me.dailyRate = 0 Then
                roleDef.tagessatzIntern = Me.tagessatzIntern
            Else
                roleDef.tagessatzIntern = Me.dailyRate
            End If
        Else
            roleDef.tagessatzIntern = Me.tagessatzIntern
        End If

        ' neu ur: 08.07.2021
        roleDef.isAggregationRole = Me.isAggregationRole
        ' ur: 25022022  this property will further be in the customization VCSetting in the Server
        'roleDef.isActDataRelevant = Me.isActDataRelevant
        roleDef.isSummaryRole = Me.isSummaryRole
        Dim kapaArray() As Double = Me.capaPerMonth

        ' jetzt die Übernahme der Kapazitäten 
        ' Rollen, die Kinder haben tragen niemals Kapa , also immer Null 
        ' ebenso Rollen, die nur den Default Wert haben 

        ' hier muss auch nur was gemacht werden, wenn subRoleId.count = 0 
        If subRoleIDs.Count = 0 Then

            Dim nrWebCapaValues As Integer
            If Not IsNothing(Me.capaPerMonth) Then
                nrWebCapaValues = Me.capaPerMonth.Length
            Else
                nrWebCapaValues = 0
            End If

            Dim lenSession As Integer = roleDef.kapazitaet.Length


            ' ' vorbesetzen mit dem Default Wert
            ' ur: 14.03.2022: changed because of exitDate 
            'For i As Integer = 1 To lenSession - 1
            '    roleDef.kapazitaet(i) = roleDef.defaultKapa
            'Next

            ' ' vorbesetzen mit dem Default Wert
            For i As Integer = entryColOfDate To exitColOfDate - 1
                roleDef.kapazitaet(i) = roleDef.defaultKapa
            Next

            ' das muss jetzt nur gemacht werden, wenn es überhaupt vom Default abweichende Werte gibt 
            ' jetzt die vom Default abweichenden Werte speichern, sofern es welche gibt ... 

            If Not IsNothing(Me.capaPerMonth) Then
                Dim startingIndex As Integer = DateDiff(DateInterval.Month, StartofCalendar, Me.startOfCal.ToLocalTime) + 1

                If awinSettings.visboDebug Then
                    logger(ptErrLevel.logInfo, "clsRollenDefinitionWeb.copyto: ", "orgaUnit: " & Me.name & " - startingIndex: " & startingIndex)
                End If


                If startingIndex > 0 Then

                    ' ur: Änderung durch TSO Orga und separate Capa-Collection
                    'For i As Integer = startingIndex To startingIndex + nrWebCapaValues - 1
                    'roleDef.kapazitaet(i) = Me.kapazitaet(i - startingIndex + 1)
                    For i As Integer = startingIndex To startingIndex + nrWebCapaValues - 1
                        roleDef.kapazitaet(i) = Me.capaPerMonth(i - startingIndex)
                    Next
                Else ' ur:2020-11-20 - wenn später der startofcalendar im customization verschoben wurde
                    startingIndex = DateDiff(DateInterval.Month, Me.startOfCal.ToLocalTime, StartofCalendar) + 1

                    ' ur: Änderung durch TSO Orga und separate Capa-Collection
                    'For i As Integer = 1 To nrWebCapaValues - startingIndex
                    'roleDef.kapazitaet(i) = Me.kapazitaet(i + startingIndex - 1)
                    For i As Integer = 0 To nrWebCapaValues - startingIndex
                        roleDef.kapazitaet(i) = Me.capaPerMonth(i + startingIndex)
                    Next
                End If

            End If

        Else
            ' andernfalls beim kapazitaet(240): jeder Wert ist bereits Null, wie es sein soll ...
        End If




        ' ------------------------------------------------------
        ' ur: 25022022: tso-organization changes because of separated capas
        '-------------------------------------------------------
        '' jetzt die Übernahme der Kapazitäten 
        '' Rollen, die Kinder haben tragen niemals Kapa , also immer Null 
        '' ebenso Rollen, die nur den Default Wert haben 

        '' hier muss auch nur was gemacht werden, wenn subRoleId.count = 0 
        'If subRoleIDs.Count = 0 Then

        '        Dim nrWebCapaValues As Integer
        '        If Not IsNothing(Me.kapazitaet) Then
        '            nrWebCapaValues = Me.kapazitaet.Length - 1
        '        Else
        '            nrWebCapaValues = 0
        '        End If

        '        Dim lenSession As Integer = roleDef.kapazitaet.Length


        '        ' ' vorbesetzen mit dem Default Wert
        '        For i As Integer = 1 To lenSession - 1
        '            roleDef.kapazitaet(i) = roleDef.defaultKapa
        '        Next

        '        ' das muss jetzt nur gemacht werden, wenn es überhaupt vom Default abweichende Werte gibt 
        '        ' jetzt die vom Default abweichenden Werte speichern, sofern es welche gibt ... 

        '        If Not IsNothing(Me.kapazitaet) Then
        '            Dim startingIndex As Integer = DateDiff(DateInterval.Month, StartofCalendar, Me.startOfCal.ToLocalTime) + 1

        '            If awinSettings.visboDebug Then
        '                logger(ptErrLevel.logInfo, "clsRollenDefinitionWeb.copyto: ", "orgaUnit: " & Me.name & " - startingIndex: " & startingIndex)
        '            End If


        '            If startingIndex > 0 Then
        '                For i As Integer = startingIndex To startingIndex + nrWebCapaValues - 1
        '                    roleDef.kapazitaet(i) = Me.kapazitaet(i - startingIndex + 1)
        '                Next
        '            Else ' ur:2020-11-20 - wenn später der startofcalendar im customization verschoben wurde
        '                startingIndex = DateDiff(DateInterval.Month, Me.startOfCal.ToLocalTime, StartofCalendar) + 1
        '                For i As Integer = 1 To nrWebCapaValues - startingIndex
        '                    roleDef.kapazitaet(i) = Me.kapazitaet(i + startingIndex - 1)
        '                Next
        '            End If

        '        End If

        '    Else
        '        ' andernfalls beim kapazitaet(240): jeder Wert ist bereits Null, wie es sein soll ...
        '    End If

    End Sub

    Public Sub copyFrom(ByVal roleDef As clsRollenDefinition)

        With roleDef

            Dim dbKapa() As Double = Nothing
            ' damit wird festgelegt, ab wo im kapazitaet240 Array die dbKApa Werte zu platzieren sind ...
            Dim startOfNonStandardValues As Date = Date.MinValue

            ' jetzt die SubRoles übernehmen 
            If .getSubRoleCount >= 1 Then
                For Each kvp As KeyValuePair(Of Integer, Double) In .getSubRoleIDs
                    Dim sr As New clsSubRoleID
                    sr.key = kvp.Key
                    'sr.value = kvp.Value.ToString
                    sr.value = kvp.Value
                    Me.subRoleIDs.Add(sr)
                Next


            End If



            If .getSkillCount >= 1 Then
                For Each kvp As KeyValuePair(Of Integer, Double) In .getSkillIDs
                    Dim sr As New clsSubRoleID
                    sr.key = kvp.Key
                    'sr.value = kvp.Value.ToString
                    sr.value = kvp.Value
                    Me.teamIDs.Add(sr)
                Next
            End If

            uid = .UID
            name = .name

            isExternRole = .isExternRole

            ' ur:25022022: isTeam replaced with type=2 and isSummaryRole = true
            'isTeam = .isSkill

            If (.isSkill) Then
                type = 2
                isSummaryRole = True
            End If

            aliases = .aliases
            defCapaMonth = .defaultKapa
            defCapaDay = .defaultDayCapa
            employeeNr = .employeeNr
            entryDate = .entryDate.ToUniversalTime
            exitDate = .exitDate.ToUniversalTime

            dailyRate = .tagessatzIntern
            tagessatzIntern = dailyRate
            tagessatz = dailyRate

            ' ur: 8.7.21 neu mit VS-943
            isAggregationRole = .isAggregationRole
            isSummaryRole = .isCombinedRole Or .isSummaryRole

            ' ur:25.02.2022: moved in customization-VCsetting
            'isActDataRelevant = .isActDataRelevant

            '' tk 17.5.20 effiziente Organisation
            '' jetzt nur den Array übergeben, der die vom Default abweichenden Werte enthält 
            ''kapazitaet = .kapazitaet
            'kapazitaet = dbKapa
            '    ' dieser startOfCal gibt jetzt an, wo der Array genau zu beginnen hat ...
            '    startOfCal = startOfNonStandardValues.ToUniversalTime


        End With
    End Sub



    Public Sub New()
        subRoleIDs = New List(Of clsSubRoleID)
        teamIDs = New List(Of clsSubRoleID)

        type = 1
        isExternRole = False

        ' am 10.1. dazugekommen 
        aliases = Nothing
        employeeNr = ""
        defaultDayCapa = -1
        entryDate = Date.MinValue.ToUniversalTime
        'exitDate = CDate("31.12.2200").ToUniversalTime
        exitDate = DateAndTime.DateSerial(2200, 12, 31)
        Dim maxDate As Date = Date.MaxValue.ToUniversalTime
        isAggregationRole = False

        ' ur:25.02.2022: moved in customization-VCsetting
        'isActDataRelevant = False
        isSummaryRole = False

        startOfCal = StartofCalendar.ToUniversalTime
    End Sub

    Public Sub New(ByVal tmpDate As Date)
        subRoleIDs = New List(Of clsSubRoleID)
        teamIDs = New List(Of clsSubRoleID)

        isTeam = False
        isExternRole = False

        ' am 10.1. dazugekommen 
        aliases = Nothing
        employeeNr = ""
        defaultDayCapa = -1
        entryDate = Date.MinValue.ToUniversalTime
        'exitDate = CDate("2200.31.12").ToUniversalTime
        exitDate = DateAndTime.DateSerial(2200, 12, 31)
        isAggregationRole = False
        ' ur:25.02.2022: moved in customization-VCsetting
        'isActDataRelevant = False
        isSummaryRole = False

        startOfCal = StartofCalendar.ToUniversalTime
    End Sub

End Class
