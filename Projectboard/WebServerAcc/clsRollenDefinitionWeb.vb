Imports ProjectBoardDefinitions
Public Class clsRollenDefinitionWeb
    ' bei subRoleIDs eigentlich integer, string), muss wegen Mongo auf String geändert werden 
    ' tk 29.5.18 in den SubroleID values steht jetzt im String nicht mehr der Name, der ist ohnehin redundant zur UID, sondern der Prozentsatz, wieviel die Rolle zur Kapa der Sammelrolle beiträgt 
    ' wenn ein nicht als double interpretierbarer Wert drinsteht (=alte Speicherungen, dann wird der Wert auf 1.0 gesetzt 
    Public subRoleIDs As List(Of clsSubRoleID)

    ' 23.11.18 neu hinzugekommen 
    Public teamIDs As List(Of clsSubRoleID)

    ' 23.11.18 neu hinzugekommen 
    Public isExternRole As Boolean
    Public isTeam As Boolean

    ' 27.04.20 ur wird nun auch in der DB gespeichert
    '9.11.20 nicht mehr in DB speichern
    'Public isTeamParent As Boolean

    ' 8.1.2020 dazugekommen
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


    ' startOfCal ist wichtig, damit die korrekte Zuordnung der Kapa-Werte zu den Monaten gemacht werden kann 
    Public startOfCal As Date

    ' war vorher : Public Sub copyTo(ByRef roleDef As clsRollenDefinition, ByVal orgaStartOfCalendar As Date)
    Public Sub copyTo(ByRef roleDef As clsRollenDefinition)


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
                    roleDef.addTeam(CInt(sr.key), tmpValue)
                Catch ex As Exception
                    Call MsgBox("1119765: not allowed to to have team-Membership and Childs ..")
                End Try

            Next
        End If

        roleDef.UID = Me.uid
        roleDef.name = Me.name
        roleDef.farbe = Me.farbe
        roleDef.defaultKapa = Me.defaultKapa

        ' tk 8.1.20
        roleDef.aliases = Me.aliases
        roleDef.defaultDayCapa = Me.defaultDayCapa
        roleDef.employeeNr = Me.employeeNr
        roleDef.entryDate = Me.entryDate.ToLocalTime
        roleDef.exitDate = Me.exitDate.ToLocalTime


        ' tk 23.11.18 
        roleDef.isExternRole = Me.isExternRole
        roleDef.isTeam = Me.isTeam
        ' 9.11.20 ur for a smart change to tagessatz
        If Not IsNothing(Me.tagessatz) Then
            roleDef.tagessatzIntern = Me.tagessatz
        Else
            roleDef.tagessatzIntern = Me.tagessatzIntern
        End If

        ' jetzt die Übernahme der Kapazitäten 
        ' Rollen, die Kinder haben tragen niemals Kapa , also immer Null 
        ' ebenso Rollen, die nur den Default Wert haben 

        ' hier muss auch nur was gemacht werden, wenn subRoleId.count = 0 
        If subRoleIDs.Count = 0 Then

            Dim nrWebCapaValues As Integer
            If Not IsNothing(Me.kapazitaet) Then
                nrWebCapaValues = Me.kapazitaet.Length - 1
            Else
                nrWebCapaValues = 0
            End If

            Dim lenSession As Integer = roleDef.kapazitaet.Length


            ' ' vorbesetzen mit dem Default Wert
            For i As Integer = 1 To lenSession - 1
                roleDef.kapazitaet(i) = roleDef.defaultKapa
            Next

            ' das muss jetzt nur gemacht werden, wenn es überhaupt vom Default abweichende Werte gibt 
            ' jetzt die vom Default abweichenden Werte speichern, sofern es welche gibt ... 

            If Not IsNothing(Me.kapazitaet) Then
                Dim startingIndex As Integer = DateDiff(DateInterval.Month, StartofCalendar, Me.startOfCal.ToLocalTime) + 1

                For i As Integer = startingIndex To startingIndex + nrWebCapaValues - 1
                    roleDef.kapazitaet(i) = Me.kapazitaet(i - startingIndex + 1)
                Next
            End If

        Else
            ' andernfalls beim kapazitaet(240): jeder Wert ist bereits Null, wie es sein soll ...
        End If





        'If orgaStartOfCalendar <> Date.MinValue Then
        '    ' neue Variante 
        '    ' erst mal mit dem Default vorbesetzen 
        '    ' Neu 17.5.20
        '    For i As Integer = 1 To lenSession - 1
        '        roleDef.kapazitaet(i) = roleDef.defaultKapa
        '    Next

        '    ' jetzt die vom Default abweichenden Werte speichern ... 
        '    Dim startingIndex As Integer = DateDiff(DateInterval.Month, StartofCalendar, Me.startOfCal.ToLocalTime) + 1

        '    For i As Integer = startingIndex To lenDB
        '        roleDef.kapazitaet(i) = Me.kapazitaet(i - startingIndex)
        '    Next

        'Else
        '    ' alte Variante 
        '    Dim anzMon As Long = DateDiff(DateInterval.Month, Me.startOfCal.ToLocalTime, StartofCalendar)
        '    If anzMon = 0 Then
        '        '  aber vorher checken ob die Dimensionen gleich sind 

        '        If lenDB = lenSession Then
        '            ' einfach kopieren ...
        '            roleDef.kapazitaet = Me.kapazitaet
        '            '.externeKapazitaet = Me.externeKapazitaet
        '        ElseIf lenDB < lenSession Then
        '            For i As Integer = 0 To lenDB
        '                roleDef.kapazitaet(i) = Me.kapazitaet(i)
        '                '.externeKapazitaet(i) = Me.externeKapazitaet(i)
        '            Next
        '            ' jetzt hinten auffüllen ..
        '            For i As Integer = lenDB + 1 To lenSession - 1
        '                roleDef.kapazitaet(i) = Me.defaultKapa
        '                '.externeKapazitaet(i) = 0
        '            Next
        '        Else
        '            For i As Integer = 0 To lenSession - 1
        '                roleDef.kapazitaet(i) = Me.kapazitaet(i)
        '                '.externeKapazitaet(i) = Me.externeKapazitaet(i)
        '            Next
        '        End If

        '    ElseIf anzMon < 0 Then
        '        ' der StartOfCalendar wurde in der Multiprojekt-Tafel mittlerweile nach vorne verschoben 
        '        ' also vorne auffülen
        '        anzMon = -1 * anzMon

        '        If lenDB = lenSession Then
        '            For i As Integer = 0 To CInt(anzMon)
        '                roleDef.kapazitaet(i) = Me.defaultKapa
        '                '.externeKapazitaet(i) = 0
        '            Next

        '            For i As Integer = CInt(anzMon + 1) To lenSession - 1
        '                roleDef.kapazitaet(i) = Me.kapazitaet(i - CInt(anzMon))
        '                '.externeKapazitaet(i) = Me.externeKapazitaet((i - CInt(anzMon)))
        '            Next
        '        ElseIf lenDB < lenSession Then
        '            ' Länge in der Datenbank ist kleiner als Länger in der Session 
        '            For i As Integer = 0 To CInt(anzMon)
        '                roleDef.kapazitaet(i) = Me.defaultKapa
        '                '.externeKapazitaet(i) = 0
        '            Next

        '            For i As Integer = CInt(anzMon + 1) To lenDB - 1
        '                roleDef.kapazitaet(i) = Me.kapazitaet(i - CInt(anzMon))
        '                '.externeKapazitaet(i) = Me.externeKapazitaet((i - CInt(anzMon)))
        '            Next

        '            For i As Integer = lenDB To lenSession - 1
        '                roleDef.kapazitaet(i) = Me.defaultKapa
        '                '.externeKapazitaet(i) = 0
        '            Next


        '        Else
        '            ' Länge in der Datenbank ist größer als Länge in der Session 
        '            For i As Integer = 0 To CInt(anzMon)
        '                roleDef.kapazitaet(i) = Me.defaultKapa
        '                '.externeKapazitaet(i) = 0
        '            Next

        '            For i As Integer = CInt(anzMon + 1) To lenSession - 1
        '                roleDef.kapazitaet(i) = Me.kapazitaet(i - CInt(anzMon))
        '                '.externeKapazitaet(i) = Me.externeKapazitaet((i - CInt(anzMon)))
        '            Next

        '        End If



        '    Else
        '        ' der StartOfCalendar wurde in der Multiprojekt-Tafel mittlerweile nach hinten verschoben 
        '        ' also ggf. hinten auffüllen 

        '        If lenDB = lenSession Then

        '            ' eas Null-Element hat keine Bedeutung 
        '            roleDef.kapazitaet(0) = 0
        '            '.externeKapazitaet(0) = 0

        '            For i As Integer = 1 To CInt(lenSession - anzMon - 1)
        '                roleDef.kapazitaet(i) = Me.kapazitaet(i + CInt(anzMon))
        '                '.externeKapazitaet(i) = Me.externeKapazitaet((i + CInt(anzMon)))
        '            Next

        '            For i As Integer = CInt(lenSession - anzMon) To lenSession - 1
        '                roleDef.kapazitaet(i) = Me.defaultKapa
        '                '.externeKapazitaet(i) = 0
        '            Next

        '        ElseIf lenDB < lenSession Then
        '            ' Länge in der Datenbank ist kleiner als Länge in der Session 
        '            For i As Integer = 0 To CInt(anzMon)
        '                roleDef.kapazitaet(i) = Me.defaultKapa
        '                '.externeKapazitaet(i) = 0
        '            Next

        '            For i As Integer = CInt(anzMon + 1) To lenDB - 1 - CInt(anzMon)
        '                roleDef.kapazitaet(i) = Me.kapazitaet(i + CInt(anzMon))
        '                '.externeKapazitaet(i) = Me.externeKapazitaet((i + CInt(anzMon)))
        '            Next

        '            For i As Integer = lenDB To lenSession - 1
        '                roleDef.kapazitaet(i) = Me.defaultKapa
        '                '.externeKapazitaet(i) = 0
        '            Next


        '        Else
        '            ' Länge in der Datenbank ist größer als Länge in der Session 
        '            For i As Integer = 0 To CInt(anzMon)
        '                roleDef.kapazitaet(i) = Me.defaultKapa
        '                '.externeKapazitaet(i) = 0
        '            Next

        '            For i As Integer = CInt(anzMon + 1) To lenSession - 1 - CInt(anzMon)
        '                roleDef.kapazitaet(i) = Me.kapazitaet(i + CInt(anzMon))
        '                '.externeKapazitaet(i) = Me.externeKapazitaet((i + CInt(anzMon)))
        '            Next

        '        End If


        '    End If

        '    ' Ende alte Variante 
        'End If



        ' alt 17.5.20


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

            Else
                ' das kann nur bei Blättern der Fall sein, alle übergeordneten Orga-Units, also solche die Kinder haben, bekommen ihre Kapa aus den "Blättern"
                ' nachsehen, ob es irgendwelche Non-Default Kapa Werte gibt 
                Dim anzahlMonate As Integer = roleDef.kapazitaet.Length - 1


                Dim startingIndex As Integer = -1
                Dim endingIndex As Integer = anzahlMonate + 1



                For i As Integer = 1 To anzahlMonate
                    If roleDef.kapazitaet(i) <> roleDef.defaultKapa Then
                        startingIndex = i
                        Exit For
                    End If
                Next

                If startingIndex = -1 Then
                    ' alle Kapa-Werte sind Standard 
                    ' das heisst man kann es bei den Voreinstellungen lassen 
                    ' 
                    dbKapa = Nothing
                    startOfNonStandardValues = Date.MinValue

                Else
                    ' startingIndex kann jetzt nur Werte zwischen 1 und 240 haben ..
                    startOfNonStandardValues = StartofCalendar.AddMonths(startingIndex - 1)

                    endingIndex = anzahlMonate

                    For i As Integer = anzahlMonate To startingIndex Step -1
                        If roleDef.kapazitaet(i) <> roleDef.defaultKapa Then
                            endingIndex = i
                            Exit For
                        End If
                    Next

                    ' es soll wie bei kapazitaet(240) sein kapazitaet (0) ist nicht relevant, es beginnt bei dbKapa(1) 
                    Dim dbDim As Integer = endingIndex - startingIndex + 1

                    ReDim dbKapa(dbDim)

                    ' Array aufbauen 
                    For i As Integer = 1 To dbDim
                        dbKapa(i) = roleDef.kapazitaet(i + startingIndex - 1)
                    Next

                End If

            End If

            If .getTeamCount >= 1 Then
                For Each kvp As KeyValuePair(Of Integer, Double) In .getTeamIDs
                    Dim sr As New clsSubRoleID
                    sr.key = kvp.Key
                    'sr.value = kvp.Value.ToString
                    sr.value = kvp.Value
                    Me.teamIDs.Add(sr)
                Next
            End If

            uid = .UID
            name = .name
            farbe = CLng(.farbe)
            defaultKapa = .defaultKapa

            ' tk 23.11.18 
            isExternRole = .isExternRole
            isTeam = .isTeam
            ' ur 27.04.20 
            ' ur 9.11.20 wird nicht mehr in DB gespeicher
            'isTeamParent = .isTeamParent

            ' tk 8.1.20
            aliases = .aliases
            defaultDayCapa = .defaultDayCapa
            employeeNr = .employeeNr
            entryDate = .entryDate.ToUniversalTime
            exitDate = .exitDate.ToUniversalTime

            tagessatz = .tagessatzIntern



            ' tk 17.5.20 effiziente Organisation
            ' jetzt nur den Array übergeben, der die vom Default abweichenden Werte enthält 
            'kapazitaet = .kapazitaet
            kapazitaet = dbKapa
            ' dieser startOfCal gibt jetzt an, wo der Array genau zu beginnen hat ...
            startOfCal = startOfNonStandardValues.ToUniversalTime


        End With
    End Sub

    '''' <summary>
    '''' true, if both Roledefinitions are identical , except timestamp 
    '''' </summary>
    '''' <param name="vglRole"></param>
    '''' <value></value>
    '''' <returns></returns>
    '''' <remarks></remarks>
    ' tk 8.1.2020 : wird nie aufgerufen , deswegen auskommentiert
    'Public ReadOnly Property isIdenticalTo(ByVal vglRole As clsRollenDefinitionWeb) As Boolean
    '    Get
    '        Dim stillok As Boolean = True

    '        If Me.subRoleIDs.Count = vglRole.subRoleIDs.Count Then
    '            If Me.subRoleIDs.Count = 0 Then
    '                stillok = True
    '            Else
    '                Dim i As Integer = 0
    '                Do While i < Me.subRoleIDs.Count And stillok
    '                    stillok = (Me.subRoleIDs.ElementAt(i).key = vglRole.subRoleIDs.ElementAt(i).key And
    '                               Me.subRoleIDs.ElementAt(i).value = vglRole.subRoleIDs.ElementAt(i).value)
    '                    i = i + 1
    '                Loop

    '                i = 0
    '                Do While i < Me.teamIDs.Count And stillok
    '                    stillok = (Me.teamIDs.ElementAt(i).key = vglRole.teamIDs.ElementAt(i).key And
    '                               Me.teamIDs.ElementAt(i).value = vglRole.teamIDs.ElementAt(i).value)
    '                    i = i + 1
    '                Loop
    '            End If
    '        Else
    '            stillok = False
    '        End If


    '        ' jetzt alle anderen Attribute überprüfen ...
    '        If stillok Then

    '            stillok = (Me.uid = vglRole.uid) And
    '                        (Me.name = vglRole.name) And
    '                        (Me.farbe = vglRole.farbe) And
    '                        (Me.defaultKapa = vglRole.defaultKapa) And
    '                        (Me.isExternRole = vglRole.isExternRole) And
    '                        (Me.isTeam = vglRole.isTeam) And
    '                        (Me.tagessatzIntern = vglRole.tagessatzIntern) And
    '                        (Me.employeeNr = vglRole.employeeNr) And
    '                        (Me.entryDate.Date = vglRole.entryDate.Date) And
    '                        (Me.exitDate.Date = vglRole.exitDate.Date) And
    '                        (Me.defaultDayCapa = vglRole.defaultDayCapa)


    '        End If

    '        ' jetzt die Kapa-Arrays vergleichen 
    '        If stillok Then
    '            stillok = Not arraysAreDifferent(Me.kapazitaet, vglRole.kapazitaet)
    '            'And
    '            '            Not arraysAreDifferent(Me.externeKapazitaet, vglRole.externeKapazitaet)
    '        End If

    '        isIdenticalTo = stillok

    '    End Get
    'End Property

    Public Sub New()
        subRoleIDs = New List(Of clsSubRoleID)
        teamIDs = New List(Of clsSubRoleID)

        isTeam = False
        isExternRole = False

        ' am 10.1. dazugekommen 
        aliases = Nothing
        employeeNr = ""
        defaultDayCapa = -1
        entryDate = Date.MinValue.ToUniversalTime
        exitDate = CDate("31.12.2200").ToUniversalTime

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
        exitDate = CDate("31.12.2200").ToUniversalTime

        startOfCal = StartofCalendar.ToUniversalTime
    End Sub

End Class
