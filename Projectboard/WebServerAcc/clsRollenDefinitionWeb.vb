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

    Public kapazitaet() As Double

    ' tk 23.11. nicht mehr relevant, bleibt drin, um alte Datenmodelle behandeln zu können 
    Public tagessatzExtern As Double = Nothing
    Public externeKapazitaet() As Double = Nothing



    Public timestamp As Date

    ' startOfCal ist wichtig, damit die korrekte Zuordnung der Kapa-Werte zu den Monaten gemacht werden kann 
    Public startOfCal As Date

    Public Sub copyTo(ByRef roleDef As clsRollenDefinition)

        With roleDef
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
                        .addSubRole(CInt(sr.key), tmpValue)
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
                        .addTeam(CInt(sr.key), tmpValue)
                    Catch ex As Exception
                        Call MsgBox("1119765: not allowed to to have team-Membership and Childs ..")
                    End Try

                Next
            End If

            .UID = Me.uid
            .name = Me.name
            .farbe = Me.farbe
            .defaultKapa = Me.defaultKapa

            ' tk 8.1.20
            .aliases = Me.aliases
            .defaultDayCapa = Me.defaultDayCapa
            .employeeNr = Me.employeeNr
            .entryDate = Me.entryDate
            .exitDate = Me.exitDate


            ' tk 23.11.18 
            .isExternRole = Me.isExternRole
            .isTeam = Me.isTeam

            .tagessatzIntern = Me.tagessatzIntern

            '.tagessatzExtern = Me.tagessatzExtern
            Dim lenDB As Integer = Me.kapazitaet.Length
            Dim lenSession As Integer = .kapazitaet.Length

            Dim anzMon As Long = DateDiff(DateInterval.Month, Me.startOfCal.ToLocalTime, StartofCalendar)
            If anzMon = 0 Then
                '  aber vorher checken ob die Dimensionen gleich sind 

                If lenDB = lenSession Then
                    ' einfach kopieren ...
                    .kapazitaet = Me.kapazitaet
                    '.externeKapazitaet = Me.externeKapazitaet
                ElseIf lenDB < lenSession Then
                    For i As Integer = 0 To lenDB
                        .kapazitaet(i) = Me.kapazitaet(i)
                        '.externeKapazitaet(i) = Me.externeKapazitaet(i)
                    Next
                    ' jetzt hinten auffüllen ..
                    For i As Integer = lenDB + 1 To lenSession - 1
                        .kapazitaet(i) = Me.defaultKapa
                        '.externeKapazitaet(i) = 0
                    Next
                Else
                    For i As Integer = 0 To lenSession - 1
                        .kapazitaet(i) = Me.kapazitaet(i)
                        '.externeKapazitaet(i) = Me.externeKapazitaet(i)
                    Next
                End If

            ElseIf anzMon < 0 Then
                ' der StartOfCalendar wurde in der Multiprojekt-Tafel mittlerweile nach vorne verschoben 
                ' also vorne auffülen
                anzMon = -1 * anzMon

                If lenDB = lenSession Then
                    For i As Integer = 0 To CInt(anzMon)
                        .kapazitaet(i) = Me.defaultKapa
                        '.externeKapazitaet(i) = 0
                    Next

                    For i As Integer = CInt(anzMon + 1) To lenSession - 1
                        .kapazitaet(i) = Me.kapazitaet(i - CInt(anzMon))
                        '.externeKapazitaet(i) = Me.externeKapazitaet((i - CInt(anzMon)))
                    Next
                ElseIf lenDB < lenSession Then
                    ' Länge in der Datenbank ist kleiner als Länger in der Session 
                    For i As Integer = 0 To CInt(anzMon)
                        .kapazitaet(i) = Me.defaultKapa
                        '.externeKapazitaet(i) = 0
                    Next

                    For i As Integer = CInt(anzMon + 1) To lenDB - 1
                        .kapazitaet(i) = Me.kapazitaet(i - CInt(anzMon))
                        '.externeKapazitaet(i) = Me.externeKapazitaet((i - CInt(anzMon)))
                    Next

                    For i As Integer = lenDB To lenSession - 1
                        .kapazitaet(i) = Me.defaultKapa
                        '.externeKapazitaet(i) = 0
                    Next


                Else
                    ' Länge in der Datenbank ist größer als Länge in der Session 
                    For i As Integer = 0 To CInt(anzMon)
                        .kapazitaet(i) = Me.defaultKapa
                        '.externeKapazitaet(i) = 0
                    Next

                    For i As Integer = CInt(anzMon + 1) To lenSession - 1
                        .kapazitaet(i) = Me.kapazitaet(i - CInt(anzMon))
                        '.externeKapazitaet(i) = Me.externeKapazitaet((i - CInt(anzMon)))
                    Next

                End If



            Else
                ' der StartOfCalendar wurde in der Multiprojekt-Tafel mittlerweile nach nach hinten verschoben 
                ' also ggf. hinten auffüllen 

                If lenDB = lenSession Then

                    ' eas Null-Element hat keine Bedeutung 
                    .kapazitaet(0) = 0
                    '.externeKapazitaet(0) = 0

                    For i As Integer = 1 To CInt(lenSession - anzMon - 1)
                        .kapazitaet(i) = Me.kapazitaet(i + CInt(anzMon))
                        '.externeKapazitaet(i) = Me.externeKapazitaet((i + CInt(anzMon)))
                    Next

                    For i As Integer = CInt(lenSession - anzMon) To lenSession - 1
                        .kapazitaet(i) = Me.defaultKapa
                        '.externeKapazitaet(i) = 0
                    Next

                ElseIf lenDB < lenSession Then
                    ' Länge in der Datenbank ist kleiner als Länge in der Session 
                    For i As Integer = 0 To CInt(anzMon)
                        .kapazitaet(i) = Me.defaultKapa
                        '.externeKapazitaet(i) = 0
                    Next

                    For i As Integer = CInt(anzMon + 1) To lenDB - 1 - CInt(anzMon)
                        .kapazitaet(i) = Me.kapazitaet(i + CInt(anzMon))
                        '.externeKapazitaet(i) = Me.externeKapazitaet((i + CInt(anzMon)))
                    Next

                    For i As Integer = lenDB To lenSession - 1
                        .kapazitaet(i) = Me.defaultKapa
                        '.externeKapazitaet(i) = 0
                    Next


                Else
                    ' Länge in der Datenbank ist größer als Länge in der Session 
                    For i As Integer = 0 To CInt(anzMon)
                        .kapazitaet(i) = Me.defaultKapa
                        '.externeKapazitaet(i) = 0
                    Next

                    For i As Integer = CInt(anzMon + 1) To lenSession - 1 - CInt(anzMon)
                        .kapazitaet(i) = Me.kapazitaet(i + CInt(anzMon))
                        '.externeKapazitaet(i) = Me.externeKapazitaet((i + CInt(anzMon)))
                    Next

                End If


            End If
        End With
    End Sub

    Public Sub copyFrom(ByVal roleDef As clsRollenDefinition)
        With roleDef

            If .getSubRoleCount >= 1 Then
                For Each kvp As KeyValuePair(Of Integer, Double) In .getSubRoleIDs
                    Dim sr As New clsSubRoleID
                    sr.key = kvp.Key
                    sr.value = kvp.Value.ToString
                    Me.subRoleIDs.Add(sr)
                Next
            End If

            If .getTeamCount >= 1 Then
                For Each kvp As KeyValuePair(Of Integer, Double) In .getTeamIDs
                    Dim sr As New clsSubRoleID
                    sr.key = kvp.Key
                    sr.value = kvp.Value.ToString
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

            ' tk 8.1.20
            aliases = .aliases
            defaultDayCapa = .defaultDayCapa
            employeeNr = .employeeNr
            entryDate = .entryDate
            exitDate = .exitDate

            tagessatzIntern = .tagessatzIntern
            kapazitaet = .kapazitaet

            ' tk 3.12.18 wird nicht mehr benötigt ...
            tagessatzExtern = Nothing
            externeKapazitaet = Nothing
            ' Id wird beim Server von der MongoDB selbst erzeugt
            'Me.Id = "Role" & "#" & CStr(Me.uid) & "#" & Date.UtcNow.ToString

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
        entryDate = Date.MinValue
        exitDate = CDate("31.12.2200")

        timestamp = Date.UtcNow
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
        entryDate = Date.MinValue
        exitDate = CDate("31.12.2200")

        timestamp = Date.UtcNow
        startOfCal = StartofCalendar.ToUniversalTime
    End Sub

End Class
