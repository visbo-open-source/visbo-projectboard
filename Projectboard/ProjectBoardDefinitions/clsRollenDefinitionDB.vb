Public Class clsRollenDefinitionDB
    ' bei subRoleIDs eigentlich integer, string), muss wegen Mongo auf String geändert werden 
    ' tk 29.5.18 in den SubroleID values steht jetzt im String nicht mehr der Name, der ist ohnehin redundant zur UID, sondern der Prozentsatz, wieviel die Rolle zur Kapa der Sammelrolle beiträgt 
    ' wenn ein nicht als double interpretierbarer Wert drinsteht (=alte Speicherungen, dann wird der Wert auf String 1.0 gesetzt 
    Public subRoleIDs As SortedList(Of String, String)

    ' tk 21.11.18 jetzt sind noch die TeamIDs mit drin  
    Public teamIDs As SortedList(Of String, String)

    Public isExternRole As Boolean
    Public isTeam As Boolean

    Public uid As Integer
    Public name As String
    Public farbe As Long
    Public defaultKapa As Double
    Public tagessatzIntern As Double
    Public kapazitaet() As Double


    Public timestamp As Date
    ' Id wird von MongoDB automatisch gesetzt 
    Public Id As String

    ' startOfCal ist wichtig, damit die korrekte Zuordnung der Kapa-Werte zu den Monaten gemacht werden kann 
    Public startOfCal As Date

    Public Sub copyTo(ByRef roleDef As clsRollenDefinition)

        With roleDef
            If subRoleIDs.Count >= 1 Then
                ' wegen Mongo müssen die Keys in String Format sein ... 

                For Each kvp As KeyValuePair(Of String, String) In Me.subRoleIDs
                    Dim tmpValue As Double = 1.0
                    If IsNumeric(kvp.Value) Then
                        tmpValue = CDbl(kvp.Value)
                        If tmpValue >= 0 And tmpValue <= 1.0 Then
                            ' alles ok
                        Else
                            tmpValue = 1.0
                        End If
                    Else
                        tmpValue = 1.0
                    End If

                    Try
                        .addSubRole(CInt(kvp.Key), tmpValue)
                    Catch ex As Exception
                        Call MsgBox("1119765: not allowed to to have team-Membership and Childs ..")
                    End Try

                Next

            End If

            ' Allianz 23.11.18 jetzt die TeamIDs kopieren 
            If teamIDs.Count >= 1 Then
                ' wegen Mongo müssen die Keys in String Format sein ... 

                For Each kvp As KeyValuePair(Of String, String) In Me.teamIDs
                    Dim tmpValue As Double = 1.0
                    If IsNumeric(kvp.Value) Then
                        tmpValue = CDbl(kvp.Value)
                        If tmpValue >= 0 And tmpValue <= 1.0 Then
                            ' alles ok
                        Else
                            tmpValue = 1.0
                        End If
                    Else
                        tmpValue = 1.0
                    End If

                    Try
                        .addTeam(CInt(kvp.Key), tmpValue)
                    Catch ex As Exception
                        Call MsgBox("1119765: not allowed to to have team-Membership and Childs ..")
                    End Try

                Next

            End If
            ' Ende 23.11.18 


            .UID = Me.uid
            .name = Me.name

            ' 23.11.18 
            .isExternRole = Me.isExternRole
            .isTeam = Me.isTeam

            .farbe = Me.farbe
            .defaultKapa = Me.defaultKapa


            .tagessatzIntern = Me.tagessatzIntern

            Dim lenDB As Integer = Me.kapazitaet.Length
            Dim lenSession As Integer = .kapazitaet.Length

            Dim anzMon As Long = DateDiff(DateInterval.Month, Me.startOfCal.ToLocalTime, StartofCalendar)
            If anzMon = 0 Then
                '  aber vorher checken ob die Dimensionen gleich sind 

                If lenDB = lenSession Then
                    ' einfach kopieren ...
                    .kapazitaet = Me.kapazitaet
                    ' tk Allianz 21.11.18 nicht mehr nötig 
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
                    Me.subRoleIDs.Add(CStr(kvp.Key), kvp.Value.ToString)
                Next
            End If

            ' tk 23.11.18 
            If .getTeamCount >= 1 Then
                For Each kvp As KeyValuePair(Of Integer, Double) In .getTeamIDs
                    Me.teamIDs.Add(CStr(kvp.Key), kvp.Value.ToString)
                Next
            End If

            Me.uid = .UID
            Me.name = .name
            Me.farbe = CLng(.farbe)
            Me.defaultKapa = .defaultKapa

            ' 23.11.18 
            Me.isExternRole = .isExternRole
            Me.isTeam = .isTeam

            Me.tagessatzIntern = .tagessatzIntern
            Me.kapazitaet = .kapazitaet


            Me.Id = "Role" & "#" & CStr(Me.uid) & "#" & Date.UtcNow.ToString

        End With
    End Sub

    ''' <summary>
    ''' true, if both Roledefinitions are identical , except timestamp 
    ''' </summary>
    ''' <param name="vglRole"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property isIdenticalTo(ByVal vglRole As clsRollenDefinitionDB) As Boolean
        Get
            Dim stillok As Boolean = True

            If Me.subRoleIDs.Count = vglRole.subRoleIDs.Count Then
                If Me.subRoleIDs.Count = 0 Then
                    stillok = True
                Else
                    Dim i As Integer = 0
                    Do While i < Me.subRoleIDs.Count And stillok
                        stillok = (Me.subRoleIDs.ElementAt(i).Key = vglRole.subRoleIDs.ElementAt(i).Key And
                                   Me.subRoleIDs.ElementAt(i).Value = vglRole.subRoleIDs.ElementAt(i).Value)
                        i = i + 1
                    Loop

                    i = 0
                    Do While i < Me.teamIDs.Count And stillok
                        stillok = (Me.teamIDs.ElementAt(i).Key = vglRole.teamIDs.ElementAt(i).Key And
                                   Me.teamIDs.ElementAt(i).Value = vglRole.teamIDs.ElementAt(i).Value)
                        i = i + 1
                    Loop

                End If
            Else
                stillok = False
            End If


            ' jetzt alle anderen Attribute überprüfen ...
            If stillok Then

                stillok = (Me.uid = vglRole.uid) And
                            (Me.name = vglRole.name) And
                            (Me.farbe = vglRole.farbe) And
                            (Me.defaultKapa = vglRole.defaultKapa) And
                            (Me.isExternRole = vglRole.isExternRole) And
                            (Me.isTeam = vglRole.isTeam) And
                            (Me.tagessatzIntern = vglRole.tagessatzIntern)


            End If

            ' jetzt die Kapa-Arrays vergleichen 
            If stillok Then
                stillok = Not arraysAreDifferent(Me.kapazitaet, vglRole.kapazitaet)

            End If

            isIdenticalTo = stillok

        End Get
    End Property

    Public Sub New()
        subRoleIDs = New SortedList(Of String, String)
        teamIDs = New SortedList(Of String, String)

        isTeam = False
        isExternRole = False

        timestamp = Date.UtcNow
        startOfCal = StartofCalendar.ToUniversalTime
        Id = ""

    End Sub

    Public Sub New(ByVal tmpDate As Date)
        subRoleIDs = New SortedList(Of String, String)
        teamIDs = New SortedList(Of String, String)

        isTeam = False
        isExternRole = False

        timestamp = Date.UtcNow
        startOfCal = StartofCalendar.ToUniversalTime
        Id = ""
    End Sub

End Class
