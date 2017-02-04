Public Class clsRollenDefinitionDB

    Public subRoleIDs As SortedList(Of Integer, String)
    Public uid As Integer
    Public name As String
    Public farbe As Long
    Public defaultKapa As Double
    Public tagessatzIntern As Double
    Public tagessatzExtern As Double
    Public kapazitaet() As Double
    Public externeKapazitaet() As Double
    Public timeStamp As Date

    ' startOfCal ist wichtig, damit die korrekte Zuordnung der Kapa-Werte zu den Monaten gemacht werden kann 
    Public startOfCal As Date

    Public Sub copyTo(ByRef costDef As clsRollenDefinition)

        With costDef
            If subRoleIDs.Count >= 1 Then
                Dim maxNr As Integer = 1000
                For Each kvp As KeyValuePair(Of Integer, String) In subRoleIDs
                    .addSubRole(kvp.Key, kvp.Value, maxNr)
                Next
            End If
            .UID = Me.uid
            .name = Me.name
            .farbe = Me.farbe
            .defaultKapa = Me.defaultKapa
            .tagessatzIntern = Me.tagessatzIntern
            .tagessatzExtern = Me.tagessatzExtern
            Dim lenDB As Integer = Me.kapazitaet.Length
            Dim lenSession As Integer = .kapazitaet.Length

            Dim anzMon As Long = DateDiff(DateInterval.Month, Me.startOfCal.ToLocalTime, StartofCalendar)
            If anzMon = 0 Then
                '  aber vorher checken ob die Dimensionen gleich sind 

                If lenDB = lenSession Then
                    ' einfach kopieren ...
                    .kapazitaet = Me.kapazitaet
                    .externeKapazitaet = Me.externeKapazitaet
                ElseIf lenDB < lenSession Then
                    For i As Integer = 0 To lenDB
                        .kapazitaet(i) = Me.kapazitaet(i)
                        .externeKapazitaet(i) = Me.externeKapazitaet(i)
                    Next
                    ' jetzt hinten auffüllen ..
                    For i As Integer = lenDB + 1 To lenSession - 1
                        .kapazitaet(i) = Me.defaultKapa
                        .externeKapazitaet(i) = 0
                    Next
                Else
                    For i As Integer = 0 To lenSession - 1
                        .kapazitaet(i) = Me.kapazitaet(i)
                        .externeKapazitaet(i) = Me.externeKapazitaet(i)
                    Next
                End If

            ElseIf anzMon < 0 Then
                ' der StartOfCalendar wurde in der Multiprojekt-Tafel mittlerweile nach vorne verschoben 
                ' also vorne auffülen
                anzMon = -1 * anzMon

                If lenDB = lenSession Then
                    For i As Integer = 0 To CInt(anzMon)
                        .kapazitaet(i) = Me.defaultKapa
                        .externeKapazitaet(i) = 0
                    Next

                    For i As Integer = CInt(anzMon + 1) To lenSession - 1
                        .kapazitaet(i) = Me.kapazitaet(i - CInt(anzMon))
                        .externeKapazitaet(i) = Me.externeKapazitaet((i - CInt(anzMon)))
                    Next
                ElseIf lenDB < lenSession Then
                    ' Länge in der Datenbank ist kleiner als Länger in der Session 
                    For i As Integer = 0 To CInt(anzMon)
                        .kapazitaet(i) = Me.defaultKapa
                        .externeKapazitaet(i) = 0
                    Next

                    For i As Integer = CInt(anzMon + 1) To lenDB - 1
                        .kapazitaet(i) = Me.kapazitaet(i - CInt(anzMon))
                        .externeKapazitaet(i) = Me.externeKapazitaet((i - CInt(anzMon)))
                    Next

                    For i As Integer = lenDB To lenSession - 1
                        .kapazitaet(i) = Me.defaultKapa
                        .externeKapazitaet(i) = 0
                    Next


                Else
                    ' Länge in der Datenbank ist größer als Länge in der Session 
                    For i As Integer = 0 To CInt(anzMon)
                        .kapazitaet(i) = Me.defaultKapa
                        .externeKapazitaet(i) = 0
                    Next

                    For i As Integer = CInt(anzMon + 1) To lenSession - 1
                        .kapazitaet(i) = Me.kapazitaet(i - CInt(anzMon))
                        .externeKapazitaet(i) = Me.externeKapazitaet((i - CInt(anzMon)))
                    Next

                End If



            Else
                ' der StartOfCalendar wurde in der Multiprojekt-Tafel mittlerweile nach nach hinten verschoben 
                ' also ggf. hinten auffüllen 

                If lenDB = lenSession Then

                    ' eas Null-Element hat keine Bedeutung 
                    .kapazitaet(0) = 0
                    .externeKapazitaet(0) = 0

                    For i As Integer = 1 To CInt(lenSession - anzMon)
                        .kapazitaet(i) = Me.kapazitaet(i + CInt(anzMon))
                        .externeKapazitaet(i) = Me.externeKapazitaet((i + CInt(anzMon)))
                    Next

                    For i As Integer = CInt(lenSession - anzMon + 1) To lenSession - 1
                        .kapazitaet(i) = Me.defaultKapa
                        .externeKapazitaet(i) = 0
                    Next

                ElseIf lenDB < lenSession Then
                    ' Länge in der Datenbank ist kleiner als Länge in der Session 
                    For i As Integer = 0 To CInt(anzMon)
                        .kapazitaet(i) = Me.defaultKapa
                        .externeKapazitaet(i) = 0
                    Next

                    For i As Integer = CInt(anzMon + 1) To lenDB - 1 - CInt(anzMon)
                        .kapazitaet(i) = Me.kapazitaet(i + CInt(anzMon))
                        .externeKapazitaet(i) = Me.externeKapazitaet((i + CInt(anzMon)))
                    Next

                    For i As Integer = lenDB To lenSession - 1
                        .kapazitaet(i) = Me.defaultKapa
                        .externeKapazitaet(i) = 0
                    Next


                Else
                    ' Länge in der Datenbank ist größer als Länge in der Session 
                    For i As Integer = 0 To CInt(anzMon)
                        .kapazitaet(i) = Me.defaultKapa
                        .externeKapazitaet(i) = 0
                    Next

                    For i As Integer = CInt(anzMon + 1) To lenSession - 1 - CInt(anzMon)
                        .kapazitaet(i) = Me.kapazitaet(i + CInt(anzMon))
                        .externeKapazitaet(i) = Me.externeKapazitaet((i + CInt(anzMon)))
                    Next

                End If


            End If
        End With
    End Sub

    Public Sub copyFrom(ByVal costDef As clsRollenDefinition)
        With costDef

            If .getSubRoleCount >= 1 Then
                For Each kvp As KeyValuePair(Of Integer, String) In .getSubRoleIDs
                    Me.subRoleIDs.Add(kvp.Key, kvp.Value)
                Next
            End If

            Me.uid = .UID
            Me.name = .name
            Me.farbe = CLng(.farbe)
            Me.defaultKapa = .defaultKapa
            Me.tagessatzIntern = .tagessatzIntern
            Me.tagessatzExtern = .tagessatzExtern
            Me.kapazitaet = .kapazitaet
            Me.externeKapazitaet = .externeKapazitaet

        End With
    End Sub

    ''' <summary>
    ''' true, if both Roledefinitions are identical , except timestamp 
    ''' </summary>
    ''' <param name="vglRole"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property isIdenticalTo(ByVal vglRole As clsRollenDefinitionDB)
        Get
            Dim stillok As Boolean = True

            If Me.subRoleIDs.Count = vglRole.subRoleIDs.Count Then
                If Me.subRoleIDs.Count = 0 Then
                    stillok = True
                Else
                    Dim i As Integer = 0
                    Do While i < Me.subRoleIDs.Count And stillok
                        i = i + 1
                        stillok = (Me.subRoleIDs.ElementAt(i).Key = vglRole.subRoleIDs.ElementAt(i).Key And _
                                   Me.subRoleIDs.ElementAt(i).Value = vglRole.subRoleIDs.ElementAt(i).Value)
                    Loop
                    
                End If
            Else
                stillok = False
            End If


            ' jetzt alle anderen Attribute überprüfen ...
            If stillok Then

                stillok = (Me.uid = vglRole.uid) And _
                            (Me.name = vglRole.name) And _
                            (Me.farbe = vglRole.farbe) And _
                            (Me.defaultKapa = vglRole.defaultKapa) And _
                            (Me.tagessatzIntern = vglRole.tagessatzIntern) And _
                            (Me.tagessatzExtern = vglRole.tagessatzExtern)

            End If

            ' jetzt die Kapa-Arrays vergleichen 
            If stillok Then
                stillok = Not arraysAreDifferent(Me.kapazitaet, vglRole.kapazitaet) And _
                            Not arraysAreDifferent(Me.externeKapazitaet, vglRole.externeKapazitaet)
            End If

            isIdenticalTo = stillok

        End Get
    End Property

    Public Sub New()
        subRoleIDs = New SortedList(Of Integer, String)
        timeStamp = Date.UtcNow
        startOfCal = StartofCalendar.ToUniversalTime
    End Sub

    Public Sub New(ByVal tmpDate As Date)
        subRoleIDs = New SortedList(Of Integer, String)
        timeStamp = Date.UtcNow
        startOfCal = StartofCalendar.ToUniversalTime
    End Sub

End Class
