''' <summary>
''' speziell für BMW Improt erstellte Klasse
''' </summary>
''' <remarks></remarks>
Public Class clsImportFileHierarchy

    Private phaseHierarchy As SortedList(Of Integer, clsPhase)

    ' ''' <summary>
    ' ''' liefert true zurück, wenn der übergebene Name bereits so in der Phasen Hierarchie vorhanden war 
    ' ''' </summary>
    ' ''' <param name="phaseName"></param>
    ' ''' <value></value>
    ' ''' <returns></returns>
    ' ''' <remarks></remarks>
    'Public ReadOnly Property dopplung(ByVal phaseName As String) As Boolean
    '    Get
    '        Dim ix As Integer = phaseHierarchy.Count
    '        Dim found As Boolean = False
    '        Dim tmpstr(30) As String
    '        Dim vgl1 As String = ""
    '        Dim vgl2 As String = ""


    '        While Not found And ix > 0

    '            tmpstr = phaseHierarchy.ElementAt(ix - 1).Value.name.Trim.Split(New Char() {CChar(" ")}, 30)
    '            For i = 1 To tmpstr.Length
    '                If i = 1 Then
    '                    vgl1 = tmpstr(i - 1)
    '                Else
    '                    vgl1 = vgl1 & tmpstr(i - 1)
    '                End If
    '            Next

    '            tmpstr = phaseName.Trim.Split(New Char() {CChar(" ")}, 30)
    '            For i = 1 To tmpstr.Length
    '                If i = 1 Then
    '                    vgl2 = tmpstr(i - 1)
    '                Else
    '                    vgl2 = vgl2 & tmpstr(i - 1)
    '                End If
    '            Next

    '            If vgl1.Trim = vgl2.Trim Then
    '                found = True
    '            Else
    '                ix = ix - 1
    '            End If
    '        End While

    '        dopplung = found

    '    End Get
    'End Property

    ''' <summary>
    ''' normiert den Namen phaseName, d.h entfernt alle leading und trailing blanks 
    ''' stellt sicher, dass Worte nur über ein Blank getrennt sind 
    ''' </summary>
    ''' <param name="phaseName"></param>
    ''' <remarks></remarks>
    ''' 
    Public Function normierung(ByVal phaseName As String) As String

        Dim ergebnis As String = Nothing
        Dim tmpstr(20) As String

        If Not IsNothing(phaseName) Then

            ergebnis = phaseName.Trim

            If ergebnis.Length > 0 Then
                tmpstr = ergebnis.Split(New Char() {CChar(" ")}, 19)
                If tmpstr.Length > 1 Then
                    ergebnis = ""
                    For i = 1 To tmpstr.Length
                        If i = 1 Then
                            ergebnis = tmpstr(i - 1)
                        Else
                            ergebnis = ergebnis & " " & tmpstr(i - 1)
                        End If
                    Next
                End If
                
            Else
                ergebnis = Nothing
            End If

        End If

        normierung = ergebnis

    End Function

    ''' <summary>
    ''' gibt die Zahl der leading blanks von phaseName zurück
    ''' wird der Index für die sortierte Liste Hierarchy 
    ''' </summary>
    ''' <param name="phaseName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getLevel(ByVal phaseName As String) As Integer

        Dim leadingBlanks As Integer = 0

        If Not IsNothing(phaseName) Then

            leadingBlanks = phaseName.TrimEnd.Length - phaseName.Trim.Length

        End If

        getLevel = leadingBlanks

    End Function

    ''' <summary>
    ''' gibt den Namen der aktuellen Ebene mit Nummer ebenenNr zurück 
    ''' kann den Wert 0 .. count-1 haben  
    ''' </summary>
    ''' <param name="ebenenNr"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getEbenenName(ByVal ebenenNr As Integer) As String

        If ebenenNr >= 0 And ebenenNr < phaseHierarchy.Count Then
            getEbenenName = phaseHierarchy.ElementAt(ebenenNr).Value.name
        Else
            getEbenenName = ""
        End If

    End Function

    ''' <summary>
    ''' ''' gibt den aktuellen Footprint bis zur Ebene mit IndentLevel kleiner level zurück 
    ''' der Hierarchie zurück 
    ''' </summary>
    ''' <param name="level">Indentlevel</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getFootPrint(ByVal level As Integer) As String

        Dim tmpValue As String = ""

        If phaseHierarchy.Count > 0 Then

            Dim index As Integer = 0
            Dim found As Boolean = False

            While index <= phaseHierarchy.Count - 1 And Not found
                If phaseHierarchy.ElementAt(index).Key < level Then
                    index = index + 1
                Else
                    found = True
                End If
            End While

            If index = 0 Then
                tmpValue = ""
            Else
                For i As Integer = 0 To index - 1
                    If i = 0 Then
                        'tmpValue = phaseHierarchy.ElementAt(i).Value.name.Trim
                        tmpValue = "."
                    Else
                        tmpValue = tmpValue & " - " & phaseHierarchy.ElementAt(i).Value.name.Trim
                    End If
                Next
            End If

        End If

        getFootPrint = tmpValue

    End Function


    ''' <summary>
    ''' nName ist bereits normiert sein. d.h alle leading, trailing und Anzahl>1 Blanks in der Mitte wurden entfernt
    ''' level ist die Anzahl an leading blanks 
    ''' Nach dem Eintrag wird die Hierarchie ggf wieder so weit zurückgesetzt, wie erforderlich
    ''' </summary>
    ''' <param name="phase">Phase</param>
    ''' <param name="level"></param>
    ''' <remarks></remarks>
    Public Sub add(ByVal phase As clsPhase, ByVal level As Integer)


        If Not IsNothing(phase) And level >= 0 Then

            If phaseHierarchy.ContainsKey(level) Then
                phaseHierarchy.Item(level) = phase
            Else
                phaseHierarchy.Add(level, phase)
            End If

            Do While phaseHierarchy.Last.Key > level

                phaseHierarchy.Remove(phaseHierarchy.Last.Key)

            Loop

        End If


    End Sub

    ''' <summary>
    ''' gibt die Phase zurück, die die letzte der angegebenen Ebene war
    ''' wenn die nicht exakt existiert, die Phase der nächste tieferen Ebene 
    ''' </summary>
    ''' <param name="level"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getPhaseBeforeLevel(ByVal level As Integer) As clsPhase
        Get

            Dim ix As Integer = phaseHierarchy.Count - 1

            If phaseHierarchy.Count = 0 Then
                getPhaseBeforeLevel = Nothing

            Else
                Do While phaseHierarchy.ElementAt(ix).Key >= level And ix > 0
                    ix = ix - 1
                Loop

                getPhaseBeforeLevel = phaseHierarchy.ElementAt(ix).Value

            End If


        End Get
    End Property

    ''' <summary>
    ''' gibt den Indent der letzten Hierarchie Stufe zurück 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getCurrentLevel As Integer
        Get
            If phaseHierarchy.Count > 0 Then
                getCurrentLevel = phaseHierarchy.Last.Key
            Else
                getCurrentLevel = -1
            End If

        End Get
    End Property

    ''' <summary>
    ''' gibt die Anzahl Einträge in der sortierten Liste zurück 
    ''' entspricht der Anzahl Hierarchie-Ebenen  
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property count() As Integer
        Get
            count = phaseHierarchy.Count
        End Get
    End Property


    Sub New()
        phaseHierarchy = New SortedList(Of Integer, clsPhase)
    End Sub



End Class
