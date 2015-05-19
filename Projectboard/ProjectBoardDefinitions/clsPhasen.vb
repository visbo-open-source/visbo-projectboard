Imports xlNS = Microsoft.Office.Interop.Excel
Public Class clsPhasen

    Private AllPhasen As SortedList(Of String, clsPhasenDefinition)


    Public Sub Add(phaseDef As clsPhasenDefinition)


        If Not AllPhasen.ContainsKey(phaseDef.name) Then
            AllPhasen.Add(phaseDef.name, phaseDef)
        End If


    End Sub

    Public ReadOnly Property Count() As Integer

        Get
            Count = AllPhasen.Count
        End Get

    End Property

    Public ReadOnly Property Contains(name As String) As Boolean
        Get
            Contains = AllPhasen.ContainsKey(name)
        End Get
    End Property

    Public ReadOnly Property getPhaseDef(ByVal myitem As String) As clsPhasenDefinition

        Get
            If AllPhasen.ContainsKey(myitem) Then
                getPhaseDef = CType(AllPhasen.Item(myitem), clsPhasenDefinition)
            Else
                getPhaseDef = AllPhasen.First.Value
            End If

        End Get

    End Property

    ''' <summary>
    ''' gibt die Phasen-Definition an der Index-Position index zurüclk: Index kann von 1 .. Anzahl Phasedefs gehen 
    ''' </summary>
    ''' <param name="index"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getPhaseDef(ByVal index As Integer) As clsPhasenDefinition

        Get
            If index < 1 Then
                index = 1
            ElseIf index > AllPhasen.Count Then
                index = AllPhasen.Count
            End If
            getPhaseDef = CType(AllPhasen.ElementAt(index - 1).Value, clsPhasenDefinition)
        End Get

    End Property

    ''' <summary>
    ''' gibt die Abkürzung, den Shortname für den Meilenstein zurück
    ''' wenn er nicht gefunden wird: "n.a."
    ''' </summary>
    ''' <param name="name">Langname Meilenstein</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getAbbrev(ByVal name As String) As String
        Get
            Dim msAbbrev As String = "n.a."

            'Dim key As String = calcKey(name, belongsTo)

            If AllPhasen.ContainsKey(name) Then
                msAbbrev = CType(AllPhasen.Item(name), clsPhasenDefinition).shortName
            End If

            getAbbrev = msAbbrev

        End Get
    End Property

    ''' <summary>
    ''' gibt die Shape Definition für die angegebene Phase zurück
    ''' wenn es die Definition für name nicht gibt, wird die Default Phasen Klasse verwendet   
    ''' </summary>
    ''' <param name="name"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getShape(ByVal name As String) As xlNS.Shape
        Get
            Dim appearanceID As String
            Dim defaultPhaseAppearance As String = "Phasen Default"


            If AllPhasen.ContainsKey(name) Then

                appearanceID = CType(AllPhasen.Item(name), clsPhasenDefinition).darstellungsKlasse
                If appearanceID = "" Then
                    appearanceID = defaultPhaseAppearance
                End If

            Else

                appearanceID = defaultPhaseAppearance

            End If

            ' jetzt ist in der AppearanceID was drin ... 
            getShape = appearanceDefinitions.Item(appearanceID).form

        End Get
    End Property

    Public Sub New()

        AllPhasen = New SortedList(Of String, clsPhasenDefinition)

    End Sub
    Public Sub Clear()

        AllPhasen.Clear()

    End Sub
End Class
