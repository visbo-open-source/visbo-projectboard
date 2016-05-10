Imports xlNS = Microsoft.Office.Interop.Excel
Public Class clsPhasen

    Private AllPhasen As SortedList(Of String, clsPhasenDefinition)


    ''' <summary>
    ''' nimmt die Phase auf; wenn der Name bereits vergeben ist, wird nichts gemacht ...
    ''' wenn PhaseDef = Nothing, wird auch nichts gemacht 
    ''' es werden keine Exceptions geworfen; wenn man an der Aufruf Stelle wissen muss, ob der Name vergeben ist, muss über .contains geprüft werden 
    ''' </summary>
    ''' <param name="phaseDef"></param>
    ''' <remarks></remarks>
    Public Sub Add(phaseDef As clsPhasenDefinition)

        If Not IsNothing(phaseDef) Then
            If Not AllPhasen.ContainsKey(phaseDef.name) Then
                AllPhasen.Add(phaseDef.name, phaseDef)
            Else
                ' nichts tun , ist ja schon da 
            End If
        Else
            ' nichts tun , es ist ja nichts aufzunehmen  
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
                'getPhaseDef = AllPhasen.First.Value
                getPhaseDef = Nothing
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
    ''' wenn er nicht gefunden wird: 
    ''' </summary>
    ''' <param name="name">Langname Phase</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getAbbrev(ByVal name As String) As String
        Get
            Dim msAbbrev As String = name

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

            ' ''Dim ok As Boolean = False
            ' ''While Not ok
            ' ''    Try
            ' ''        ' jetzt ist in der AppearanceID was drin ... 
            ' ''        getShape = appearanceDefinitions.Item(appearanceID).form
            ' ''        If Not IsNothing(getShape) Then
            ' ''            ok = True
            ' ''        Else
            ' ''            Call MsgBox("nothing")
            ' ''        End If
            ' ''    Catch ex As Exception
            ' ''        Call MsgBox("getshape fehlerhaft")
            ' ''        getShape = Nothing
            ' ''    End Try

            ' ''End While

            ' jetzt ist in der AppearanceID was drin ... 
            getShape = appearanceDefinitions.Item(appearanceID).form

        End Get
    End Property

    ''' <summary>
    ''' löscht die Phasen-Definition mit dem übergebenen Namen aus der Liste , sofern vorhanden
    ''' wenn nicht vorhanden, keine Änderung; aber auch keine Mitteilung 
    ''' </summary>
    ''' <param name="name"></param>
    ''' <remarks></remarks>
    Public Sub remove(ByVal name As String)

        If AllPhasen.ContainsKey(name) Then
            AllPhasen.Remove(name)
        End If

    End Sub
    
    ''' <summary>
    ''' leert die komplette Liste 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Clear()

        AllPhasen.Clear()

    End Sub


    Public Sub New()

        AllPhasen = New SortedList(Of String, clsPhasenDefinition)


    End Sub
End Class
