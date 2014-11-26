Imports xlNS = Microsoft.Office.Interop.Excel
Public Class clsPhasen

    Private AllPhasen As Collection


    Public Sub Add(phase As clsPhasenDefinition)

        AllPhasen.Add(phase, phase.name)

    End Sub

    Public Sub Remove(myitem As Object)

        AllPhasen.Remove(myitem)

    End Sub

    Public ReadOnly Property Count() As Integer

        Get
            Count = AllPhasen.Count
        End Get

    End Property

    Public ReadOnly Property Contains(name As String) As Boolean
        Get
            Contains = AllPhasen.Contains(name)
        End Get
    End Property

    Public ReadOnly Property getPhaseDef(ByVal myitem As String) As clsPhasenDefinition

        Get
            getPhaseDef = CType(AllPhasen.Item(myitem), clsPhasenDefinition)
        End Get

    End Property

    Public ReadOnly Property getPhaseDef(ByVal myitem As Integer) As clsPhasenDefinition

        Get
            getPhaseDef = CType(AllPhasen.Item(myitem), clsPhasenDefinition)
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


            If AllPhasen.Contains(name) Then

                appearanceID = Me.getPhaseDef(name).darstellungsKlasse
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

        AllPhasen = New Collection

    End Sub

End Class
