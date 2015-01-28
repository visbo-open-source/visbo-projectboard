Public Class clsAwinEvents
    Private AllEvents As Collection


    Public Sub Clear()
        AllEvents.Clear()
    End Sub

    ''' <summary>
    ''' fügt der global verfügbaren Collection eine Eventvariable hinzu, so daß die eventvaiable ständig verfügbar ist
    ''' </summary>
    ''' <param name="eventvariable"></param>
    ''' <remarks> eventvariable muss in der Awin Event Klasse als Public withevents deklariert sein </remarks>
    Public Sub Add(eventvariable As Object)

        AllEvents.Add(eventvariable)

    End Sub

    ''' <summary>
    ''' löscht ein Element aus der collection von Eventvariablen
    ''' </summary>
    ''' <param name="myitem"></param>
    ''' <remarks></remarks>
    Public Sub Remove(myitem As Integer)

        AllEvents.Remove(myitem)

    End Sub

    Public ReadOnly Property Count() As Integer

        Get
            Count = AllEvents.Count
        End Get

    End Property


    Public ReadOnly Property getEvent(myitem As Integer) As Object


        Get
            getEvent = AllEvents.Item(myitem)
        End Get


    End Property




    Public Sub New()

        AllEvents = New Collection

    End Sub

End Class
