Public Class clsKostenartDefinition

    Public name As String
    Public farbe As Object

    Private _budget() As Double
    Private _uuid As Integer            ' muss eindeutig sein, da in der Liste allKostenarten danach sortiert

    Public Property UID() As Integer
        Get
            UID = _uuid
        End Get
        Set(value As Integer)
            _uuid = value
        End Set
    End Property

    ''' <summary>
    ''' true, if both costdefinitions are identical , except timestamp 
    ''' </summary>
    ''' <param name="vglCost"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property isIdenticalTo(ByVal vglCost As clsKostenartDefinition) As Boolean
        Get

            isIdenticalTo = (Me.name = vglCost.name And _
                             CLng(Me.farbe) = CLng(vglCost.farbe) And _
                             Me.UID = vglCost.UID)


        End Get
    End Property

    Public Sub New()

    End Sub
End Class
