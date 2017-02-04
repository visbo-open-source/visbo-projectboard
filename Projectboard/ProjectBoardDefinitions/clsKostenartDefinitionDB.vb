Public Class clsKostenartDefinitionDB

    Public name As String
    Public farbe As Long
    Public uid As Integer
    Public timeStamp As Date

    Public Sub copyTo(ByRef costDef As clsKostenartDefinition)
        With costDef
            .name = Me.name
            .UID = Me.uid
            .farbe = Me.farbe
        End With
    End Sub

    Public Sub copyFrom(ByVal costDef As clsKostenartDefinition)
        With costDef
            Me.name = .name
            Me.uid = .UID
            Me.farbe = CLng(.farbe)
        End With
    End Sub

    ''' <summary>
    ''' true, if both costdefinitions are identical , except timestamp 
    ''' </summary>
    ''' <param name="vglCost"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property isIdenticalTo(ByVal vglCost As clsKostenartDefinitionDB)
        Get

            isIdenticalTo = (Me.name = vglCost.name And _
                             Me.farbe = vglCost.farbe And _
                             Me.uid = vglCost.uid)

        End Get
    End Property

    Public Sub New()
        timeStamp = Date.UtcNow
    End Sub

    Public Sub New(ByVal tmpDate As Date)
        timeStamp = Date.UtcNow
    End Sub
End Class
