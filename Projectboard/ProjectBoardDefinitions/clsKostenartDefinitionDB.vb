Public Class clsKostenartDefinitionDB

    Public name As String
    Public farbe As Long
    Public uid As Integer
    Public timestamp As Date
    ' Id wird von MongoDB automatisch gesetzt 
    Public Id As String

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
            Me.Id = "Cost" & "#" & CStr(Me.uid) & "#" & Date.UtcNow.ToString
        End With
    End Sub

    ''' <summary>
    ''' true, if both costdefinitions are identical , except timestamp 
    ''' </summary>
    ''' <param name="vglCost"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property isIdenticalTo(ByVal vglCost As clsKostenartDefinitionDB) As Boolean
        Get

            isIdenticalTo = (Me.name = vglCost.name And _
                             Me.farbe = vglCost.farbe And _
                             Me.uid = vglCost.uid)

        End Get
    End Property

    Public Sub New()
        timestamp = Date.UtcNow
        Id = ""
    End Sub

    Public Sub New(ByVal tmpDate As Date)
        timestamp = Date.UtcNow
        Id = ""
    End Sub
End Class
