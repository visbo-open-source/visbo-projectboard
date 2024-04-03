Imports ProjectBoardDefinitions
Public Class clsKostenartDefinitionWeb
    'Inherits clsKostenartDefinitionDB

    Public name As String
    Public farbe As Long
    Public uid As Integer
    Public timestamp As Date
    ' Id wird von MongoDB automatisch gesetzt 
    Public Id As String

    Public subCostIDs As List(Of clsSubCostID)

    Public Sub copyTo(ByRef costDef As clsKostenartDefinition)
        With costDef

            If subCostIDs.Count >= 1 Then
                ' wegen Mongo müssen die Keys in String Format sein ... 

                For Each subCostItem As clsSubCostID In subCostIDs
                    Dim tmpValue As Double = 1.0
                    If IsNumeric(subCostItem.value) Then
                        tmpValue = CDbl(subCostItem.value)
                        If tmpValue >= 0 And tmpValue <= 1.0 Then
                            ' alles ok
                        Else
                            tmpValue = 1.0
                        End If
                    Else
                        tmpValue = 1.0
                    End If

                    Try
                        .addSubCost(CInt(subCostItem.key), tmpValue)
                    Catch ex As Exception

                    End Try

                Next

            End If

            .name = Me.name
            .UID = Me.uid

            '.farbe = Me.farbe
        End With
    End Sub

    Public Sub copyFrom(ByVal costDef As clsKostenartDefinition)
        With costDef

            If .getSubCostCount >= 1 Then
                For Each kvp As KeyValuePair(Of Integer, Double) In .getSubCostIDs
                    Dim tmpCost As New clsSubCostID
                    tmpCost.key = CStr(kvp.Key)
                    tmpCost.value = CStr(kvp.Value)
                    'subCostIDs.Add(CStr(kvp.Key), kvp.Value.ToString)
                    subCostIDs.Add(tmpCost)
                Next
            End If

            name = .name
            uid = .UID
            'Me.farbe = CLng(.farbe)
            Id = "Cost" & "#" & CStr(Me.uid) & "#" & Date.UtcNow.ToString
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

            Dim stillok As Boolean = True

            If subCostIDs.Count = vglCost.subCostIDs.Count Then
                If subCostIDs.Count = 0 Then
                    stillok = True
                Else
                    Dim i As Integer = 0
                    Do While i < subCostIDs.Count And stillok
                        stillok = (subCostIDs.ElementAt(i).Key = vglCost.subCostIDs.ElementAt(i).Key And
                                   subCostIDs.ElementAt(i).Value = vglCost.subCostIDs.ElementAt(i).Value)
                        i = i + 1
                    Loop

                End If
            Else
                stillok = False
            End If

            If stillok Then
                stillok = (name = vglCost.name And
                             uid = vglCost.uid)
            End If

            isIdenticalTo = stillok

        End Get
    End Property

    Public Sub New()
        subCostIDs = New List(Of clsSubCostID)
        'timestamp = Date.UtcNow
        Id = ""
    End Sub

    'Public Sub New(ByVal tmpDate As Date)
    '    'timestamp = Date.UtcNow
    '    Id = ""
    'End Sub

End Class
