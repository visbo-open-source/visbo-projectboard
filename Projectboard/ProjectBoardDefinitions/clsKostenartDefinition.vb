Public Class clsKostenartDefinition

    Public name As String
    Public farbe As Object = visboFarbeBlau

    Private _budget() As Double
    Private _uuid As Integer            ' muss eindeutig sein, da in der Liste allKostenarten danach sortiert

    Private _subCostIDs As SortedList(Of Integer, Double)

    Public ReadOnly Property getSubCostIDs As SortedList(Of Integer, Double)
        Get
            getSubCostIDs = _subCostIDs
        End Get
    End Property

    Public Sub addSubCost(ByVal subCostUid As Integer, ByVal addOn As Double)
        If Not _subCostIDs.ContainsKey(subCostUid) Then

            _subCostIDs.Add(subCostUid, addOn)

        End If
    End Sub

    'Public ReadOnly Property farbe As Integer
    '    Get
    '        farbe = visboFarbeBlau
    '    End Get
    'End Property

    Public ReadOnly Property getSubCostCount As Integer
        Get
            Dim tmpValue As Integer = 0
            If Not IsNothing(_subCostIDs) Then
                tmpValue = _subCostIDs.Count
            Else
                tmpValue = 0
            End If

            getSubCostCount = tmpValue
        End Get
    End Property

    Public ReadOnly Property isCombinedCost As Boolean
        Get
            Dim tmpValue As Boolean = False
            If IsNothing(_subCostIDs) Then
                tmpValue = False
            ElseIf _subCostIDs.Count >= 1 Then
                tmpValue = True
            Else
                tmpValue = False
            End If

            isCombinedCost = tmpValue
        End Get
    End Property

    Public Property UID() As Integer
        Get
            UID = _uuid
        End Get
        Set(value As Integer)
            _uuid = value
        End Set
    End Property

    Public ReadOnly Property hasAnyOfThemAsChild(ByVal tmpCollection As Collection) As Boolean
        Get
            Dim tmpCheck As Boolean = False
            Dim myRoleName As String = Me.name

            For Each kvp As KeyValuePair(Of Integer, Double) In getSubCostIDs
                Dim tmpName As String = CostDefinitions.getCostDefByID(kvp.Key).name
                If tmpCollection.Contains(tmpName) Then
                    tmpCheck = True
                Else
                    ' 
                    If CostDefinitions.containsUid(kvp.Key) Then
                        Dim tmpCostDef As clsKostenartDefinition = CostDefinitions.getCostDefByID(kvp.Key)
                        If tmpCostDef.isCombinedCost Then
                            tmpCheck = tmpCostDef.hasAnyOfThemAsChild(tmpCollection)
                        End If
                    End If

                End If

                If tmpCheck = True Then
                    Exit For
                End If

            Next

            hasAnyOfThemAsChild = tmpCheck
        End Get
    End Property


    ''' <summary>
    ''' true, if both costdefinitions are identical 
    ''' </summary>
    ''' <param name="vglCost"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property isIdenticalTo(ByVal vglCost As clsKostenartDefinition) As Boolean
        Get
            Dim stillok As Boolean = True

            If _subCostIDs.Count = vglCost.getSubCostIDs.Count Then
                If _subCostIDs.Count = 0 Then
                    stillok = True
                Else

                    Dim i As Integer = 0
                    Do While i < _subCostIDs.Count And stillok
                        stillok = (_subCostIDs.ElementAt(i).Key = vglCost.getSubCostIDs.ElementAt(i).Key And
                                   _subCostIDs.ElementAt(i).Value = vglCost.getSubCostIDs.ElementAt(i).Value)
                        i = i + 1
                    Loop

                End If
            Else
                stillok = False
            End If

            If stillok Then
                stillok = (name = vglCost.name And
                             CLng(farbe) = CLng(vglCost.farbe) And
                             UID = vglCost.UID)

            End If

            isIdenticalTo = stillok

        End Get
    End Property

    Public Sub New()

    End Sub
End Class
