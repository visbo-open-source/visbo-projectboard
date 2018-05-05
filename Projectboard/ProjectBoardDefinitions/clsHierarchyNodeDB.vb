''' <summary>
''' Klassen-Definition für Hierarchie-Knoten
''' </summary>
''' <remarks></remarks>
Public Class clsHierarchyNodeDB
    Public elemName As String
    Public origName As String
    Public indexOfElem As Integer
    Public parentNodeKey As String
    Public childNodeKeys As List(Of String)

    ' 
    ''' <summary>
    ''' kopiert einen HAuptspeicher Hierarchie Knoten in einen DB Hierarchie Knoten 
    ''' </summary>
    ''' <param name="hryNode"></param>
    ''' <remarks></remarks>
    Sub copyFrom(ByVal hryNode As clsHierarchyNode)

        Dim childID As String
        With hryNode
            Me.elemName = .elemName
            ' ist seit 29.5 niht mehr Bestandteil eines Hierarchie Knotens
            'Me.origName = .origName
            Me.indexOfElem = .indexOfElem
            Me.parentNodeKey = .parentNodeKey
            For i As Integer = 1 To .childCount
                childID = .getChild(i)
                Me.childNodeKeys.Add(childID)
            Next
        End With

    End Sub

    ''' <summary>
    ''' kopiert einen DB Hierarchie-Knoten in einen Hauptspeicher Hierarchie Knoten 
    ''' </summary>
    ''' <param name="hryNode"></param>
    ''' <remarks></remarks>
    Sub copyTo(ByRef hryNode As clsHierarchyNode)

        Dim childID As String
        With hryNode
            .elemName = Me.elemName
            ' ist seit 29.5 nicht mehr Bestandteil eines Hierarchie-Knotens 
            '.origName = Me.origName
            .indexOfElem = Me.indexOfElem
            .parentNodeKey = Me.parentNodeKey
            For i As Integer = 1 To Me.childNodeKeys.Count
                childID = Me.childNodeKeys.Item(i - 1)
                .addChild(childID)
            Next
        End With

    End Sub

    Sub New()

        childNodeKeys = New List(Of String)

    End Sub

End Class
