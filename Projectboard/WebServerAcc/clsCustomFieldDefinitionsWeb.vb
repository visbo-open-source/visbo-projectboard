Imports ProjectBoardDefinitions
Public Class clsCustomFieldDefinitionsWeb

    Private listOfDefinitions As List(Of clsCustomFieldDefinition)


    ''' <summary>
    ''' gibt die sortierte Liste der Custom Fields zurück
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property liste As List(Of clsCustomFieldDefinition)
        Get
            liste = listOfDefinitions
        End Get
    End Property
    Public Sub copyFrom(ByVal custFieldDef As clsCustomFieldDefinitions)

        With custFieldDef

            For Each kvp As KeyValuePair(Of Integer, clsCustomFieldDefinition) In .liste
                Dim cfd As New clsCustomFieldDefinition
                cfd = kvp.Value
                Me.listOfDefinitions.Add(cfd)
            Next

        End With
    End Sub
    Public Sub copyTo(ByVal custFieldDef As clsCustomFieldDefinitions)

        With custFieldDef

            For Each cfdWeb As clsCustomFieldDefinition In Me.liste

                .liste.Add(cfdWeb.uid, cfdWeb)

            Next

        End With
    End Sub
    Sub New()
            listOfDefinitions = New List(Of clsCustomFieldDefinition)
        End Sub

End Class
