Imports ProjectBoardDefinitions
Public Class clsWebVPfItem
    Inherits clsConstellationDB.clsConstellationItemDB

    Public Property name As String
    Public Property vpid As String
    Public Property _id As String


    Sub New()
        _vpid = ""
        _id = ""
        _name = ""
    End Sub


    Overloads Sub copyfrom(ByVal item As clsWebVPfItem)

        With item
            Me.projectName = .name
            Me.variantName = .variantName
            Me.Start = .Start.ToUniversalTime
            Me.show = .show
            Me.zeile = .zeile
            Me.reasonToInclude = .reasonToInclude
            Me.reasonToExclude = .reasonToExclude
        End With
    End Sub

    Overloads Sub copyto(ByRef item As clsConstellationItem)

        With item
            .projectName = Me.name
            .variantName = Me.variantName
            .start = Me.Start.ToLocalTime
            .show = Me.show
            .zeile = Me.zeile
            .reasonToInclude = Me.reasonToInclude
            .reasonToExclude = Me.reasonToExclude
        End With

    End Sub

End Class
