''' <summary>
''' Klassen-Definition für Rolle mit Ressourcenbedarf 
''' </summary>
''' <remarks></remarks>
Public Class clsRolleDB
    ' tk 24.11.18 Rollentyp ist die RollenID
    Public RollenTyp As Integer
    Public Bedarf() As Double
    ' neu hinzugekommen 
    Public teamID As Integer

    ' deprecated 24.11.18 , immer mit Nothing / Null lesen/schreiben
    Public name As String
    Public farbe As Object
    Public startkapa As Integer
    Public tagessatzIntern As Double
    Public tagessatzExtern As Double
    Public isCalculated As Boolean

    Sub copyFrom(ByVal role As clsRolle)

        With role
            Me.RollenTyp = .RollenTyp
            Me.Bedarf = .Xwerte
            Me.teamID = .teamID

            ' 24.11.18 deprecated
            Me.name = Nothing
            Me.farbe = Nothing
            Me.startkapa = Nothing
            Me.tagessatzIntern = Nothing
            Me.tagessatzExtern = Nothing
            Me.isCalculated = Nothing

        End With

    End Sub

    Sub copyto(ByRef role As clsRolle)

        With role
            .RollenTyp = Me.RollenTyp
            .Xwerte = Me.Bedarf
            .teamID = Me.teamID
            '.isCalculated = Me.isCalculated
        End With

    End Sub

    Sub New()
        isCalculated = False
    End Sub

    Sub New(ByVal laenge As Integer)

        ReDim Bedarf(laenge)
        isCalculated = False

    End Sub

End Class
