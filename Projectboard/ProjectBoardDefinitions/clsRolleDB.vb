''' <summary>
''' Klassen-Definition für Rolle mit Ressourcenbedarf 
''' </summary>
''' <remarks></remarks>
Public Class clsRolleDB

    Public RollenTyp As Integer
    Public name As String
    Public farbe As Object
    Public startkapa As Integer
    Public tagessatzIntern As Double
    Public tagessatzExtern As Double
    Public Bedarf() As Double
    Public isCalculated As Boolean

    Sub copyFrom(ByVal role As clsRolle)

        With role
            Me.RollenTyp = .RollenTyp
            Me.name = .name
            Me.farbe = .farbe
            Me.startkapa = CInt(.Startkapa)
            Me.tagessatzIntern = .tagessatzIntern
            Me.tagessatzExtern = .tagessatzExtern
            Bedarf = .Xwerte
            Me.isCalculated = .isCalculated
        End With

    End Sub

    Sub copyto(ByRef role As clsRolle)

        With role
            .RollenTyp = Me.RollenTyp
            .Xwerte = Me.Bedarf
            .isCalculated = Me.isCalculated
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
