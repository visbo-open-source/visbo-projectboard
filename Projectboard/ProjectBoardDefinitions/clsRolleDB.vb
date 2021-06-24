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
    Public isCalculated As Boolean

    Sub copyFrom(ByVal role As clsRolle)

        With role

            ' now in case there is true: onePersonHasOneRole
            ' added 21.6.21
            Try
                If awinSettings.onePersonOneRole Then
                    ' substitute team by the one and only one team / role that person has 
                    ' is relevant for instart and all other customers working with Junior, Expert, Senior Roles
                    Dim myRole As clsRollenDefinition = RoleDefinitions.getRoleDefByID(.uid)
                    If Not IsNothing(myRole) Then
                        If Not myRole.isCombinedRole Then
                            Dim mySkills As SortedList(Of Integer, Double) = myRole.getSkillIDs
                            If mySkills.Count = 1 Then
                                .teamID = mySkills.First.Key
                            End If
                        End If
                    End If

                End If

            Catch ex As Exception

            End Try


            Me.RollenTyp = .uid
            Me.Bedarf = .Xwerte
            Me.teamID = .teamID

            ' 24.11.18 deprecated
            Me.name = Nothing
            Me.farbe = Nothing
            Me.startkapa = Nothing
            Me.isCalculated = Nothing

        End With

    End Sub

    Sub copyto(ByRef role As clsRolle)

        With role
            .uid = Me.RollenTyp
            .Xwerte = Me.Bedarf
            .teamID = Me.teamID
            ' now in case there is true: onePersonHasOneRole
            Try
                If awinSettings.onePersonOneRole Then
                    ' substitute team by the one and only one team / role that person has 
                    ' is relevant for instart and all other customers working with Junior, Expert, Senior Roles
                    Dim myRole As clsRollenDefinition = RoleDefinitions.getRoleDefByID(.uid)
                    If Not IsNothing(myRole) Then
                        If Not myRole.isCombinedRole Then
                            Dim mySkills As SortedList(Of Integer, Double) = myRole.getSkillIDs
                            If mySkills.Count = 1 Then
                                .teamID = mySkills.First.Key
                            End If
                        End If
                    End If

                End If

            Catch ex As Exception

            End Try

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
