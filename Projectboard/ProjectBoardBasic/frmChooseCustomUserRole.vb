Imports ProjectBoardDefinitions
Public Class frmChooseCustomUserRole

    Public selectedIndex As Integer = -1
    Public myUserRoles As Collection

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnOK.Click
        If Not IsNothing(dgv_customUserRoles.SelectedRows) Then
            If dgv_customUserRoles.SelectedRows.Count > 0 Then
                selectedIndex = CInt(dgv_customUserRoles.SelectedRows.Item(0).Tag)
            End If
        End If
    End Sub

    Private Sub frmChooseCustomUserRole_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Call listeAufbauen()

    End Sub

    Private Sub listeAufbauen()


        If Not IsNothing(myUserRoles) Then

            Dim anzChangeItems As Integer = myUserRoles.Count
            If anzChangeItems > 0 Then

                dgv_customUserRoles.Rows.Add(anzChangeItems)

                For i As Integer = 1 To anzChangeItems

                    Dim currentItem As clsCustomUserRole = myUserRoles.Item(i)

                    With currentItem
                        dgv_customUserRoles.Rows(i - 1).Tag = (i).ToString
                        dgv_customUserRoles.Rows(i - 1).Cells(0).Value = .customUserRole.ToString
                        If .customUserRole = ptCustomUserRoles.PortfolioManager Then
                            dgv_customUserRoles.Rows(i - 1).Cells(1).Value = ""
                        ElseIf .customUserRole = ptCustomUserRoles.RessourceManager Then
                            Dim tmpteamID As Integer = -1
                            dgv_customUserRoles.Rows(i - 1).Cells(1).Value = RoleDefinitions.getRoleDefByIDKennung(.specifics, tmpteamID).name
                        Else
                            dgv_customUserRoles.Rows(i - 1).Cells(1).Value = ""
                        End If

                    End With

                    'changeListTable.Rows(i).Tag = changeliste.getShapeNameFromChangeList(i + 1)

                Next
            End If

            If anzChangeItems = 1 Then
                ' es muss eine zusätzliche Zeile hinzugefügt werden, sonst ist diese eine Zeile nicht zu selektieren 
                dgv_customUserRoles.Rows.Add(1)
                dgv_customUserRoles.Rows.Item(0).Selected = False
                dgv_customUserRoles.Rows.Item(1).Selected = True
            End If

        End If
    End Sub

End Class