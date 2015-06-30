Imports ProjectBoardDefinitions
Imports Microsoft.Office.Interop.Excel
Public Class frmAnnotateProject

    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click

        Dim worksheetShapes As Excel.Shapes
        Dim projectshape As Excel.Shape

        Try

            worksheetShapes = CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet).Shapes

        Catch ex As Exception
            Call MsgBox("keine Shapes Zuordnung möglich ")
            Exit Sub
        End Try

        Me.Cursor = Cursors.WaitCursor
        If selectedProjekte.Count = 0 Then
            For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste
                Try
                    projectshape = worksheetShapes.Item(kvp.Value.name)
                    Call annotateProject(projectshape, annotatePhases.Checked, annotateMilestones.Checked, _
                                         showStdNames.Checked, showAbbrev.Checked)
                Catch ex As Exception

                End Try
            Next
        Else
            OKButton.Text = "Beschriften"

            For Each kvp As KeyValuePair(Of String, clsProjekt) In selectedProjekte.Liste

                Try
                    projectshape = worksheetShapes.Item(kvp.Value.name)
                    Call annotateProject(projectshape, annotatePhases.Checked, annotateMilestones.Checked, _
                                         showStdNames.Checked, showAbbrev.Checked)
                Catch ex As Exception

                End Try


                'Try
                '    worksheetShapes = CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet).Shapes
                '    projectshape = worksheetShapes.Item(kvp.Value.name)
                'Catch ex As Exception

                'End Try
            Next
        End If
        Me.Cursor = Cursors.Default
        

    End Sub

    Private Sub showAbbrev_CheckedChanged(sender As Object, e As EventArgs) Handles showAbbrev.CheckedChanged

        If showAbbrev.Checked = True Then
            showStdNames.Checked = True
        End If

    End Sub

    
    Private Sub showOrigNames_CheckedChanged(sender As Object, e As EventArgs) Handles showOrigNames.CheckedChanged
        If showOrigNames.Checked = True Then
            showAbbrev.Checked = False
        End If
    End Sub

    Private Sub frmAnnotateProject_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If selectedProjekte.Count = 0 Then
            OKButton.Text = "Alle beschriften"
        End If
    End Sub
End Class