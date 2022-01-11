Imports System.Windows.Forms
Imports ProjectBoardDefinitions
Public Class frmMeRcEinstellungen
    Private Sub frmMeRcEinstellungen_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Call getFrmPosition(PTfrm.other, Top, Left)
        Call languageSettings()

        If awinSettings.meDateForLastPlan > StartofCalendar Then
            VersionDatePicker.Value = awinSettings.meDateForLastPlan
        Else
            VersionDatePicker.Value = Date.Now.AddMonths(-1)
        End If

        If awinSettings.meCompareVsLastPlan Then
            VersionDatePicker.Enabled = True
        Else
            VersionDatePicker.Enabled = False
        End If

        If visboZustaende.projectBoardMode = ptModus.massEditRessSkills Then
            chkbx_allowOvertime.Visible = True
            chkbx_compareProjectTotals.Visible = True
            chkbx_allowOvertime.Checked = awinSettings.meAllowOverTime
        Else
            chkbx_allowOvertime.Visible = False
            chkbx_allowOvertime.Enabled = False
            chkbx_compareProjectTotals.Visible = False
        End If

        chkbx_showHeader.Checked = appInstance.ActiveWindow.DisplayHeadings

        chkbx_compareWithVersion.Checked = awinSettings.meCompareVsLastPlan
        chkbx_compareProjectTotals.Checked = awinSettings.considerProjectTotals
        chkbx_AutoDistribution.Checked = Not awinSettings.noNewCalculation

    End Sub
    Private Sub languageSettings()

        If awinSettings.englishLanguage Then
            ' auf Englisch darstellen 
            Me.Text = "Edit Resources / Costs Settings"
            chkbx_showHeader.Text = "show Header"
            chkbx_compareWithVersion.Text = "compare with version from"
            chkbx_compareProjectTotals.Text = "compare project totals"
            chkbx_allowOvertime.Text = "allow Overtime"
            chkbx_AutoDistribution.Text = "auto distribution of values"
            cancel_btn.Text = "Cancel"

        End If

    End Sub

    Private Sub frmMeRcEinstellungen_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed

        Try
            frmCoord(PTfrm.other, PTpinfo.top) = Me.Top
            frmCoord(PTfrm.other, PTpinfo.left) = Me.Left
        Catch ex As Exception

        End Try

    End Sub

    Private Sub ok_Btn_Click(sender As Object, e As EventArgs) Handles ok_Btn.Click

        Try
            If chkbx_showHeader.Checked Then
                With appInstance.ActiveWindow
                    .DisplayHeadings = True
                End With
            Else
                With appInstance.ActiveWindow
                    .DisplayHeadings = False
                End With
            End If
        Catch ex As Exception

        End Try

        Try
            awinSettings.meAllowOverTime = chkbx_allowOvertime.Checked

            awinSettings.meCompareVsLastPlan = chkbx_compareWithVersion.Checked
            awinSettings.considerProjectTotals = chkbx_compareProjectTotals.Checked
            awinSettings.meDateForLastPlan = VersionDatePicker.Value

            awinSettings.noNewCalculation = Not chkbx_AutoDistribution.Checked
        Catch ex As Exception

        End Try


        MyBase.Close()

    End Sub

    Private Sub chkbx_compareWithVersion_CheckedChanged(sender As Object, e As EventArgs) Handles chkbx_compareWithVersion.CheckedChanged

        VersionDatePicker.Enabled = chkbx_compareWithVersion.Checked

    End Sub
End Class