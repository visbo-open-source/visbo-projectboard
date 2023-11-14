Imports ProjectBoardDefinitions
Public Class frmProjectEditSettings
    Private topPos As Double = Me.Top
    Private leftPos As Double = Me.Left

    Private Sub AdjustResourceNeeds_CheckedChanged(sender As Object, e As EventArgs) Handles AdjustResourceNeeds.CheckedChanged
        'awinSettings.propAnpassRess = AdjustResourceNeeds.Checked
    End Sub

    Private Sub newCalculation_CheckedChanged(sender As Object, e As EventArgs) Handles newCalculation.CheckedChanged
        'awinSettings.noNewCalculation = Not newCalculation.Checked
    End Sub

    Private Sub invoices_CheckedChanged(sender As Object, e As EventArgs) Handles invoices.CheckedChanged
        'awinSettings.enableInvoices = invoices.Checked
    End Sub

    Private Sub adjustChilds_CheckedChanged(sender As Object, e As EventArgs) Handles adjustChilds.CheckedChanged
        ' awinSettings.autoAjustChilds = adjustChilds.Checked
    End Sub

    Private Sub frmProjectEditSettings_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        If awinSettings.englishLanguage Then
            AdjustResourceNeeds.Text = "Adjust resource needs proportionally (when duration of phase has changed)"
            newCalculation.Text = "Distribute values automatically over time  (when duration or start/end of phase has changed)"
            invoices.Text = "Edit Revenue/Benefit and contractual penalties"
            adjustChilds.Text = "Date changes also affect the subordinate tasks (only used for Time)"
            allowOverUtilization.Text = "overloading resources is allowed"

        Else
            AdjustResourceNeeds.Text = "Ressourcen Bedarfe proportional anpassen (bei Änderung der Dauer der Phase) "
            newCalculation.Text = "Werte autom. über die Zeit verteilen  (bei Änderung Dauer oder Start/Ende der Phase) "
            invoices.Text = "Umsätze/Benefits und Vertrags-Strafen bearbeiten"
            adjustChilds.Text = "Start- und Endedatum der 'Kinder' automatisch anpassen (nur für Zeitangaben verwendet)"
            allowOverUtilization.Text = "Überlastung erlauben"

        End If

        AdjustResourceNeeds.Checked = awinSettings.propAnpassRess
        newCalculation.Checked = Not awinSettings.noNewCalculation
        invoices.Checked = awinSettings.enableInvoices
        adjustChilds.Checked = awinSettings.autoAjustChilds
        ' tk added 23.11.22: to be able to allow over-Utilization
        allowOverUtilization.Checked = awinSettings.meAllowOverTime

        If currentProjektTafelModus <> ptModus.massEditTermine Then
            AdjustResourceNeeds.Enabled = False
            newCalculation.Enabled = False
            invoices.Enabled = False
            adjustChilds.Enabled = False
            allowOverUtilization.Enabled = True
        Else
            AdjustResourceNeeds.Enabled = True
            newCalculation.Enabled = True
            invoices.Enabled = True
            adjustChilds.Enabled = True
            allowOverUtilization.Enabled = False
        End If
    End Sub

    Private Sub frmProjectEditSettings_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        awinSettings.propAnpassRess = AdjustResourceNeeds.Checked
        awinSettings.noNewCalculation = Not newCalculation.Checked
        awinSettings.enableInvoices = invoices.Checked
        awinSettings.autoAjustChilds = adjustChilds.Checked
        ' tk added 23.11.22 to be able to allow more days when capacity is available
        awinSettings.meAllowOverTime = allowOverUtilization.Checked

        topPos = Me.Top
        leftPos = Me.Left

    End Sub

    Private Sub allowOverUtilization_CheckedChanged(sender As Object, e As EventArgs) Handles allowOverUtilization.CheckedChanged

    End Sub
End Class