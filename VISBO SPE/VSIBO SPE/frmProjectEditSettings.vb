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
            AdjustResourceNeeds.Text = "Adjust resource needs proportionally (only used for Time)"
            newCalculation.Text = "Distribute values automatically over time  (only used for Time)"
            invoices.Text = "Process invoices and contractual penalties"
            adjustChilds.Text = "Date changes also affect the subordinate tasks (only used for Time)"

        Else
            AdjustResourceNeeds.Text = "Ressourcen Bedarfe proportional anpassen (nur für Zeitangaben verwendet) "
            newCalculation.Text = "Werte autom. über die Zeit verteilen  (nur für Zeitangaben verwendet) "
            invoices.Text = "Rechnungen und Vertrags-Strafen bearbeiten"
            adjustChilds.Text = "Start- und Endedatum der 'Kinder' automatisch anpassen (nur für Zeitangaben verwendet)"

        End If

        AdjustResourceNeeds.Checked = awinSettings.propAnpassRess
        newCalculation.Checked = Not awinSettings.noNewCalculation
        invoices.Checked = awinSettings.enableInvoices
        adjustChilds.Checked = awinSettings.autoAjustChilds

        If currentProjektTafelModus <> ptModus.massEditTermine Then
            AdjustResourceNeeds.Enabled = False
            newCalculation.Enabled = False
            invoices.Enabled = False
            adjustChilds.Enabled = False
        Else
            AdjustResourceNeeds.Enabled = True
            newCalculation.Enabled = True
            invoices.Enabled = True
            adjustChilds.Enabled = True
        End If
    End Sub

    Private Sub frmProjectEditSettings_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        awinSettings.propAnpassRess = AdjustResourceNeeds.Checked
        awinSettings.noNewCalculation = Not newCalculation.Checked
        awinSettings.enableInvoices = invoices.Checked
        awinSettings.autoAjustChilds = adjustChilds.Checked

        topPos = Me.Top
        leftPos = Me.Left

    End Sub
End Class