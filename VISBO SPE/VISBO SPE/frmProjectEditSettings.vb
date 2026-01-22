Imports ProjectBoardDefinitions
Public Class frmProjectEditSettings
    Private topPos As Double = Me.Top
    Private leftPos As Double = Me.Left


    Private Sub frmProjectEditSettings_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        If awinSettings.englishLanguage Then
            AdjustResourceNeeds.Text = "Adjust resource needs proportionally (when duration of phase has changed)"
            newCalculation.Text = "Distribute values automatically over time  (when duration or start/end of phase has changed)"
            showForecastMonthsOnly.Text = "in resource/cost view: show forecast only"
            adjustChilds.Text = "Date changes also affect the subordinate tasks (only used for Time)"
            avoidOverutlization.Text = "Avoid overloading resources"

        Else
            AdjustResourceNeeds.Text = "Ressourcen Bedarfe proportional anpassen (bei Änderung der Dauer der Phase) "
            newCalculation.Text = "Werte autom. über die Zeit verteilen  (bei Änderung Dauer oder Start/Ende der Phase) "
            showForecastMonthsOnly.Text = "in Ressourcen/Kosten-View: nur Forecast Monate zeigen"
            adjustChilds.Text = "Start- und Endedatum der 'Kinder' automatisch anpassen (nur für Zeitangaben verwendet)"
            avoidOverutlization.Text = "Überlastung vermeiden"

        End If

        AdjustResourceNeeds.Checked = awinSettings.propAnpassRess
        newCalculation.Checked = Not awinSettings.noNewCalculation
        ' showForecastMonthsOnly.Checked = awinSettings.enableInvoices
        showForecastMonthsOnly.Checked = Not awinSettings.noMatterActualData
        adjustChilds.Checked = awinSettings.autoAjustChilds
        ' tk added 23.11.22: to be able to allow over-Utilization
        avoidOverutlization.Checked = Not awinSettings.meAllowOverTime

        If currentProjektTafelModus <> ptModus.massEditTermine Then
            AdjustResourceNeeds.Enabled = False
            newCalculation.Enabled = False
            showForecastMonthsOnly.Enabled = False
            adjustChilds.Enabled = False
            avoidOverutlization.Enabled = True
        Else
            AdjustResourceNeeds.Enabled = True
            newCalculation.Enabled = True
            showForecastMonthsOnly.Enabled = True
            adjustChilds.Enabled = True
            avoidOverutlization.Enabled = True
        End If
    End Sub

    Private Sub frmProjectEditSettings_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        awinSettings.propAnpassRess = AdjustResourceNeeds.Checked
        awinSettings.noNewCalculation = Not newCalculation.Checked

        ' awinSettings.enableInvoices = showForecastMonthsOnly.Checked
        awinSettings.noMatterActualData = Not showForecastMonthsOnly.Checked
        awinSettings.autoAjustChilds = adjustChilds.Checked

        ' tk added 23.11.22 to be able to allow more days when capacity is available
        awinSettings.meAllowOverTime = Not avoidOverutlization.Checked

        topPos = Me.Top
        leftPos = Me.Left

    End Sub


End Class