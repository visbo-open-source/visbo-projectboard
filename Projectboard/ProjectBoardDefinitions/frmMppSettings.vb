Public Class frmMppSettings

    Public calledfrom As String

    Private Sub languageSettings()

        If awinSettings.englishLanguage Then
            shwProjectLine.Text = "Projectline"
            notStrictly.Text = "one element in Timespan: show all other as well"
            useOriginalNames.Text = "use Original Names"
            shwAmpeln.Text = "Traffic-Lights"
            filterEmptyProjects.Text = "filter empty Projects"
            shwPhaseText.Text = "annotate Phase-Name"
            shwPhaseDate.Text = "annotate Phase-Date"
            useAbbrev.Text = "use Abbreviation"
            ShwMilestoneText.Text = "annotate Milestone-Name"
            ShwMilestoneDate.Text = "annotate Milestone-Date"
            KwInMilestone.Text = "annotate CW in Milestone"
            shwVerticals.Text = "draw vertical Lines"
            shwHorizontals.Text = "draw horizontal Lines"
            shwLegend.Text = "create Legend"
            allOnOnePage.Text = "all-on-1-Page"
            sortiertNachDauer.Text = "sorted by Duration"
            shwExtendedMode.Text = "Extended Mode"
            Me.Text = "Settings"
        End If

    End Sub

    Private Sub frmMppSettings_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Call languageSettings()

        With awinSettings
            shwProjectLine.Checked = .mppShowProjectLine
            notStrictly.Checked = .mppShowAllIfOne
            shwAmpeln.Checked = .mppShowAmpel
            useOriginalNames.Checked = .mppUseOriginalNames

            shwPhaseText.Checked = .mppShowPhName
            shwPhaseDate.Checked = .mppShowPhDate
            useAbbrev.Checked = .mppUseAbbreviation
            ShwMilestoneText.Checked = .mppShowMsName
            ShwMilestoneDate.Checked = .mppShowMsDate
            KwInMilestone.Checked = .mppKwInMilestone

            shwVerticals.Checked = .mppVertikalesRaster
            shwLegend.Checked = .mppShowLegend
            sortiertNachDauer.Checked = .mppSortiertDauer
            shwHorizontals.Checked = .mppShowHorizontals
            allOnOnePage.Checked = .mppOnePage
            shwExtendedMode.Checked = .mppExtendedMode

            filterEmptyProjects.Checked = Not .mppProjectsWithNoMPmayPass

            If .mppSortiertDauer Then
                .mppShowAllIfOne = True
            End If

        End With

    End Sub



    Private Sub okButton_Click(sender As Object, e As EventArgs) Handles okButton.Click


        awinSettings.mppShowProjectLine = shwProjectLine.Checked
        awinSettings.mppShowAllIfOne = notStrictly.Checked
        awinSettings.mppShowAmpel = shwAmpeln.Checked
        awinSettings.mppUseAbbreviation = useOriginalNames.Checked

        awinSettings.mppShowPhName = shwPhaseText.Checked
        awinSettings.mppShowPhDate = shwPhaseDate.Checked
        awinSettings.mppShowMsName = ShwMilestoneText.Checked
        awinSettings.mppShowMsDate = ShwMilestoneDate.Checked
        awinSettings.mppUseAbbreviation = useAbbrev.Checked
        awinSettings.mppKwInMilestone = KwInMilestone.Checked


        awinSettings.mppVertikalesRaster = shwVerticals.Checked
        awinSettings.mppShowHorizontals = shwHorizontals.Checked
        awinSettings.mppShowLegend = shwLegend.Checked
        awinSettings.mppOnePage = allOnOnePage.Checked

        awinSettings.mppSortiertDauer = sortiertNachDauer.Checked
        awinSettings.mppExtendedMode = shwExtendedMode.Checked

        awinSettings.mppProjectsWithNoMPmayPass = Not filterEmptyProjects.Checked

        If awinSettings.mppSortiertDauer Then
            awinSettings.mppShowAllIfOne = True
        End If


        MyBase.Close()

    End Sub


    Private Sub notStrictly_CheckedChanged(sender As Object, e As EventArgs) Handles notStrictly.CheckedChanged

    End Sub

    Private Sub shwExtendedMode_CheckedChanged(sender As Object, e As EventArgs) Handles shwExtendedMode.CheckedChanged

    End Sub

    Private Sub filterEmptyProjects_CheckedChanged(sender As Object, e As EventArgs) Handles filterEmptyProjects.CheckedChanged

    End Sub

    Private Sub shwProjectLine_CheckedChanged(sender As Object, e As EventArgs) Handles shwProjectLine.CheckedChanged

    End Sub
End Class