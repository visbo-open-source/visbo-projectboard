Imports System.Windows.Forms

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
            shwInvoices.Text = "Invoice/Penalty"
            Me.Text = "Settings"
        End If

    End Sub

    Private Sub Visibility()

        shwAmpeln.Visible = False
        useOriginalNames.Visible = False
        filterEmptyProjects.Left = useOriginalNames.Left

        shwLegend.Visible = False
        sortiertNachDauer.Visible = False
        allOnOnePage.Visible = False
        shwExtendedMode.Top = allOnOnePage.Top


    End Sub

    Private Sub frmMppSettings_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Call getFrmPosition(PTfrm.other, Top, Left)

        Call languageSettings()

        Call Visibility()

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
            shwInvoices.Checked = .mppInvoicesPenalties

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

        awinSettings.mppInvoicesPenalties = shwInvoices.Checked


        MyBase.Close()

    End Sub

    Private Sub frmMppSettings_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Try
            frmCoord(PTfrm.other, PTpinfo.top) = Me.Top
            frmCoord(PTfrm.other, PTpinfo.left) = Me.Left
        Catch ex As Exception

        End Try
    End Sub
End Class