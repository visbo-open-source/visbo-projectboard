Public Class frmMppSettings

    Public calledfrom As String

    Private Sub frmMppSettings_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        If calledfrom = "frmShowPlanElements" Then

            With awinSettings

                .eppExtendedMode = False

                If .mppSortiertDauer Then
                    .mppShowAllIfOne = True
                End If

                shwProjectLine.Checked = .mppShowProjectLine
                notStrictly.Checked = .mppShowAllIfOne
                shwAmpeln.Checked = .mppShowAmpel
                shwPhaseText.Checked = .mppShowPhName
                shwPhaseDate.Checked = .mppShowPhDate
                'phaseFullyContained.Checked = .mppFullyContained
                ShwMilestoneText.Checked = .mppShowMsName
                ShwMilestoneDate.Checked = .mppShowMsDate
                shwVerticals.Checked = .mppVertikalesRaster
                shwLegend.Checked = .mppShowLegend
                sortiertNachDauer.Checked = .mppSortiertDauer
                allOnOnePage.Checked = .mppOnePage
                shwExtendedMode.Checked = .mppExtendedMode
                KwInMilestone.Checked = .mppKwInMilestone

            End With
        ElseIf calledfrom = "frmSelectPPTTempl" Then

            With awinSettings

                .eppExtendedMode = True

                If .mppSortiertDauer Then
                    .mppShowAllIfOne = True
                End If

                shwProjectLine.Checked = .mppShowProjectLine
                notStrictly.Visible = False
                notStrictly.Checked = .mppShowAllIfOne
                shwAmpeln.Checked = .mppShowAmpel
                shwPhaseText.Checked = .mppShowPhName
                shwPhaseDate.Checked = .mppShowPhDate
                'phaseFullyContained.Checked = .mppFullyContained
                ShwMilestoneText.Checked = .mppShowMsName
                ShwMilestoneDate.Checked = .mppShowMsDate
                KwInMilestone.Checked = .mppKwInMilestone

                shwVerticals.Checked = .mppVertikalesRaster
                shwLegend.Checked = .mppShowLegend
                sortiertNachDauer.Visible = False
                'sortiertNachDauer.Checked = .mppSortiertDauer
                allOnOnePage.Visible = False
                'allOnOnePage.Checked = .mppOnePage
                shwExtendedMode.Visible = False
                'shwExtendedMode.Checked = .mppExtendedMode

                'ur: 21.04.2015: noch zu tun: zuvor alten Wert sichern
                shwExtendedMode.Checked = True
            End With

        ElseIf calledfrom = "frmBHTC" Then

            With awinSettings

                ' die folgenden werden im MS Project AddIn gar nicht gezeigt 
                ' sie müssen auch entsprechend auf die für den Project AddIn sinnvllen werte gesetzt werden 
                '
                shwProjectLine.Visible = False
                shwProjectLine.Checked = False
                .mppShowProjectLine = False

                shwAmpeln.Visible = False
                shwAmpeln.Checked = False
                .mppShowAmpel = False

                notStrictly.Visible = False
                notStrictly.Checked = False
                .mppShowAllIfOne = False

                sortiertNachDauer.Visible = False
                sortiertNachDauer.Checked = False
                .mppSortiertDauer = False
                '

                shwExtendedMode.Visible = False
                shwExtendedMode.Checked = True
                .eppExtendedMode = True
                .mppExtendedMode = True

                ' jetzt kommen die im MS Project sichtbaren Checkboxes
                shwPhaseText.Checked = .mppShowPhName
                shwPhaseDate.Checked = .mppShowPhDate

                ShwMilestoneText.Checked = .mppShowMsName
                ShwMilestoneDate.Checked = .mppShowMsDate
                KwInMilestone.Visible = True
                KwInMilestone.Checked = .mppKwInMilestone

                shwVerticals.Checked = .mppVertikalesRaster
                shwHorizontals.Checked = .mppShowHorzizontals

                shwLegend.Checked = .mppShowLegend
                allOnOnePage.Checked = .mppOnePage

                ' jetzt müssen die Checkboxes und der OK-Button noch hochgeschoben werden 
                ' ausserdem die Höhe des Formulars verändert werden 
                Dim offset As Integer = shwPhaseText.Top - shwProjectLine.Top
                shwPhaseText.Top = shwPhaseText.Top - offset
                shwPhaseDate.Top = shwPhaseDate.Top - offset

                ShwMilestoneText.Top = ShwMilestoneText.Top - offset
                ShwMilestoneDate.Top = ShwMilestoneDate.Top - offset
                KwInMilestone.Top = KwInMilestone.Top - offset

                useAbbrev.Top = useAbbrev.Top - offset

                shwVerticals.Top = shwVerticals.Top - offset
                shwHorizontals.Top = shwHorizontals.Top - offset

                shwLegend.Top = shwLegend.Top - offset
                allOnOnePage.Top = allOnOnePage.Top - offset

                okButton.Top = okButton.Top - offset

                Me.Height = Me.Height - offset


            End With


        End If

    End Sub



    Private Sub okButton_Click(sender As Object, e As EventArgs) Handles okButton.Click

        awinSettings.mppShowProjectLine = shwProjectLine.Checked
        awinSettings.mppShowAllIfOne = notStrictly.Checked
        awinSettings.mppShowAmpel = shwAmpeln.Checked

        awinSettings.mppShowPhName = shwPhaseText.Checked
        awinSettings.mppShowPhDate = shwPhaseDate.Checked
        awinSettings.mppShowMsName = ShwMilestoneText.Checked
        awinSettings.mppShowMsDate = ShwMilestoneDate.Checked
        awinSettings.mppUseAbbreviation = useAbbrev.Checked


        awinSettings.mppVertikalesRaster = shwVerticals.Checked
        awinSettings.mppShowHorzizontals = shwHorizontals.Checked
        awinSettings.mppShowLegend = shwLegend.Checked
        awinSettings.mppOnePage = allOnOnePage.Checked

        awinSettings.mppSortiertDauer = sortiertNachDauer.Checked
        awinSettings.mppExtendedMode = shwExtendedMode.Checked

        If awinSettings.mppSortiertDauer Then
            awinSettings.mppShowAllIfOne = True
        End If
        

        MyBase.Close()

    End Sub


    Private Sub notStrictly_CheckedChanged(sender As Object, e As EventArgs) Handles notStrictly.CheckedChanged

    End Sub

    Private Sub shwExtendedMode_CheckedChanged(sender As Object, e As EventArgs) Handles shwExtendedMode.CheckedChanged

    End Sub
End Class