
Imports System.Globalization
Imports ProjectBoardBasic
Imports ProjectBoardDefinitions

Public Class frmEinstellungen

    Private dontFire As Boolean = False
    Private Sub frmEinstellungen_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        'chkboxMassEdit.Checked = awinSettings.meExtendedColumnsView
        chkboxAmpel.Checked = awinSettings.mppShowAmpel
        chkboxPropAnpass.Checked = awinSettings.propAnpassRess
        loadPFV.Checked = awinSettings.loadPFV

        dontFire = True
        If awinSettings.meCompareWithLastVersion Then
            rdbLast.Checked = True
        Else
            rdbFirst.Checked = True
        End If
        dontFire = False

        Dim xxx As String = awinSettings.ReportLanguage

        SprachAusw.Items.Add("Deutsch")
        SprachAusw.Items.Add("Englisch")
        ' ''SprachAusw.Items.Add("Französisch")
        ' ''SprachAusw.Items.Add("Spanisch")

        Select Case repCult.Name

            Case ReportLang(PTSprache.deutsch).Name
                SprachAusw.SelectedIndex = PTSprache.deutsch
            Case ReportLang(PTSprache.englisch).Name
                SprachAusw.SelectedIndex = PTSprache.englisch
            Case ReportLang(PTSprache.französisch).Name
                SprachAusw.SelectedIndex = PTSprache.französisch
            Case ReportLang(PTSprache.spanisch).Name
                SprachAusw.SelectedIndex = PTSprache.spanisch
            Case Else
                SprachAusw.SelectedIndex = PTSprache.deutsch

        End Select

        statusLabel.Enabled = False
        statusLabel.Visible = True
        statusLabel.Text = ""
    End Sub

    Private Sub setVisibility()
        loadPFV.Visible = myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager
        Label1.Visible = False
        SprachAusw.Visible = False
    End Sub

    'Private Sub chkboxMassEdit_CheckedChanged(sender As Object, e As EventArgs) Handles chkboxMassEdit.CheckedChanged

    '    If chkboxMassEdit.Checked Then
    '        awinSettings.meExtendedColumnsView = True
    '    Else
    '        awinSettings.meExtendedColumnsView = False
    '    End If

    'End Sub

    Private Sub chkboxPropAnpass_CheckedChanged(sender As Object, e As EventArgs) Handles chkboxPropAnpass.CheckedChanged

        If chkboxPropAnpass.Checked Then
            awinSettings.propAnpassRess = True
        Else
            awinSettings.propAnpassRess = False
        End If
    End Sub

    Private Sub chkboxAmpel_CheckedChanged(sender As Object, e As EventArgs) Handles chkboxAmpel.CheckedChanged

        If chkboxAmpel.Checked Then
            awinSettings.mppShowAmpel = True
        Else
            awinSettings.mppShowAmpel = False
        End If
    End Sub

    Private Sub SprachAusw_SelectedIndexChanged(sender As Object, e As EventArgs) Handles SprachAusw.SelectedIndexChanged

        Select Case SprachAusw.SelectedIndex
            Case PTSprache.deutsch
                repCult = ReportLang(SprachAusw.SelectedIndex)
            Case PTSprache.englisch
                repCult = ReportLang(SprachAusw.SelectedIndex)
            Case PTSprache.französisch
                repCult = ReportLang(SprachAusw.SelectedIndex)
            Case PTSprache.spanisch
                repCult = ReportLang(SprachAusw.SelectedIndex)
            Case Else
                repCult = ReportLang(PTSprache.deutsch)

        End Select

        awinSettings.ReportLanguage = repCult.Name

        repMessages = XMLImportReportMsg(repMsgFileName, awinSettings.ReportLanguage)

        Call setLanguageMessages()

        'statusLabel.Text = "Spracheinstellung aktuell auf " & repCult.DisplayName & " gesetzt!!"

        'Me.Close()

    End Sub

    Private Sub rdbFirst_CheckedChanged(sender As Object, e As EventArgs) Handles rdbFirst.CheckedChanged
        If rdbFirst.Checked = True And Not dontFire Then
            awinSettings.meCompareWithLastVersion = False
        End If
    End Sub

    Private Sub rdbLast_CheckedChanged(sender As Object, e As EventArgs) Handles rdbLast.CheckedChanged
        If rdbLast.Checked = True And Not dontFire Then
            awinSettings.meCompareWithLastVersion = True
        End If
    End Sub

    Private Sub loadPFV_CheckedChanged(sender As Object, e As EventArgs) Handles loadPFV.CheckedChanged
        awinSettings.loadPFV = loadPFV.Checked
    End Sub
End Class