
Imports System.Globalization
Imports ProjectBoardBasic
Imports ProjectBoardDefinitions

Public Class frmEinstellungen

    Private dontFire As Boolean = False
    Private Sub frmEinstellungen_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        Call languageSettings()

        'chkboxMassEdit.Checked = awinSettings.meExtendedColumnsView
        chkboxAmpel.Checked = awinSettings.mppShowAmpel
        chkboxPropAnpass.Checked = awinSettings.propAnpassRess
        chkbxPhasesAnteilig.Checked = awinSettings.phasesProzentual
        chkbxInvoices.Checked = awinSettings.enableInvoices

        If chkbx_TakeCapaFromOldOrga.Checked <> awinSettings.takeCapasFromOldOrga Then
            dontFire = True
            chkbx_TakeCapaFromOldOrga.Checked = awinSettings.takeCapasFromOldOrga
            dontFire = False
        End If

        If chkbx_KUG_active.Checked <> awinSettings.kurzarbeitActivated Then
            dontFire = True
            chkbx_KUG_active.Checked = awinSettings.kurzarbeitActivated
            dontFire = False
        End If


        If myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager Then
            loadPFV.Checked = awinSettings.loadPFV
        Else
            loadPFV.Checked = awinSettings.filterPFV
        End If

        dontFire = True
        If awinSettings.meCompareWithLastVersion Then
            rdbLast.Checked = True
        Else
            rdbFirst.Checked = True
        End If
        dontFire = False

        'Dim xxx As String = awinSettings.ReportLanguage

        If awinSettings.englishLanguage Then
            SprachAusw.Items.Add("German")
            SprachAusw.Items.Add("English")
        Else
            SprachAusw.Items.Add("Deutsch")
            SprachAusw.Items.Add("Englisch")
        End If

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

    ''' <summary>
    ''' setzt die Texte der Buttons, auch in Abhängigleit von der Rolle 
    ''' </summary>
    Private Sub languageSettings()

        If awinSettings.englishLanguage Then
            GroupBox1.Text = "Version to compare"
            rdbFirst.Text = "First"
            rdbLast.Text = "Last"
            If myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager Then
                loadPFV.Text = "always load baseline(s"
            Else
                loadPFV.Text = "filter by baseline(s and load as planning version"
            End If
            chkboxPropAnpass.Text = "adjust ressource needs proportionally"
            chkbxPhasesAnteilig.Text = "use monthly overlap percentage in phase bottleneck diagrams"
            chkboxAmpel.Text = "show traffic lights"
            chkbxInvoices.Text = "Edit Invoices & Penalties"
            chkbx_KUG_active.Text = "Short-time work possible"
            chkbx_TakeCapaFromOldOrga.Text = "Take capacities from old Organisation"
            Label1.Text = "Language for Reports"
        Else
            GroupBox1.Text = "Vergleich mit welcher Version"
            rdbFirst.Text = "Erster"
            rdbLast.Text = "Letzter"
            If myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager Then
                loadPFV.Text = "immer Vorgabe laden"
            Else
                loadPFV.Text = "auf Vorgaben filtern und als Planungs-Version laden"
            End If
            chkboxPropAnpass.Text = "Ressourcenbedarfe proportional anpassen"
            chkbxPhasesAnteilig.Text = "Phasen in Monats-Häufigkeitsdiagrammen anteilig berechnen"
            chkboxAmpel.Text = "Ampel anzeigen"
            chkbxInvoices.Text = "Rechnungen und Vertrags-Strafen bearbeiten"
            chkbx_KUG_active.Text = "Kurzarbeit ist möglich"
            Label1.Text = "Sprache für Reports"
        End If

    End Sub

    ''' <summary>
    ''' steuert, wie die Phase in Häufigkeits-Digrammen gezählt wird
    ''' true: wenn die Phase den Monat zu 10% überdeckt, wird 0,1 gerechnet, wenn sie den Monat zu 100% abdeckt wird 1 gezählt
    ''' false: sobald die Phase den Monat auch nur ein kleines bisschen überdeckt wird , 1 gezählt 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub chkbxPhasesAnteilig_CheckedChanged(sender As Object, e As EventArgs) Handles chkbxPhasesAnteilig.CheckedChanged
        If chkbxPhasesAnteilig.Checked Then
            awinSettings.phasesProzentual = True
        Else
            awinSettings.phasesProzentual = False
        End If
    End Sub

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
        If myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager Then
            awinSettings.loadPFV = loadPFV.Checked
        Else
            awinSettings.filterPFV = loadPFV.Checked
        End If

    End Sub

    Private Sub chkbxInvoices_CheckedChanged(sender As Object, e As EventArgs) Handles chkbxInvoices.CheckedChanged
        If chkbxInvoices.Checked Then
            awinSettings.enableInvoices = True
        Else
            awinSettings.enableInvoices = False
        End If
    End Sub

    Private Sub chkbx_KUG_active_CheckedChanged(sender As Object, e As EventArgs) Handles chkbx_KUG_active.CheckedChanged
        If chkbx_KUG_active.Checked Then
            awinSettings.kurzarbeitActivated = True
        Else
            awinSettings.kurzarbeitActivated = False
        End If
    End Sub

    Private Sub chkbx_TakeCapaFromOldOrga_CheckedChanged(sender As Object, e As EventArgs) Handles chkbx_TakeCapaFromOldOrga.CheckedChanged

        If Not dontFire Then
            If chkbx_TakeCapaFromOldOrga.Checked Then
                awinSettings.takeCapasFromOldOrga = True
            Else
                awinSettings.takeCapasFromOldOrga = False
            End If
        End If

    End Sub
End Class