
Imports ProjectBoardDefinitions
Imports ProjectBoardBasic
Imports MongoDbAccess
Imports ClassLibrary1
Imports WpfWindow
Imports WPFPieChart
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Excel
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Security.Principal
Imports System.Diagnostics
Imports System.Drawing
Imports System.Windows.Forms


Public Module PBBModules


    ''' <summary>
    ''' wird aus der Multiprojekt-Tafel zum Testen der Report Erstellungs-Routinen 
    ''' und aus dem MS Project AddIn aufgerufen 
    ''' </summary>
    ''' <param name="controlID"></param>
    ''' <remarks></remarks>

    Sub PBBBHTCHierarchySelAction(controlID As String, ByVal reportprofil As clsReport)

        Dim hryFormular As New frmHierarchySelection
        Dim returnValue As DialogResult
        Dim formerSettings(3) As Boolean

        If controlID = "PT1G1B3" Then
            hryFormular.calledFrom = "Multiprojekt-Tafel"

            With awinSettings
                formerSettings(0) = .mppExtendedMode
                formerSettings(1) = .mppShowAllIfOne
                formerSettings(2) = .mppShowAmpel
                formerSettings(3) = .mppFullyContained
            End With

            With awinSettings
                .mppExtendedMode = True
                .mppShowAllIfOne = False
                .mppShowAmpel = False
                .mppFullyContained = False
            End With

        Else
            hryFormular.calledFrom = "MS-Project"

            hryFormular.repProfil = New clsReport
            reportprofil.CopyTo(hryFormular.repProfil)
        End If


        ' Dim formerSettings(3) As Boolean
        With awinSettings
            formerSettings(0) = .mppExtendedMode
            formerSettings(1) = .mppShowAllIfOne
            formerSettings(2) = .mppShowAmpel
            formerSettings(3) = .mppFullyContained
        End With

        With awinSettings
            .mppExtendedMode = True
            .mppShowAllIfOne = False
            .mppShowAmpel = False
            .mppFullyContained = False
        End With

        awinSettings.useHierarchy = True
        With hryFormular

            .Text = "Projekt-Report erzeugen"
            .OKButton.Text = "Bericht erstellen"
            .menuOption = PTmenue.reportBHTC

            ' hier müssen die für BHTC nicht wählbaren Optionen gesetzt werden 
            With awinSettings
                .mppShowProjectLine = False
                .mppShowAmpel = False
                .mppShowAllIfOne = False
                .mppSortiertDauer = False
                .mppExtendedMode = True
                '.eppExtendedMode = True
            End With

            .statusLabel.Text = ""
            .statusLabel.Visible = True

            .AbbrButton.Visible = False
            .AbbrButton.Enabled = False

            .chkbxOneChart.Checked = False
            .chkbxOneChart.Visible = False


            .hryStufenLabel.Visible = False
            .hryStufen.Value = 50
            .hryStufen.Visible = False



            ' Reports
            .repVorlagenDropbox.Visible = True
            .labelPPTVorlage.Visible = True
            .einstellungen.Visible = True

            ' Filter
            .filterDropbox.Visible = True
            .filterLabel.Visible = True
            .filterLabel.Text = "Name Report-Profil"

            If Not IsNothing(reportprofil) Then
                .filterDropbox.Text = reportprofil.name
            Else
                .filterDropbox.Text = ""
            End If



            Try
                If .calledFrom = "MS-Project" Then

                    Dim lic As New clsLicences
                    Try
                        lic = XMLImportLicences(licFileName)
                    Catch ex As Exception

                    End Try

                    ' nur mit dem Recht für ProjectAdmin können ReportProfile gespeichert werden
                    If lic.validLicence(myWindowsName, LizenzKomponenten(PTSWKomp.ProjectAdmin)) Then

                        .auswSpeichern.Visible = True
                        .filterDropbox.Enabled = True
                    Else
                        .auswSpeichern.Visible = False
                        .filterDropbox.Enabled = False
                    End If
                Else

                    .auswSpeichern.Visible = False
                    .filterDropbox.Enabled = False
                End If

            Catch ex As Exception
                .auswSpeichern.Visible = False
                .filterDropbox.Enabled = False
            End Try


            ' bei Verwendung Background Worker muss Aufruf so erfolgen: 
            returnValue = .ShowDialog
        End With


        With awinSettings
            .mppExtendedMode = formerSettings(0)
            .mppShowAllIfOne = formerSettings(1)
            .mppShowAmpel = formerSettings(2)
            .mppFullyContained = formerSettings(3)
        End With


    End Sub

    ''' <summary>
    ''' wird aus der Multiprojekt-Tafel aufgerufen 
    ''' </summary>
    ''' <param name="controlID"></param>
    ''' <remarks></remarks>
    Sub PBBNameHierarchySelAction(controlID As String)


        Dim nameFormular As New frmNameSelection
        Dim hryFormular As New frmHierarchySelection
        Dim awinSelection As Excel.ShapeRange
        Dim returnValue As DialogResult

        Call projektTafelInit()

        hryFormular.calledFrom = "Multiprojekt-Tafel"


        ' gibt es überhaupt Objekte, zu denen man was anzeigen kann ? 
        'If ShowProjekte.Count > 0 And showRangeRight - showRangeLeft > 5 Then

        If controlID = "Pt6G3M1B1" Then
            ' normale, volle Auswahl des filters ; Namens-Definition

            With nameFormular

                .Text = "Datenbank Filter definieren"
                .OKButton.Text = "Speichern"
                .menuOption = PTmenue.filterdefinieren
                .statusLabel.Text = ""
                .statusLabel.Visible = True

                .rdbRoles.Enabled = True
                .rdbCosts.Enabled = True

                .rdbBU.Visible = True
                .pictureBU.Visible = True

                .rdbTyp.Visible = True
                .pictureTyp.Visible = True

                .einstellungen.Visible = False

                .chkbxOneChart.Checked = False
                .chkbxOneChart.Visible = False

                ' Reports 
                .repVorlagenDropbox.Visible = False
                .labelPPTVorlage.Visible = False
                .einstellungen.Visible = False

                ' Filter
                .filterDropbox.Visible = True
                .filterLabel.Visible = True
                .filterLabel.Text = "Name des Filters"

                ' Auswahl Speichern
                .auswSpeichern.Visible = False
                .auswSpeichern.Enabled = False

                returnValue = .ShowDialog

            End With

        ElseIf controlID = "Pt6G3M1B2" Then

            awinSettings.useHierarchy = True

            With hryFormular

                .Text = "Datenbank Filter definieren"
                .OKButton.Text = "Speichern"
                .menuOption = PTmenue.filterdefinieren
                .statusLabel.Text = ""
                .statusLabel.Visible = True

                .AbbrButton.Visible = False
                .AbbrButton.Enabled = False

                .chkbxOneChart.Checked = False
                .chkbxOneChart.Visible = False

                ' Reports 
                .repVorlagenDropbox.Visible = False
                .labelPPTVorlage.Visible = False
                .einstellungen.Visible = False

                ' Filter
                .filterDropbox.Visible = True
                .filterLabel.Visible = True
                .filterLabel.Text = "Name des Filters"

                ' Auswahl Speichern
                .auswSpeichern.Visible = False
                .auswSpeichern.Enabled = False

                .einstellungen.Visible = False

                returnValue = .ShowDialog
            End With


        ElseIf ShowProjekte.Count > 0 Then

            If awinSettings.isHryNameFrmActive Then
                Call MsgBox("es kann nur ein Fenster zur Hierarchie- bzw. Namenauswahl geöffnet sein ...")

            ElseIf controlID = "PTXG1B4" Or controlID = "PT0G1B8" Then
                ' Namen auswählen, Visualisieren
                awinSettings.useHierarchy = False
                With nameFormular
                    .Text = "Plan-Elemente visualisieren"
                    .OKButton.Text = "Anzeigen"
                    .menuOption = PTmenue.visualisieren
                    .statusLabel.Text = ""
                    .statusLabel.Visible = True


                    .rdbBU.Visible = False
                    .pictureBU.Visible = False
                    .rdbTyp.Visible = False
                    .pictureTyp.Visible = False
                    .rdbRoles.Visible = True
                    .pictureRoles.Visible = True
                    .rdbCosts.Visible = True
                    .pictureCosts.Visible = True

                    ' Leistbarkeits-Charts
                    .chkbxOneChart.Checked = False
                    .chkbxOneChart.Visible = False

                    ' Reports 
                    .repVorlagenDropbox.Visible = False
                    .labelPPTVorlage.Visible = False
                    .einstellungen.Visible = False

                    ' Filter
                    .filterDropbox.Visible = True
                    .filterLabel.Visible = True
                    .filterLabel.Text = "Auswahl"


                    ' Nicht Modal anzeigen
                    .Show()
                    'returnValue = .ShowDialog
                End With

            ElseIf controlID = "PTXG1B5" Or controlID = "PT0G1B9" Then
                ' Hierarchie auswählen, visualisieren
                awinSettings.useHierarchy = True
                With hryFormular
                    .Text = "Plan-Elemente visualisieren"
                    .OKButton.Text = "Anzeigen"
                    .AbbrButton.Visible = False
                    .AbbrButton.Enabled = False
                    .menuOption = PTmenue.visualisieren
                    .statusLabel.Text = ""
                    .statusLabel.Visible = True

                    .chkbxOneChart.Checked = False
                    .chkbxOneChart.Visible = False

                    ' Reports
                    .repVorlagenDropbox.Visible = False
                    .labelPPTVorlage.Visible = False
                    .einstellungen.Visible = False

                    ' Filter
                    .filterDropbox.Visible = True
                    .filterLabel.Visible = True
                    .filterLabel.Text = "Auswahl"


                    ' Nicht Modal anzeigen
                    .Show()
                    'returnValue = .ShowDialog
                End With
            ElseIf controlID = "PTXG1B6" Then
                ' Namen auswählen, Leistbarkeit

                awinSettings.useHierarchy = False
                With nameFormular
                    .Text = "Leistbarkeits-Charts erstellen"
                    .OKButton.Text = "Charts erstellen"
                    .menuOption = PTmenue.leistbarkeitsAnalyse
                    .statusLabel.Text = ""
                    .statusLabel.Visible = True


                    .rdbBU.Visible = False
                    .pictureBU.Visible = False
                    .rdbTyp.Visible = False
                    .pictureTyp.Visible = False

                    .rdbRoles.Visible = True
                    .pictureRoles.Visible = True
                    .rdbCosts.Visible = True
                    .pictureCosts.Visible = True

                    ' Leistbarkeits-Charts
                    .chkbxOneChart.Checked = False
                    .chkbxOneChart.Visible = True

                    ' Reports 
                    .repVorlagenDropbox.Visible = False
                    .labelPPTVorlage.Visible = False

                    ' Filter
                    .filterDropbox.Visible = True
                    .filterLabel.Visible = True
                    .filterLabel.Text = "Auswahl"

                    ' Nicht Modal anzeigen
                    .Show()
                    'returnValue = .ShowDialog
                End With
            ElseIf controlID = "PTXG1B7" Then
                ' Hierarchie auswählen, Leistbarkeit
                awinSettings.useHierarchy = True
                With hryFormular
                    .Text = "Leistbarkeits-Charts erstellen"
                    .OKButton.Text = "Charts erstellen"
                    .AbbrButton.Visible = False
                    .AbbrButton.Enabled = False
                    .menuOption = PTmenue.leistbarkeitsAnalyse
                    .statusLabel.Text = ""
                    .statusLabel.Visible = True


                    .chkbxOneChart.Checked = False
                    .chkbxOneChart.Visible = True

                    ' Reports
                    .repVorlagenDropbox.Visible = False
                    .labelPPTVorlage.Visible = False
                    .einstellungen.Visible = False

                    ' Filter
                    .filterDropbox.Visible = True
                    .filterLabel.Visible = True
                    .filterLabel.Text = "Auswahl"

                    ' Nicht Modal anzeigen
                    .Show()
                    'returnValue = .ShowDialog
                End With


            ElseIf controlID = "PT1G1M1B1" Then
                ' Namen auswählen, Einzelprojekt Berichte 

                Try
                    awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
                Catch ex As Exception
                    awinSelection = Nothing
                End Try

                If awinSelection Is Nothing Then
                    Call MsgBox("vorher Projekt/e selektieren ...")
                Else

                    ' wenn nachher .showdialog aufgerufen wird, müssen die beiden Settings erst auf 
                    ' false, dann auf True gesetzt werden
                    ' bei .show darf das nicht gemacht werden ! 
                    appInstance.ScreenUpdating = False
                    appInstance.EnableEvents = False

                    With nameFormular

                        .Text = "Projekt-Varianten Report erzeugen"
                        .OKButton.Text = "Bericht erstellen"
                        .menuOption = PTmenue.einzelprojektReport
                        .statusLabel.Text = ""
                        .statusLabel.Visible = True

                        .rdbRoles.Enabled = False
                        .rdbCosts.Enabled = False

                        .rdbBU.Enabled = False
                        .rdbBU.Visible = False
                        .pictureBU.Visible = False

                        .rdbTyp.Enabled = False
                        .rdbTyp.Visible = False
                        .pictureTyp.Visible = False


                        .einstellungen.Visible = True

                        .chkbxOneChart.Checked = False
                        .chkbxOneChart.Visible = False

                        .repVorlagenDropbox.Visible = True
                        .labelPPTVorlage.Visible = True

                        ' Filter
                        .filterDropbox.Visible = True
                        .filterLabel.Visible = True
                        .filterLabel.Text = "Auswahl"


                        '.Show()
                        ' bei Reports mit der Background Worker Behandlung 
                        returnValue = .ShowDialog()
                    End With

                    appInstance.ScreenUpdating = True
                    appInstance.EnableEvents = True

                End If

            ElseIf controlID = "PT1G1M1B2" Then

                Try
                    awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
                Catch ex As Exception
                    awinSelection = Nothing
                End Try

                If awinSelection Is Nothing Then
                    Call MsgBox("vorher Projekt/e selektieren ...")
                Else


                    ' wenn nachher .showdialog aufgerufen wird, müssen die beiden Settings erst auf 
                    ' dalse, dann auf True gesetzt werden
                    ' bei .show darf das nicht gemacht werden ! 
                    appInstance.ScreenUpdating = False
                    appInstance.EnableEvents = False

                    awinSettings.useHierarchy = True
                    With hryFormular
                        .Text = "Projekt-Varianten Report erzeugen"
                        .OKButton.Text = "Bericht erstellen"
                        .menuOption = PTmenue.einzelprojektReport
                        .statusLabel.Text = ""
                        .statusLabel.Visible = True

                        .AbbrButton.Visible = False
                        .AbbrButton.Enabled = False

                        .chkbxOneChart.Checked = False
                        .chkbxOneChart.Visible = False


                        ' Reports
                        .repVorlagenDropbox.Visible = True
                        .labelPPTVorlage.Visible = True
                        .einstellungen.Visible = True

                        ' Filter
                        .filterDropbox.Visible = True
                        .filterLabel.Visible = True
                        .filterLabel.Text = "Name des Filters"


                        ' bei Verwendung Background Worker muss Modal erfolgen 
                        '.Show()
                        returnValue = .ShowDialog
                    End With

                    appInstance.ScreenUpdating = True
                    appInstance.EnableEvents = True
                End If

            ElseIf controlID = "PT1G1M2B1" Then


                If showRangeLeft > 0 And showRangeRight > showRangeLeft Then
                    ' Namen Auswahl, Multiprojekt Report
                    ' wenn nachher .showdialog aufgerufen wird, müssen die beiden Settings erst auf 
                    ' dalse, dann auf True gesetzt werden
                    ' bei .show darf das nicht gemacht werden ! 
                    appInstance.ScreenUpdating = False
                    appInstance.EnableEvents = False

                    With nameFormular

                        .Text = "Multiprojekt Reports erzeugen"
                        .OKButton.Text = "Bericht erstellen"
                        .menuOption = PTmenue.multiprojektReport
                        .statusLabel.Text = ""
                        .statusLabel.Visible = True

                        .rdbRoles.Enabled = True
                        .rdbCosts.Enabled = True

                        .rdbBU.Enabled = False
                        .rdbBU.Visible = False
                        .pictureBU.Visible = False

                        .rdbTyp.Enabled = False
                        .rdbTyp.Visible = False
                        .pictureTyp.Visible = False


                        .einstellungen.Visible = True

                        .chkbxOneChart.Checked = False
                        .chkbxOneChart.Visible = False

                        .repVorlagenDropbox.Visible = True
                        .labelPPTVorlage.Visible = True

                        ' Filter
                        .filterDropbox.Visible = True
                        .filterLabel.Visible = True
                        .filterLabel.Text = "Auswahl"

                        ' .show; bei Verwendung mit Background Worker Funktion muss das modal erfolgen
                        returnValue = .ShowDialog
                    End With

                    appInstance.ScreenUpdating = True
                    appInstance.EnableEvents = True

                Else

                    Call MsgBox("Bitte wählen Sie den Zeitraum aus, für den der Report erstellt werden soll!")

                End If

            ElseIf controlID = "PT1G1M2B2" Then

                If showRangeLeft > 0 And showRangeRight > showRangeLeft Then

                    ' Hierarchie Auswahl, Multiprojekt Report
                    ' wenn nachher .showdialog aufgerufen wird, müssen die beiden Settings erst auf 
                    ' dalse, dann auf True gesetzt werden
                    ' bei .show darf das nicht gemacht werden ! 
                    appInstance.ScreenUpdating = False
                    appInstance.EnableEvents = False

                    awinSettings.useHierarchy = True
                    With hryFormular

                        .Text = "Multiprojekt Reports erzeugen"
                        .OKButton.Text = "Bericht erstellen"
                        .menuOption = PTmenue.multiprojektReport
                        .statusLabel.Text = ""
                        .statusLabel.Visible = True

                        .AbbrButton.Visible = False
                        .AbbrButton.Enabled = False

                        .chkbxOneChart.Checked = False
                        .chkbxOneChart.Visible = False

                        ' Reports
                        .repVorlagenDropbox.Visible = True
                        .labelPPTVorlage.Visible = True
                        .einstellungen.Visible = True

                        ' Filter
                        .filterDropbox.Visible = True
                        .filterLabel.Visible = True
                        .filterLabel.Text = "Auswahl"


                        ' .show; bei Verwendung mit Background Worker Funktion muss das modal erfolgen
                        returnValue = .ShowDialog
                    End With

                    appInstance.ScreenUpdating = True
                    appInstance.EnableEvents = True

                Else

                    Call MsgBox("Bitte wählen Sie den Zeitraum aus, für den der Report erstellt werden soll!")


                End If

            ElseIf controlID = "PT4G1M0B1" Then
                ' Auswahl über Namen, Typ II Export
                ' wenn nachher .showdialog aufgerufen wird, müssen die beiden Settings erst auf 
                ' dalse, dann auf True gesetzt werden
                ' bei .show darf das nicht gemacht werden ! 
                appInstance.ScreenUpdating = False
                appInstance.EnableEvents = False

                With nameFormular

                    .Text = "Excel Report erzeugen"
                    .OKButton.Text = "Report erstellen"
                    .menuOption = PTmenue.excelExport
                    .statusLabel.Text = ""

                    .rdbRoles.Enabled = False
                    .rdbCosts.Enabled = False

                    .rdbBU.Visible = True
                    .pictureBU.Visible = True

                    .rdbTyp.Visible = True
                    .pictureTyp.Visible = True

                    .einstellungen.Visible = False

                    .chkbxOneChart.Checked = False
                    .chkbxOneChart.Visible = False

                    .repVorlagenDropbox.Visible = False
                    .labelPPTVorlage.Visible = False

                    ' Filter
                    .filterDropbox.Visible = True
                    .filterLabel.Visible = True
                    .filterLabel.Text = "Auswahl"


                    returnValue = .ShowDialog
                End With

                appInstance.ScreenUpdating = True
                appInstance.EnableEvents = True

            ElseIf controlID = "PT4G1M0B2" Then

                ' Auswahl über Hierarchie, Typ II Export
                ' wenn nachher .showdialog aufgerufen wird, müssen die beiden Settings erst auf 
                ' dalse, dann auf True gesetzt werden
                ' bei .show darf das nicht gemacht werden ! 
                appInstance.ScreenUpdating = False
                appInstance.EnableEvents = False

                awinSettings.useHierarchy = True

                With hryFormular

                    .Text = "Excel Report erzeugen"
                    .OKButton.Text = "Report erstellen"
                    .menuOption = PTmenue.excelExport
                    .statusLabel.Text = ""

                    .AbbrButton.Visible = False
                    .AbbrButton.Enabled = False

                    .chkbxOneChart.Checked = False
                    .chkbxOneChart.Visible = False

                    ' Reports
                    .repVorlagenDropbox.Visible = False
                    .labelPPTVorlage.Visible = False

                    ' Filter
                    .filterDropbox.Visible = True
                    .filterLabel.Visible = True
                    .filterLabel.Text = "Auswahl"

                    .einstellungen.Visible = False

                    ' Nicht Modal anzeigen
                    '.Show()
                    returnValue = .ShowDialog
                End With

                appInstance.ScreenUpdating = True
                appInstance.EnableEvents = True

            ElseIf controlID = "PT4G1M2B1" Then
                ' Auswahl über Namen, Vorlagen erzeugen
                ' Auswahl über Hierarchie, Typ II Export
                ' wenn nachher .showdialog aufgerufen wird, müssen die beiden Settings erst auf 
                ' dalse, dann auf True gesetzt werden
                ' bei .show darf das nicht gemacht werden ! 
                appInstance.ScreenUpdating = False
                appInstance.EnableEvents = False

                With nameFormular

                    .Text = "modulare Vorlagen erzeugen"
                    .OKButton.Text = "Vorlage erstellen"
                    .menuOption = PTmenue.vorlageErstellen
                    .statusLabel.Text = ""

                    .rdbRoles.Enabled = False
                    .rdbCosts.Enabled = False

                    .rdbBU.Visible = False
                    .pictureBU.Visible = False

                    .rdbTyp.Visible = False
                    .pictureTyp.Visible = False

                    .einstellungen.Visible = False

                    .chkbxOneChart.Checked = False
                    .chkbxOneChart.Visible = False

                    .repVorlagenDropbox.Visible = False
                    .labelPPTVorlage.Visible = False

                    ' Filter
                    .filterDropbox.Visible = True
                    .filterLabel.Visible = True
                    .filterLabel.Text = "Auswahl"

                    returnValue = .ShowDialog
                End With

                appInstance.ScreenUpdating = True
                appInstance.EnableEvents = True


            ElseIf controlID = "PT4G1M2B2" Then
                ' Auswahl über Hierarchie, Vorlagen Export

                ' wenn nachher .showdialog aufgerufen wird, müssen die beiden Settings erst auf 
                ' dalse, dann auf True gesetzt werden
                ' bei .show darf das nicht gemacht werden ! 
                appInstance.ScreenUpdating = False
                appInstance.EnableEvents = False

                awinSettings.useHierarchy = True
                With hryFormular

                    .Text = "modulare Vorlagen erzeugen"
                    .OKButton.Text = "Vorlage erstellen"
                    .menuOption = PTmenue.vorlageErstellen
                    .statusLabel.Text = ""

                    .AbbrButton.Visible = False
                    .AbbrButton.Enabled = False

                    .chkbxOneChart.Checked = False
                    .chkbxOneChart.Visible = False

                    ' Reports
                    .repVorlagenDropbox.Visible = False
                    .labelPPTVorlage.Visible = False
                    .einstellungen.Visible = False

                    ' Filter
                    .filterDropbox.Visible = True
                    .filterLabel.Visible = True
                    .filterLabel.Text = "Auswahl"

                    ' Nicht Modal anzeigen
                    '.Show()
                    returnValue = .ShowDialog
                End With

                appInstance.ScreenUpdating = True
                appInstance.EnableEvents = True

            ElseIf controlID = "PT0G1M2B7" Then
                ' Auswahl über Namen, Meilensteine für Meilenstein Trendanalyse
                Try
                    awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
                Catch ex As Exception
                    awinSelection = Nothing
                End Try

                If awinSelection Is Nothing Then
                    Call MsgBox("vorher Projekt/e selektieren ...")
                Else

                    ' wenn nachher .showdialog aufgerufen wird, müssen die beiden Settings erst auf 
                    ' dalse, dann auf True gesetzt werden
                    ' bei .show darf das nicht gemacht werden ! 
                    appInstance.ScreenUpdating = False
                    appInstance.EnableEvents = False

                    With nameFormular

                        .Text = "Meilenstein Trendanalyse erzeugen"
                        .OKButton.Text = "Anzeigen"
                        .menuOption = PTmenue.meilensteinTrendanalyse
                        .statusLabel.Text = ""
                        .statusLabel.Visible = True

                        .headerLine.Text = "Meilensteine"

                        .picturePhasen.Visible = False
                        .rdbPhases.Visible = False
                        .rdbPhases.Checked = False
                        .rdbPhases.Enabled = False

                        .pictureMilestones.Visible = False
                        .rdbMilestones.Visible = False
                        .rdbMilestones.Checked = True
                        .rdbMilestones.Enabled = False

                        .pictureRoles.Visible = False
                        .rdbRoles.Visible = False
                        .rdbRoles.Checked = False
                        .rdbRoles.Enabled = False

                        .pictureCosts.Visible = False
                        .rdbCosts.Visible = False
                        .rdbCosts.Checked = False
                        .rdbCosts.Enabled = False

                        .rdbBU.Visible = False
                        .pictureBU.Visible = False

                        .rdbTyp.Visible = False
                        .pictureTyp.Visible = False

                        .einstellungen.Visible = False

                        .chkbxOneChart.Checked = False
                        .chkbxOneChart.Visible = False

                        .repVorlagenDropbox.Visible = False
                        .labelPPTVorlage.Visible = False

                        .auswSpeichern.Visible = False

                        returnValue = .ShowDialog()
                    End With

                    appInstance.ScreenUpdating = True
                    appInstance.EnableEvents = True

                End If


            End If
        Else
            Call MsgBox("Es sind keine Projekte sichtbar!  ")
        End If



        ' oben ist es de-aktiviert 
        'appInstance.EnableEvents = True
        'enableOnUpdate = True

    End Sub

    Sub PBBAnalyseLeistbarkeit001(ByVal ControlID As String)

        Dim namensFormular As New frmNameSelection
        Dim hierarchieFormular As New frmHierarchySelection
        Dim returnValue As DialogResult


        Call projektTafelInit()

        enableOnUpdate = False
        appInstance.EnableEvents = False

        ' gibt es überhaupt Objekte, zu denen man was anzeigen kann ? 
        If ShowProjekte.Count > 0 And showRangeRight - showRangeLeft > 5 Then

            If ControlID = "PTXG1B6" Then
                ' Auswahl über Namen

                With namensFormular
                    .Text = "Leistbarkeit analysieren"

                    .rdbBU.Visible = False
                    .pictureBU.Visible = False

                    .rdbTyp.Visible = False
                    .pictureTyp.Visible = False

                    .rdbRoles.Visible = True
                    .pictureRoles.Visible = True

                    .rdbCosts.Visible = True
                    .pictureCosts.Visible = True

                    '.chkbxShowObjects = False

                    .chkbxOneChart.Checked = False
                    .chkbxOneChart.Visible = True

                    '.chkbxCreateCharts = True


                    .repVorlagenDropbox.Visible = False
                    .labelPPTVorlage.Visible = False

                    '.showModePortfolio = True

                    .menuOption = PTmenue.leistbarkeitsAnalyse
                    .OKButton.Text = "Charts erstellen"

                    '.Show()
                    returnValue = .ShowDialog
                End With


            Else
                ' Auswahl über Hierarchie
                ' Hierarchie
                awinSettings.useHierarchy = True
                With hierarchieFormular
                    .Text = "Leistbarkeit analysieren"

                    .chkbxOneChart.Checked = False
                    .chkbxOneChart.Visible = True

                    '.chkbxCreateCharts = False


                    .repVorlagenDropbox.Visible = False
                    .labelPPTVorlage.Visible = False

                    '.showModePortfolio = True
                    .menuOption = PTmenue.leistbarkeitsAnalyse

                    .OKButton.Text = "Charts erstellen"

                    '.Show()
                    returnValue = .ShowDialog
                End With

            End If

        ElseIf ShowProjekte.Count = 0 Then

            Call MsgBox("Es sind keine Projekte geladen!  ")

        ElseIf showRangeRight - showRangeLeft <= 5 Then

            Call MsgBox("bitte zuerst einen Zeitraum markieren! ")

        End If



        appInstance.EnableEvents = True
        enableOnUpdate = True



    End Sub
    ''' <summary>
    ''' eine neue Variante anlegen 
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub PBBVarianteNeu(control As IRibbonControl)

        Dim hproj As clsProjekt
        Dim awinSelection As Excel.ShapeRange
        Dim neueVariante As New frmCreateNewVariant
        Dim resultat As DialogResult
        ''Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
        Dim newproj As clsProjekt
        Dim key As String
        Dim phaseList As New Collection
        Dim milestoneList As New Collection
        Dim neuerVariantenName As String = ""
        Dim ok As Boolean = True
        Dim zaehler As Integer = 1
        Dim nameCollection As New Collection
        Dim abbruch As Boolean = False


        Call projektTafelInit()

        enableOnUpdate = False

        Try
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            For i As Integer = 1 To awinSelection.count
                nameCollection.Add(awinSelection.Item(i).Name)
            Next

            While zaehler <= nameCollection.Count And Not abbruch

                ' jetzt die Aktion durchführen ...
                Dim pName As String = CStr(nameCollection.Item(zaehler))

                Try
                    hproj = ShowProjekte.getProject(pName)
                    pName = hproj.name
                    phaseList = projectboardShapes.getPhaseList(hproj.name)
                    milestoneList = projectboardShapes.getMilestoneList(hproj.name)
                Catch ex As Exception
                    Call MsgBox("Projekt " & pName & " nicht gefunden ...")
                    enableOnUpdate = True
                    Exit Sub
                End Try

                ' enableevents wird hier nicht false gesetzt; wenn dann wird das im Formular gemacht 
                ' screenupdating wird hier ebenso nicht auf false gesetzt 

                ' jetzt wird hier das Formular aufgerufen, wo eine neue Variante eingegeben werden kann 
                With neueVariante
                    .projektName.Text = hproj.name
                    .variantenName.Text = hproj.variantName
                    .newVariant.Text = neuerVariantenName
                End With

                resultat = neueVariante.ShowDialog
                If resultat = DialogResult.OK Then

                    newproj = New clsProjekt
                    hproj.copyTo(newproj)

                    If newproj.dauerInDays <> hproj.dauerInDays Then
                        'Call MsgBox("ungleich: " & newproj.dauerInDays & " versus " & hproj.dauerInDays)
                    End If

                    neuerVariantenName = neueVariante.newVariant.Text

                    With newproj
                        .name = hproj.name
                        .variantName = neuerVariantenName
                        .ampelErlaeuterung = hproj.ampelErlaeuterung
                        .ampelStatus = hproj.ampelStatus
                        .timeStamp = Date.Now
                        .shpUID = hproj.shpUID
                        .tfZeile = hproj.tfZeile
                        .Status = ProjektStatus(0)
                        If Not IsNothing(hproj.budgetWerte) Then
                            .budgetWerte = hproj.budgetWerte
                        End If

                    End With

                    ' jetzt muss die bisherige Variante aus Showprojekte rausgenommen werden ..
                    ShowProjekte.Remove(hproj.name)

                    ' die neue Variante wird aufgenommen
                    key = calcProjektKey(newproj)
                    AlleProjekte.Add(key, newproj)
                    ShowProjekte.Add(newproj)

                    ' wenn bestimmte Projekte beim Suchen nach einem Platz nicht berücksichtigt werden sollen,
                    ' dann müssen sie in einer Collection an ZeichneProjektinPlanTafel übergeben werden 
                    Try

                        Dim tmpCollection As New Collection
                        Call ZeichneProjektinPlanTafel(tmpCollection, newproj.name, newproj.tfZeile, phaseList, milestoneList)

                    Catch ex As Exception

                        Call MsgBox("Fehler bei Zeichnen Projekt: " & ex.Message)

                    End Try

                    zaehler = zaehler + 1
                Else
                    abbruch = True
                End If

            End While

        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If

        enableOnUpdate = True

    End Sub
    ''' <summary>
    ''' Es werden Projekte, die Varianten haben angezeigt in einem TreeView
    ''' Hier können Varianten ausgewählt werden, die gelöscht werden sollen
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub PBBVarianteLoeschen(control As IRibbonControl)

        Dim deletedProj As Integer = 0
        'Dim returnValue As DialogResult

        'Dim activateVariant As New frmDeleteProjects
        Dim deleteVariant As New frmProjPortfolioAdmin

        Try

            With deleteVariant
                .Text = "Variante löschen"
                .aKtionskennung = PTTvActions.deleteV
                .OKButton.Visible = True
                .OKButton.Text = "Löschen"
                '' '' ''.portfolioName.Visible = False
                '' '' ''.Label1.Visible = False
            End With

            'returnValue = activateVariant.ShowDialog
            deleteVariant.Show()

            'If returnValue = DialogResult.OK Then
            '    'deletedProj = RemoveSelectedProjectsfromDB(deleteProjects.selectedItems)    ' es werden die selektierten Projekte in der DB gespeichert, die Anzahl gespeicherter Projekte sind das Ergebnis

            'Else
            '    ' returnValue = DialogResult.Cancel

            'End If

        Catch ex As Exception

            Call MsgBox(ex.Message)
        End Try


    End Sub
    ''' <summary>
    ''' Projekt löschen
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub PBBLoeschen(control As IRibbonControl)

        Dim bestaetigeLoeschen As New frmconfirmDeletePrj
        Dim singleShp As Excel.Shape
        Dim awinSelection As Excel.ShapeRange
        Dim returnValue As DialogResult

        Call projektTafelInit()

        appInstance.EnableEvents = False
        enableOnUpdate = False

        Try
            'awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            bestaetigeLoeschen.botschaft = "Bitte bestätigen Sie das Löschen" & vbLf & _
                                            "Vorsicht: alle Varianten werden mitgelöscht ..."
            returnValue = bestaetigeLoeschen.ShowDialog

            If returnValue = DialogResult.Cancel Then

                appInstance.EnableEvents = True
                enableOnUpdate = True
                Exit Sub

            End If



            ' jetzt die Aktion durchführen ...


            For Each singleShp In awinSelection


                Dim shapeArt As Integer
                shapeArt = kindOfShape(singleShp)

                With singleShp
                    If isProjectType(shapeArt) Then

                        Try
                            Call awinDeleteProjectInSession(pName:=.Name)

                        Catch ex As Exception
                            Exit For
                        End Try

                    End If
                End With


            Next

            ' ein oder mehrere Projekte wurden gelöscht  - typus = 3
            Call awinNeuZeichnenDiagramme(3)

        Else

            Dim deletedProj As Integer = 0

            If AlleProjekte.Count = 0 Then
                Call MsgBox("es sind keine Projekte geladen !")
            Else

                'Dim deleteProjects As New frmDeleteProjects
                Dim deleteProjects As New frmProjPortfolioAdmin
                Try

                    With deleteProjects
                        .Text = "Projekte, Varianten aus der Session löschen"
                        .aKtionskennung = PTTvActions.delFromSession
                        .OKButton.Text = "Löschen"
                        '' '' ''.portfolioName.Visible = False
                        '' '' ''.Label1.Visible = False
                    End With

                    returnValue = deleteProjects.ShowDialog

                    If returnValue = DialogResult.OK Then

                        'Call MsgBox("ok, aus Session gelöscht  !")

                    Else
                        ' returnValue = DialogResult.Cancel

                    End If

                Catch ex As Exception

                    Call MsgBox(ex.Message)
                End Try

            End If



        End If

        Call awinDeSelect()

        enableOnUpdate = True
        appInstance.EnableEvents = True

    End Sub



    ''' <summary>
    ''' lädt die gewählten Projekte und gewählten Varianten in die Session
    ''' </summary>
    ''' <param name="Control"></param>
    ''' <remarks></remarks>
    Public Sub PBBDatenbankLoadProjekte(Control As IRibbonControl)

        Dim deletedProj As Integer = 0
        Dim returnValue As DialogResult

        'Dim deleteProjects As New frmDeleteProjects
        Dim loadProjectsForm As New frmProjPortfolioAdmin

        Try

            With loadProjectsForm
                .Text = "Projekte und Varianten in die Session laden "
                .aKtionskennung = PTTvActions.loadPV
                .OKButton.Text = "Laden"
                '' '' ''.portfolioName.Visible = False
                '' '' ''.Label1.Visible = False
            End With

            returnValue = loadProjectsForm.ShowDialog

            If returnValue = DialogResult.OK Then
                'deletedProj = RemoveSelectedProjectsfromDB(deleteProjects.selectedItems)    ' es werden die selektierten Projekte in der DB gespeichert, die Anzahl gespeicherter Projekte sind das Ergebnis

            Else
                ' returnValue = DialogResult.Cancel

            End If

        Catch ex As Exception

            Call MsgBox(ex.Message)
        End Try




    End Sub
    ''' <summary>
    ''' löscht die ausgewählten Projekte aus der Datenbank 
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub PBBDeleteProjectsInDB(control As IRibbonControl)


        Dim deletedProj As Integer = 0
        Dim returnValue As DialogResult

        'Dim deleteProjects As New frmDeleteProjects
        Dim deleteProjects As New frmProjPortfolioAdmin

        Try

            With deleteProjects
                .Text = "Projekte, Varianten bzw. Snapshots in der Datenbank löschen"
                .aKtionskennung = PTTvActions.delFromDB
                .OKButton.Text = "Löschen"
                '' '' ''.portfolioName.Visible = False
                '' '' ''.Label1.Visible = False
            End With

            returnValue = deleteProjects.ShowDialog

            ' die Operation ist bereits ausgeführt - deswegen muss hier nichts mehr unterschieden werden 

            If returnValue = DialogResult.OK Then
                ' everything is done ... 

            Else
                ' everything is done ... 

            End If

        Catch ex As Exception

            Call MsgBox(ex.Message)
        End Try



    End Sub
    ''' <summary>
    ''' aktiviert die selektierte Variante 
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub PBBVarianteAktiv(control As IRibbonControl)

        Dim deletedProj As Integer = 0
        'Dim returnValue As DialogResult

        'Dim activateVariant As New frmDeleteProjects
        Dim activateVariant As New frmProjPortfolioAdmin

        Try

            With activateVariant
                .Text = "Variante aktivieren"
                .aKtionskennung = PTTvActions.activateV
                .OKButton.Visible = False
                '.OKButton.Text = "Löschen"
                '' '' ''.portfolioName.Visible = False
                '' '' ''.Label1.Visible = False
            End With

            'returnValue = activateVariant.ShowDialog
            activateVariant.Show()

            'If returnValue = DialogResult.OK Then
            '    'deletedProj = RemoveSelectedProjectsfromDB(deleteProjects.selectedItems)    ' es werden die selektierten Projekte in der DB gespeichert, die Anzahl gespeicherter Projekte sind das Ergebnis

            'Else
            '    ' returnValue = DialogResult.Cancel

            'End If

        Catch ex As Exception

            Call MsgBox(ex.Message)
        End Try


    End Sub

    Sub PBBShowTimeMachine(control As IRibbonControl)

        Dim hproj As clsProjekt
        Dim pName As String, variantName As String
        Dim vglName As String = " "
        Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
        Dim singleShp As Excel.Shape
        Dim showCharacteristics As New frmShowProjCharacteristics
        'Dim returnValue As DialogResult
        Dim awinSelection As Excel.ShapeRange
        Dim grueneAmpel As String = awinPath & "gruen.gif"
        Dim gelbeAmpel As String = awinPath & "gelb.gif"
        Dim roteAmpel As String = awinPath & "rot.gif"
        Dim graueAmpel As String = awinPath & "grau.gif"

        If timeMachineIsOn Then
            Call MsgBox("bitte erst Time Machine beenden ...")
            Exit Sub
        End If

        Call projektTafelInit()

        enableOnUpdate = False
        appInstance.EnableEvents = True


        Try
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then


            If awinSelection.Count = 1 And isProjectType(kindOfShape(awinSelection.Item(1))) Then
                ' jetzt die Aktion durchführen ...
                singleShp = awinSelection.Item(1)
                hproj = ShowProjekte.getProject(singleShp.Name, True)
                With hproj
                    pName = .name
                    variantName = .variantName
                    'Try
                    '    variantName = .variantName.Trim
                    'Catch ex As Exception
                    '    variantName = ""
                    'End Try

                End With

                If Not projekthistorie Is Nothing Then
                    If projekthistorie.Count > 0 Then
                        vglName = projekthistorie.First.getShapeText
                    End If

                Else
                    projekthistorie = New clsProjektHistorie
                End If

                If vglName <> hproj.getShapeText Then

                    If request.pingMongoDb() Then
                        ' projekthistorie muss nur dann neu geladen werden, wenn sie nicht bereits für dieses Projekt geholt wurde
                        projekthistorie.liste = request.retrieveProjectHistoryFromDB(projectname:=pName, variantName:=variantName, _
                                                                            storedEarliest:=StartofCalendar, storedLatest:=Date.Now)
                        If projekthistorie.Count <> 0 Then

                            projekthistorie.Add(Date.Now, hproj)

                        End If

                    Else
                        Call MsgBox("Datenbank-Verbindung ist unterbrochen")
                        projekthistorie.clear()
                    End If

                Else
                    ' der aktuelle Stand hproj muss hinzugefügt werden 
                    Dim lastElem As Integer = projekthistorie.Count - 1
                    projekthistorie.RemoveAt(lastElem)
                    projekthistorie.Add(Date.Now, hproj)
                End If


                Dim nrSnapshots As Integer = projekthistorie.Count

                If nrSnapshots > 0 Then

                    With showCharacteristics

                        .Text = "Historie für Projekt " & pName.Trim & vbLf & _
                                "( " & projekthistorie.getZeitraum & " )"
                        .timeSlider.Minimum = 0
                        .timeSlider.Maximum = nrSnapshots - 1

                        '.ampelErlaeuterung.Text = kvp.Value.ampelErlaeuterung

                        'If kvp.Value.ampelStatus = 1 Then
                        '    .ampelPicture.LoadAsync(grueneAmpel)
                        'ElseIf kvp.Value.ampelStatus = 2 Then
                        '    .ampelPicture.LoadAsync(gelbeAmpel)
                        'ElseIf kvp.Value.ampelStatus = 3 Then
                        '    .ampelPicture.LoadAsync(roteAmpel)
                        'Else
                        '    .ampelPicture.LoadAsync(graueAmpel)
                        'End If

                        '.snapshotDate.Text = kvp.Value.timeStamp.ToString
                        ' das ist ja der aktuelle Stand ..
                        .snapshotDate.Text = "Aktueller Stand"
                        ' Designer 
                        'Dim zE As String = "(" & awinSettings.zeitEinheit & ")"
                        '.engpass1.Text = "Designer:          " & kvp.Value.getRessourcenBedarf(3).Sum.ToString("###.#") & zE
                        '.engpass2.Text = "Personalkosten: " & kvp.Value.getAllPersonalKosten.Sum.ToString("###.#") & " (T€)"
                        '.engpass3.Text = "Sonstige Kosten:   " & kvp.Value.getGesamtAndereKosten.Sum.ToString("###.#") & " (T€)"


                    End With


                    ' jetzt wird das Form aufgerufen ... 

                    'returnValue = showCharacteristics.ShowDialog
                    showCharacteristics.Show()

                Else
                    Call MsgBox("es gibt noch keine Planungs-Historie")
                End If

            Else
                Call MsgBox("bitte nur ein Projekt selektieren")
                'For Each singleShp In awinSelection
                '    With singleShp
                '        If .AutoShapeType = MsoAutoShapeType.msoShapeRoundedRectangle Then
                '            nrSelPshp = nrSelPshp + 1
                '            SID = .ID.ToString
                '        End If
                '    End With
                'Next
            End If
        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If

        enableOnUpdate = True


    End Sub
End Module
