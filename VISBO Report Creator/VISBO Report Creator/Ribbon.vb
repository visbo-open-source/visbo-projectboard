'TODO:  Führen Sie diese Schritte aus, um das Element auf dem Menüband (XML) zu aktivieren:

'1: Kopieren Sie folgenden Codeblock in die ThisAddin-, ThisWorkbook- oder ThisDocument-Klasse.

'Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
'    Return New Ribbon()
'End Function

'2. Erstellen Sie Rückrufmethoden im Abschnitt "Menübandrückrufe" dieser Klasse, um Benutzeraktionen
'   zu behandeln, zum Beispiel das Klicken auf eine Schaltfläche. Hinweis: Wenn Sie dieses Menüband aus dem
'   Menüband-Designer exportiert haben, verschieben Sie den Code aus den Ereignishandlern in die Rückrufmethoden, und
'   ändern Sie den Code für die Verwendung mit dem Programmiermodell für die Menübanderweiterung (RibbonX).

'3. Weisen Sie den Steuerelementtags in der Menüband-XML-Datei Attribute zu, um die entsprechenden Rückrufmethoden im Code anzugeben.

'Weitere Informationen erhalten Sie in der Menüband-XML-Dokumentation in der Hilfe zu Visual Studio-Tools für Office.

Imports ProjectBoardDefinitions
Imports ProjectBoardBasic
Imports System.Windows.Forms

<Runtime.InteropServices.ComVisible(True)> _
Public Class Ribbon
    Implements Office.IRibbonExtensibility

    Private ribbon As Office.IRibbonUI

    Public Sub New()
    End Sub

    Public Function GetCustomUI(ByVal ribbonID As String) As String Implements Office.IRibbonExtensibility.GetCustomUI
        Return GetResourceText("VISBO_Report_Creator.Ribbon.xml")
    End Function

#Region "Menübandrückrufe"
    'Erstellen Sie hier Rückrufmethoden. Weitere Informationen zum Hinzufügen von Rückrufmethoden finden Sie unter https://go.microsoft.com/fwlink/?LinkID=271226.
    Public Sub Ribbon_Load(ByVal ribbonUI As Office.IRibbonUI)
        Me.ribbon = ribbonUI
        Me.ribbon.Invalidate()
    End Sub

    ''' <summary>
    ''' kann sowohl Projekt wie auch Portfolio Reports erstellen 
    ''' </summary>
    ''' <param name="control"></param>
    Sub createReports(control As Microsoft.Office.Core.IRibbonControl)

        Dim returnValue As Windows.Forms.DialogResult
        Dim errMsg As String = ""

        Dim loadProjectsForm As New frmProjPortfolioAdmin
        Dim weitermachen As Boolean = True
        If noDBLoginPPT Then
            ' einloggen, dann Visbo Center wählen, dann Orga einlesen, dann user roles, dann customization und appearance classes ... 
            weitermachen = successfulLoginAndSetup(errMsg)
        End If

        If weitermachen Then
            ' jetzt hat ja alles geklappt: login, Settings lesen, ... 

            noDBLoginPPT = False
            Try

                With loadProjectsForm

                    .aKtionskennung = PTTvActions.loadPVInPPT

                    '' '' ''.portfolioName.Visible = False
                    '' '' ''.Label1.Visible = False
                End With

                returnValue = loadProjectsForm.ShowDialog

                If returnValue = Windows.Forms.DialogResult.OK Then

                    ' tk 7.10.19 jetzt werden die Platzhalter umgewandelt ...
                    Dim hproj As clsProjekt = Nothing
                    Dim anzP As Integer = ShowProjekte.Count
                    If selectedProjekte.Count = 1 Then
                        hproj = selectedProjekte.getProject(1)

                        Dim tmpCollection As New Collection
                        Call createPPTSlidesFromProjectWithinPPT(hproj, tmpCollection, tmpCollection, tmpCollection, tmpCollection, tmpCollection, tmpCollection, 0.0, 12.0)
                        ' tk 7.10 selectedProjekte wieder zurücksetzen ..
                        ShowProjekte.Clear(False)
                        selectedProjekte.Clear(False)
                        showRangeLeft = 0
                        showRangeRight = 0


                        Try
                            ' jetzt den Namen auf das Projekt setzen, wenn er nicht schon vorher gesetzt wurde .. 

                            Dim savePath As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments
                            Dim fullFileName As String = My.Computer.FileSystem.CombinePath(savePath, hproj.name)
                            If anzP > 1 Then
                                fullFileName = My.Computer.FileSystem.CombinePath(savePath, "Multiprojekt-Report")
                            End If
                            curPresentation.SaveAs(fullFileName)
                        Catch ex As Exception

                        End Try


                    Else
                        Dim msgtxt As String = "kein Projekt ausgewählt ... Abbruch"
                        If awinSettings.englishLanguage Then
                            msgtxt = "no project selected ... Exit"
                        End If
                        Call MsgBox(msgtxt)
                    End If



                Else
                    ' returnValue = DialogResult.Cancel

                End If

            Catch ex As Exception

                Call MsgBox(ex.Message)
            End Try
        Else
            Call MsgBox("Cancelled ...")
        End If

    End Sub

    Sub createPortfolioReports(control As Microsoft.Office.Core.IRibbonControl)
        Call MsgBox("creates Portfolio report ...")
    End Sub

    Private Function successfulLoginAndSetup(ByRef errMsg As String) As Boolean
        Dim wasSuccessful As Boolean = False
        Dim err As New clsErrorCodeMsg
        awinSettings.databaseURL = My.Settings.dbURL
        awinSettings.databaseName = My.Settings.dbName
        awinSettings.visboServer = True
        awinSettings.proxyURL = My.Settings.proxyURL
        awinSettings.DBWithSSL = My.Settings.mongoDBSSL
        awinSettings.rememberUserPwd = My.Settings.rememberUserPWD
        awinSettings.userNamePWD = My.Settings.userNamePWD

        If logInToMongoDB(True) Then
            ' weitermachen ...

            Try
                ' die dem User zugeodneten Visbo Center lesen ...
                ' jetzt muss geprüft werden, ob es mehr als ein zugelassenes VISBO Center gibt , ist dann der Fall wenn es ein # im awinsettings.databaseNAme gibt 
                Dim listOfVCs As List(Of String) = CType(databaseAcc, DBAccLayer.Request).retrieveVCsForUser(err)

                If listOfVCs.Count > 1 Then
                    Dim chooseVC As New frmSelectOneItem
                    chooseVC.itemsCollection = listOfVCs
                    If chooseVC.ShowDialog = DialogResult.OK Then
                        ' alles ok 
                        awinSettings.databaseName = chooseVC.itemList.SelectedItem.ToString
                        Dim changeOK As Boolean = CType(databaseAcc, DBAccLayer.Request).updateActualVC(awinSettings.databaseName, err)

                        If Not changeOK Then
                            Throw New ArgumentException("bad Selection of VISBO project Center ... program ends  ...")
                        End If
                    Else
                        Throw New ArgumentException("no Selection of VISBO project Center ... program ends  ...")
                    End If

                End If

                ' lesen der Customization und Appearance Classes; hier wird der SOC , der StartOfCalendar gesetzt ...  

                appearanceDefinitions = CType(databaseAcc, DBAccLayer.Request).retrieveAppearancesFromDB("", Date.Now, False, err)
                If IsNothing(appearanceDefinitions) Then
                    Throw New ArgumentException("Appearance classes do not exist")
                End If

                Dim customizations As clsCustomization = CType(databaseAcc, DBAccLayer.Request).retrieveCustomizationFromDB("", Date.Now, False, err)
                If IsNothing(customizations) Then
                    Throw New ArgumentException("Customization does not exist")
                Else
                    ' alle awinSettings... mit den customizations... besetzen
                    'For Each kvp As KeyValuePair(Of Integer, clsBusinessUnit) In businessUnitDefinitions
                    '    customizations.businessUnitDefinitions.Add(kvp.Key, kvp.Value)
                    'Next
                    businessUnitDefinitions = customizations.businessUnitDefinitions

                    'For Each kvp As KeyValuePair(Of String, clsPhasenDefinition) In PhaseDefinitions.liste
                    '    customizations.phaseDefinitions.Add(kvp.Value)
                    'Next
                    PhaseDefinitions = customizations.phaseDefinitions

                    'For Each kvp As KeyValuePair(Of String, clsMeilensteinDefinition) In MilestoneDefinitions.liste
                    '    customizations.milestoneDefinitions.Add(kvp.Value)
                    'Next
                    MilestoneDefinitions = customizations.milestoneDefinitions
                    ' die Struktur clsCustomization besetzen und in die DB dieses VCs eintragen

                    showtimezone_color = customizations.showtimezone_color
                    noshowtimezone_color = customizations.noshowtimezone_color
                    calendarFontColor = customizations.calendarFontColor
                    nrOfDaysMonth = customizations.nrOfDaysMonth
                    farbeInternOP = customizations.farbeInternOP
                    farbeExterne = customizations.farbeExterne
                    iProjektFarbe = customizations.iProjektFarbe
                    iWertFarbe = customizations.iWertFarbe
                    vergleichsfarbe0 = customizations.vergleichsfarbe0
                    vergleichsfarbe1 = customizations.vergleichsfarbe1
                    'customizations.vergleichsfarbe2 = vergleichsfarbe2

                    awinSettings.SollIstFarbeB = customizations.SollIstFarbeB
                    awinSettings.SollIstFarbeL = customizations.SollIstFarbeL
                    awinSettings.SollIstFarbeC = customizations.SollIstFarbeC
                    awinSettings.AmpelGruen = customizations.AmpelGruen
                    'tmpcolor = CType(.Range("AmpelGruen").Interior.Color, Microsoft.Office.Interop.Excel.ColorFormat)
                    awinSettings.AmpelGelb = customizations.AmpelGelb
                    awinSettings.AmpelRot = customizations.AmpelRot
                    awinSettings.AmpelNichtBewertet = customizations.AmpelNichtBewertet
                    awinSettings.glowColor = customizations.glowColor

                    awinSettings.timeSpanColor = customizations.timeSpanColor
                    awinSettings.showTimeSpanInPT = customizations.showTimeSpanInPT

                    awinSettings.gridLineColor = customizations.gridLineColor

                    awinSettings.missingDefinitionColor = customizations.missingDefinitionColor

                    awinSettings.allianzIstDatenReferate = customizations.allianzIstDatenReferate

                    awinSettings.autoSetActualDataDate = customizations.autoSetActualDataDate

                    awinSettings.actualDataMonth = customizations.actualDataMonth
                    ergebnisfarbe1 = customizations.ergebnisfarbe1
                    ergebnisfarbe2 = customizations.ergebnisfarbe2
                    weightStrategicFit = customizations.weightStrategicFit
                    awinSettings.kalenderStart = customizations.kalenderStart
                    awinSettings.zeitEinheit = customizations.zeitEinheit
                    awinSettings.kapaEinheit = customizations.kapaEinheit
                    awinSettings.offsetEinheit = customizations.offsetEinheit
                    awinSettings.EinzelRessExport = customizations.EinzelRessExport
                    awinSettings.zeilenhoehe1 = customizations.zeilenhoehe1
                    awinSettings.zeilenhoehe2 = customizations.zeilenhoehe2
                    awinSettings.spaltenbreite = customizations.spaltenbreite
                    awinSettings.autoCorrectBedarfe = customizations.autoCorrectBedarfe
                    awinSettings.propAnpassRess = customizations.propAnpassRess
                    awinSettings.showValuesOfSelected = customizations.showValuesOfSelected

                    awinSettings.mppProjectsWithNoMPmayPass = customizations.mppProjectsWithNoMPmayPass
                    awinSettings.fullProtocol = customizations.fullProtocol
                    awinSettings.addMissingPhaseMilestoneDef = customizations.addMissingPhaseMilestoneDef
                    awinSettings.alwaysAcceptTemplateNames = customizations.alwaysAcceptTemplateNames
                    awinSettings.eliminateDuplicates = customizations.eliminateDuplicates
                    awinSettings.importUnknownNames = customizations.importUnknownNames
                    awinSettings.createUniqueSiblingNames = customizations.createUniqueSiblingNames

                    awinSettings.readWriteMissingDefinitions = customizations.readWriteMissingDefinitions
                    awinSettings.meExtendedColumnsView = customizations.meExtendedColumnsView
                    awinSettings.meDontAskWhenAutoReduce = customizations.meDontAskWhenAutoReduce
                    awinSettings.readCostRolesFromDB = customizations.readCostRolesFromDB

                    awinSettings.importTyp = customizations.importTyp

                    awinSettings.meAuslastungIsInclExt = customizations.meAuslastungIsInclExt

                    awinSettings.englishLanguage = customizations.englishLanguage

                    awinSettings.showPlaceholderAndAssigned = customizations.showPlaceholderAndAssigned
                    awinSettings.considerRiskFee = customizations.considerRiskFee

                    ' noch zu tun, sonst in readOtherdefinitions
                    StartofCalendar = awinSettings.kalenderStart
                    'StartofCalendar = StartofCalendar.ToLocalTime()

                    historicDate = StartofCalendar
                    Try
                        If awinSettings.englishLanguage Then
                            menuCult = ReportLang(PTSprache.englisch)
                            repCult = menuCult
                            awinSettings.kapaEinheit = "PD"
                        Else
                            awinSettings.kapaEinheit = "PT"
                            menuCult = ReportLang(PTSprache.deutsch)
                            repCult = menuCult
                        End If
                    Catch ex As Exception
                        awinSettings.englishLanguage = False
                        awinSettings.kapaEinheit = "PT"
                        menuCult = ReportLang(PTSprache.deutsch)
                        repCult = menuCult
                    End Try
                End If

                ' Lesen der CustomField-Definitions
                ' Auslesen der Custom Field Definitions aus den VCSettings über ReST-Server

                customFieldDefinitions = CType(databaseAcc, DBAccLayer.Request).retrieveCustomFieldsFromDB(err)

                If IsNothing(customFieldDefinitions) Then
                    customFieldDefinitions = New clsCustomFieldDefinitions
                    'Call MsgBox("no Custom-Field-Definitions in database")
                End If


                ' lesen der Organisation und Kapazitäten
                Dim currentOrga As clsOrganisation = CType(databaseAcc, DBAccLayer.Request).retrieveOrganisationFromDB("", Date.Now, False, err)
                If IsNothing(currentOrga) Then

                ElseIf currentOrga.count > 0 Then
                    validOrganisations.addOrga(currentOrga)
                    CostDefinitions = currentOrga.allCosts
                    RoleDefinitions = currentOrga.allRoles
                Else
                    RoleDefinitions = New clsRollen
                    CostDefinitions = New clsKostenarten
                End If

                ' lesen der Custom User Roles 
                Dim meldungen As New Collection
                Try

                    Call setUserRoles(meldungen)
                Catch ex As Exception
                    ' hier bekommt der Nutzer die Rolle Projektleiter 
                    myCustomUserRole = New clsCustomUserRole

                    With myCustomUserRole
                        .customUserRole = ptCustomUserRoles.ProjektLeitung
                        .specifics = ""
                        .userName = dbUsername
                    End With
                    ' jetzt gibt es eine currentUserRole: myCustomUserRole - die gelten aktuell nur für Excel Projectboard, haben aber keine auswirkungen auf PPT Report Creation Addin
                    Call myCustomUserRole.setNonAllowances()
                End Try


                wasSuccessful = True
            Catch ex As Exception
                wasSuccessful = False
                errMsg = ex.Message
            End Try

        Else
            wasSuccessful = False
        End If

        successfulLoginAndSetup = wasSuccessful
    End Function

    Sub protectReport(control As Microsoft.Office.Core.IRibbonControl)

    End Sub

    Sub createLegend(control As Microsoft.Office.Core.IRibbonControl)

    End Sub

    Sub cpsettings(control As Microsoft.Office.Core.IRibbonControl)
        Dim mppFrm As New frmMppSettings
        Dim dialogreturn As DialogResult
        Dim calledFrom As String = "Powerpoint"

        If calledFrom = "MS-Project" Then
            mppFrm.calledfrom = calledFrom
        Else
            mppFrm.calledfrom = "frmSelectPPTTempl"
        End If

        dialogreturn = mppFrm.ShowDialog
    End Sub

#End Region

#Region "Hilfsprogramme"

    Private Shared Function GetResourceText(ByVal resourceName As String) As String
        Dim asm As Reflection.Assembly = Reflection.Assembly.GetExecutingAssembly()
        Dim resourceNames() As String = asm.GetManifestResourceNames()
        For i As Integer = 0 To resourceNames.Length - 1
            If String.Compare(resourceName, resourceNames(i), StringComparison.OrdinalIgnoreCase) = 0 Then
                Using resourceReader As IO.StreamReader = New IO.StreamReader(asm.GetManifestResourceStream(resourceNames(i)))
                    If resourceReader IsNot Nothing Then
                        Return resourceReader.ReadToEnd()
                    End If
                End Using
            End If
        Next
        Return Nothing
    End Function

#End Region

End Class
