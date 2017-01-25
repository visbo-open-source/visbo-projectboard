
Imports ProjectBoardDefinitions
Imports ProjectBoardBasic
Imports MongoDbAccess
Imports ClassLibrary1
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

                .menuOption = PTmenue.filterdefinieren
                returnValue = .ShowDialog

            End With

        ElseIf controlID = "Pt6G3M1B2" Then

            awinSettings.useHierarchy = True

            With hryFormular

                .menuOption = PTmenue.filterdefinieren
                returnValue = .ShowDialog

            End With


        ElseIf ShowProjekte.Count > 0 Then

            If awinSettings.isHryNameFrmActive Then
                Call MsgBox("es kann nur ein Fenster zur Hierarchie- bzw. Namenauswahl geöffnet sein ...")

            ElseIf controlID = "PTXG1B4" Or controlID = "PT0G1B8" Then
                ' Namen auswählen, Visualisieren
                awinSettings.useHierarchy = False
                With nameFormular
                    
                    .menuOption = PTmenue.visualisieren
                    ' Nicht Modal anzeigen
                    .Show()
                    'returnValue = .ShowDialog

                End With

            ElseIf controlID = "PTXG1B5" Or controlID = "PT0G1B9" Then
                ' Hierarchie auswählen, visualisieren
                awinSettings.useHierarchy = True

                With hryFormular
                    
                    .menuOption = PTmenue.visualisieren
                    ' Nicht Modal anzeigen
                    .Show()
                    'returnValue = .ShowDialog

                End With
            ElseIf controlID = "PTXG1B6" Or controlID = "PTMEC1" Then
                ' Namen auswählen, Leistbarkeit

                awinSettings.useHierarchy = False
                With nameFormular

                    .ribbonButtonID = controlID
                    .menuOption = PTmenue.leistbarkeitsAnalyse
                    ' Nicht Modal anzeigen
                    .Show()
                    'returnValue = .ShowDialog

                End With

            ElseIf controlID = "PTXG1B7" Then
                ' Hierarchie auswählen, Leistbarkeit
                awinSettings.useHierarchy = True
                With hryFormular
                    
                    .menuOption = PTmenue.leistbarkeitsAnalyse
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

                        
                        .menuOption = PTmenue.einzelprojektReport
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
                        
                        .menuOption = PTmenue.einzelprojektReport
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

                        .menuOption = PTmenue.multiprojektReport
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

                        .menuOption = PTmenue.multiprojektReport
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

                    
                    .menuOption = PTmenue.excelExport
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

                    .menuOption = PTmenue.excelExport
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

                    
                    .menuOption = PTmenue.vorlageErstellen
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

                    .menuOption = PTmenue.vorlageErstellen
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

                        .menuOption = PTmenue.meilensteinTrendanalyse
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
        If ShowProjekte.Count > 0 And showRangeRight - showRangeLeft >= minColumns - 1 Then

            If ControlID = "PTXG1B6" Then
                ' Auswahl über Namen

                With namensFormular

                    .ribbonButtonID = ControlID
                    .menuOption = PTmenue.leistbarkeitsAnalyse
                    '.Show()
                    returnValue = .ShowDialog

                End With


            Else
                ' Auswahl über Hierarchie
                ' Hierarchie
                awinSettings.useHierarchy = True
                With hierarchieFormular

                    .menuOption = PTmenue.leistbarkeitsAnalyse
                    '.Show()
                    returnValue = .ShowDialog

                End With

            End If

        ElseIf ShowProjekte.Count = 0 Then

            Call MsgBox("Es sind keine Projekte geladen!  ")

        ElseIf showRangeRight - showRangeLeft < minColumns - 1 Then

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

        Dim variantDescription As String = ""


        Call projektTafelInit()

        enableOnUpdate = False

        If control.Id = "PT2G1M1B0" Then
            ' neue Variante anlegen 
            Try
                awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
            Catch ex As Exception
                awinSelection = Nothing
            End Try

            If Not awinSelection Is Nothing Then

                For i As Integer = 1 To awinSelection.Count
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
                        .txtDescription.Text = variantDescription
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

                        With neueVariante
                            neuerVariantenName = .newVariant.Text
                            variantDescription = .txtDescription.Text
                        End With


                        With newproj
                            .name = hproj.name
                            .variantName = neuerVariantenName
                            .variantDescription = variantDescription
                            .ampelErlaeuterung = hproj.ampelErlaeuterung
                            .ampelStatus = hproj.ampelStatus
                            .timeStamp = Date.Now
                            .shpUID = hproj.shpUID
                            .tfZeile = hproj.tfZeile
                            .Status = ProjektStatus(0)

                        End With

                        If Not currentConstellationName.EndsWith("(*)") Then
                            currentConstellationName = currentConstellationName & " (*)"
                        End If

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

        Else
            ' nur Varianten Erläuterung editieren ... 

        End If

        Call storeSessionConstellation("Last")

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

                .aKtionskennung = PTTvActions.deleteV

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
        Dim outputCollection As New Collection
        Call projektTafelInit()

        appInstance.EnableEvents = False
        enableOnUpdate = False

        Try
            'awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not IsNothing(awinSelection) Then

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
                            Dim hproj As clsProjekt = ShowProjekte.getProject(.Name)
                            If Not IsNothing(hproj) Then
                                If notReferencedByAnyPortfolio(hproj.name, hproj.variantName) Then
                                    Call awinDeleteProjectInSession(pName:=.Name)
                                Else
                                    Dim outputline As String = "Löschen verweigert " & hproj.name & " wird in Szenarien referenziert: "
                                    outputline = outputline & projectConstellations.getSzenarioNamesWith(hproj.name, hproj.variantName)
                                    outputCollection.Add(outputline)
                                End If
                            End If


                        Catch ex As Exception
                            Exit For
                        End Try

                    End If
                End With


            Next

            If Not currentConstellationName.EndsWith("(*)") Then
                currentConstellationName = currentConstellationName & " (*)"
            End If

            ' ein oder mehrere Projekte wurden gelöscht  - typus = 3
            Call awinNeuZeichnenDiagramme(3)

            If outputCollection.Count > 0 Then
                Call showOutPut(outputCollection, "Löschen von Projekten", "folgende Fehler sind aufgetreten:")
            End If

        Else

            Dim deletedProj As Integer = 0

            If AlleProjekte.Count = 0 Then
                Call MsgBox("es sind keine Projekte geladen !")
            Else

                'Dim deleteProjects As New frmDeleteProjects
                Dim deleteProjects As New frmProjPortfolioAdmin
                Try

                    With deleteProjects

                        .aKtionskennung = PTTvActions.delFromSession

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

        Call storeSessionConstellation("Last")

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

                .aKtionskennung = PTTvActions.loadPV

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
    ''' alle aktuell in AlleProjekte geladenen PRojekte und Varianten werden angezeigt und können 
    ''' aktiv / de-aktiv gesetzt werden 
    ''' on-the-fly werden evtl gezeigte Portfolio Charts aktualisiert
    ''' erst mit OK werden die Projekte gezeichnet und als lastConstellation gespeichert , oder unter dem angegebenen Namen
    ''' </summary>
    ''' <remarks></remarks>
    Sub PBBChangeCurrentPortfolio()


        'Dim returnValue As DialogResult

        'Dim deleteProjects As New frmDeleteProjects
        Dim changePortfolio As New frmProjPortfolioAdmin


        If AlleProjekte.Count > 0 Then
            ' das letzte Portfolio speichern 
            Call storeSessionConstellation("Last")

            Try

                With changePortfolio

                    .aKtionskennung = PTTvActions.chgInSession

                End With

                'Call awinClearPlanTafel()

                changePortfolio.Show()

                'returnValue = changePortfolio.ShowDialog

                '' die Operation ist bereits ausgeführt - deswegen muss hier nichts mehr unterschieden werden 

                'If returnValue = DialogResult.OK Then
                '    ' das aktuelle Portfolio speichern 

                '    ' dann die Projekt-Tafel neu zeichnen 

                'Else
                '    ' das last-Portfolio wiederherstellen 
                '    Call loadSessionConstellation("Last", False, False, False)

                '    ' gezeichnet werden muss nix ... 

                'End If

            Catch ex As Exception

                Call MsgBox(ex.Message)
            End Try
        Else

            Call MsgBox("keine Projekte geladen ...")
        End If
        


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

                If control.Id = "Pt5G3B4" Then
                    .aKtionskennung = PTTvActions.delAllExceptFromDB
                ElseIf control.Id = "Pt5G3B3" Then
                    .aKtionskennung = PTTvActions.delFromDB
                End If

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

                .aKtionskennung = PTTvActions.activateV

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
