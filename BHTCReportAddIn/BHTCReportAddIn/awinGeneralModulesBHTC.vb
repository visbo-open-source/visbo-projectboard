
Option Explicit On
'Option Strict On

Imports ProjectBoardDefinitions
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.ComponentModel
Imports System.Windows
Imports System.Windows.Forms

Imports System
Imports System.Runtime.Serialization
Imports System.Xml
Imports System.Xml.Serialization
Imports System.IO
Imports Microsoft.VisualBasic
Imports ProjectBoardBasic
Imports System.Security.Principal



Module awinGeneralModulesBHTC

    Private Enum ptInventurSpalten
        Name = 0
        Vorlage = 1
        Start = 2
        Ende = 3
        startElement = 4
        endElement = 5
        Dauer = 6
        Budget = 7
        Risiko = 8
        Strategie = 9
        Volumen = 10
        Komplexitaet = 11
        Businessunit = 12
        Beschreibung = 13
    End Enum

    Private Enum ptModuleSpalten
        produktlinie = 0
        name = 1
        projektTyp = 2
        abhaengigVon = 3
        strategicFit = 4
        risiko = 5
        volume = 6
        budget = 7
    End Enum

 

    ''' <summary>
    ''' liest das Customization File aus und initialisiert die globalen Variablen entsprechend
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub awinsetTypenNEW(ByVal special As String)

        Try

            Dim xlsCustomization As Excel.Workbook = Nothing

            ReDim importOrdnerNames(6)
            ReDim exportOrdnerNames(4)


            ' Auslesen des Window Namens 
            Dim accountToken As IntPtr = WindowsIdentity.GetCurrent().Token
            Dim myUser As New WindowsIdentity(accountToken)
            myWindowsName = myUser.Name
            ''Call logfileSchreiben("Windows-User: ", myWindowsName, anzFehler)

            ' hier werden die Ordner Namen für den Import wie Export festgelegt ... 
            'awinPath = appInstance.ActiveWorkbook.Path & "\"

            globalPath = awinSettings.globalPath

            ' awinpath kann relativ oder absolut angegeben sein, beides möglich
            Dim curUserDir As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments
            awinPath = My.Computer.FileSystem.CombinePath(curUserDir, awinSettings.awinPath)
            If Not awinPath.EndsWith("\") Then
                awinPath = awinPath & "\"
            End If


            If awinPath = "" And (globalPath <> "" And My.Computer.FileSystem.DirectoryExists(globalPath)) Then
                awinPath = globalPath
            ElseIf globalPath = "" And (awinPath <> "" And My.Computer.FileSystem.DirectoryExists(awinPath)) Then
                globalPath = awinPath
            ElseIf globalPath = "" Or awinPath = "" Then
                Throw New ArgumentException("Globaler Ordner " & awinSettings.globalPath & " und Lokaler Ordner " & awinSettings.awinPath & " existieren nicht")
            End If

            If My.Computer.FileSystem.DirectoryExists(globalPath) And (Dir(globalPath, vbDirectory) = "") Then
                Throw New ArgumentException("Requirementsordner " & awinSettings.globalPath & " existiert nicht")
            End If

            If awinPath <> globalPath And My.Computer.FileSystem.DirectoryExists(globalPath) Then
                Call synchronizeGlobalToLocalFolder()
            Else
                If My.Computer.FileSystem.DirectoryExists(awinPath) And (Dir(awinPath, vbDirectory) = "") Then
                    Throw New ArgumentException("Requirementsordner " & awinSettings.awinPath & " existiert nicht")
                End If

            End If

            ' Benutzer arbeitet auf dem awinPath-Directories ohne Synchronisation

            ' Erzeugen des Report Ordners, wenn er nicht schon existiert .. 
            reportOrdnerName = awinPath & "Reports\"
            Try
                My.Computer.FileSystem.CreateDirectory(reportOrdnerName)
            Catch ex As Exception

            End Try

            importOrdnerNames(PTImpExp.visbo) = awinPath & "Import\VISBO Steckbriefe"
            importOrdnerNames(PTImpExp.rplan) = awinPath & "Import\RPLAN-Excel"
            importOrdnerNames(PTImpExp.msproject) = awinPath & "Import\MSProject"
            importOrdnerNames(PTImpExp.simpleScen) = awinPath & "Import\einfache Szenarien"
            importOrdnerNames(PTImpExp.modulScen) = awinPath & "Import\modulare Szenarien"
            importOrdnerNames(PTImpExp.addElements) = awinPath & "Import\addOn Regeln"
            importOrdnerNames(PTImpExp.rplanrxf) = awinPath & "Import\RXF Files"

            exportOrdnerNames(PTImpExp.visbo) = awinPath & "Export\VISBO Steckbriefe"
            exportOrdnerNames(PTImpExp.rplan) = awinPath & "Export\RPLAN-Excel"
            exportOrdnerNames(PTImpExp.msproject) = awinPath & "Export\MSProject"
            exportOrdnerNames(PTImpExp.simpleScen) = awinPath & "Export\einfache Szenarien"
            exportOrdnerNames(PTImpExp.modulScen) = awinPath & "Export\modulare Szenarien"



            StartofCalendar = StartofCalendar.Date

            LizenzKomponenten(PTSWKomp.ProjectAdmin) = "ProjectAdmin"
            LizenzKomponenten(PTSWKomp.Swimlanes2) = "Swimlanes2"
            LizenzKomponenten(PTSWKomp.SWkomp2) = "SWkomp2"
            LizenzKomponenten(PTSWKomp.SWkomp3) = "SWkomp3"
            LizenzKomponenten(PTSWKomp.SWkomp4) = "SWkomp4"

            ProjektStatus(0) = "geplant"
            ProjektStatus(1) = "beauftragt"
            ProjektStatus(2) = "beauftragt, Änderung noch nicht freigegeben"
            ProjektStatus(3) = "beendet" ' ein Projekt wurde in seinem Verlauf beendet, ohne es plangemäß abzuschliessen
            ProjektStatus(4) = "abgeschlossen"


            DiagrammTypen(0) = "Phase"
            DiagrammTypen(1) = "Rolle"
            DiagrammTypen(2) = "Kostenart"
            DiagrammTypen(3) = "Portfolio"
            DiagrammTypen(4) = "Ergebnis"
            DiagrammTypen(5) = "Meilenstein"
            DiagrammTypen(6) = "Meilenstein Trendanalyse"

            ergebnisChartName(0) = "Earned Value"
            ergebnisChartName(1) = "Earned Value - gewichtet"
            ergebnisChartName(2) = "Verbesserungs-Potential"
            ergebnisChartName(3) = "Risiko-Abschlag"

            ReDim portfolioDiagrammtitel(21)
            portfolioDiagrammtitel(PTpfdk.Phasen) = "Phasen - Übersicht"
            portfolioDiagrammtitel(PTpfdk.Rollen) = "Rollen - Übersicht"
            portfolioDiagrammtitel(PTpfdk.Kosten) = "Kosten - Übersicht"
            portfolioDiagrammtitel(PTpfdk.ErgebnisWasserfall) = summentitel1
            portfolioDiagrammtitel(PTpfdk.FitRisiko) = summentitel2
            portfolioDiagrammtitel(PTpfdk.Auslastung) = summentitel9
            portfolioDiagrammtitel(PTpfdk.UeberAuslastung) = summentitel10
            portfolioDiagrammtitel(PTpfdk.Unterauslastung) = summentitel11
            portfolioDiagrammtitel(PTpfdk.ZieleV) = summentitel6
            portfolioDiagrammtitel(PTpfdk.ZieleF) = summentitel7
            portfolioDiagrammtitel(PTpfdk.ComplexRisiko) = "Komplexität, Risiko und Volumen"
            portfolioDiagrammtitel(PTpfdk.ZeitRisiko) = "Zeit, Risiko und Volumen"
            portfolioDiagrammtitel(PTpfdk.AmpelFarbe) = ""
            portfolioDiagrammtitel(PTpfdk.ProjektFarbe) = ""
            portfolioDiagrammtitel(PTpfdk.Meilenstein) = "Meilenstein - Übersicht"
            portfolioDiagrammtitel(PTpfdk.FitRisikoVol) = "strategischer Fit, Risiko & Volumen"
            portfolioDiagrammtitel(PTpfdk.Dependencies) = "Abhängigkeiten: Aktive bzw passive Beeinflussung"
            portfolioDiagrammtitel(PTpfdk.betterWorseL) = "Abweichungen zum letztem Stand"
            portfolioDiagrammtitel(PTpfdk.betterWorseB) = "Abweichungen zur Beauftragung"
            portfolioDiagrammtitel(PTpfdk.Budget) = "Budget Übersicht"
            portfolioDiagrammtitel(PTpfdk.FitRisikoDependency) = "strategischer Fit, Risiko & Ausstrahlung"


            autoSzenarioNamen(0) = "vor Optimierung"
            autoSzenarioNamen(1) = "1. Optimum"
            autoSzenarioNamen(2) = "2. Optimum"
            autoSzenarioNamen(3) = "3. Optimum"

            windowNames(0) = "Cockpit Phasen"
            windowNames(1) = "Cockpit Rollen"
            windowNames(2) = "Cockpit Kosten"
            windowNames(3) = "Cockpit Wertigkeit"
            windowNames(4) = "Cockpit Ergebnisse"
            windowNames(5) = "Projekt Tafel"

            '
            ' die Namen der Worksheets Ressourcen und Portfolio verfügbar machen
            '
            arrWsNames(1) = "Portfolio"
            arrWsNames(2) = "Vorlage"                          ' dient als Hilfs-Sheet für Anzeige in Plantafel 
            arrWsNames(3) = "Tabelle1"
            arrWsNames(4) = "Einstellungen"
            arrWsNames(5) = "Tabelle2"
            arrWsNames(6) = "Edit Allgemein"
            arrWsNames(7) = "Darstellungsklassen"                          ' war Kosten ; ist nicht mehr notwendig
            arrWsNames(8) = "Phasen-Mappings"
            arrWsNames(9) = "Tabelle3"
            arrWsNames(10) = "Meilenstein-Mappings"
            arrWsNames(11) = "Projekt editieren"
            arrWsNames(12) = "Projektdefinition Erloese"
            arrWsNames(13) = "Projekt iErloese"
            arrWsNames(14) = "Objekte"
            arrWsNames(15) = "Portfolio Vorlage"


            awinSettings.applyFilter = False

            showRangeLeft = 0
            showRangeRight = 0

            'selectedRoleNeeds = 0
            'selectedCostNeeds = 0

            ' '' '' bestimmen der maximalen Breite und Höhe 
            '' ''Dim formerSU As Boolean = appInstance.ScreenUpdating
            '' ''appInstance.ScreenUpdating = False


            ' '' '' um dahinter temporär die Darstellungsklassen kopieren zu können  
            '' ''Dim projectBoardSheet As Excel.Worksheet = CType(appInstance.ActiveSheet, _
            '' ''                                        Global.Microsoft.Office.Interop.Excel.Worksheet)



            '' ''With appInstance.ActiveWindow


            '' ''    If .WindowState = Excel.XlWindowState.xlMaximized Then
            '' ''        maxScreenHeight = .Height
            '' ''        maxScreenWidth = .Width
            '' ''    Else
            '' ''        Dim formerState As Excel.XlWindowState = .WindowState
            '' ''        .WindowState = Excel.XlWindowState.xlMaximized
            '' ''        maxScreenHeight = .Height
            '' ''        maxScreenWidth = .Width
            '' ''        .WindowState = formerState
            '' ''    End If


            '' ''End With

            '' ''miniHeight = maxScreenHeight / 6
            '' ''miniWidth = maxScreenWidth / 10



            '' ''Dim oGrenze As Integer = UBound(frmCoord, 1)
            ' '' '' hier werden die Top- & Left- Default Positionen der Formulare gesetzt 
            '' ''For i = 0 To oGrenze
            '' ''    frmCoord(i, PTpinfo.top) = maxScreenHeight * 0.3
            '' ''    frmCoord(i, PTpinfo.left) = maxScreenWidth * 0.4
            '' ''Next

            ' '' '' jetzt setzen der Werte für Status-Information und Milestone-Information
            '' ''frmCoord(PTfrm.projInfo, PTpinfo.top) = 125
            '' ''frmCoord(PTfrm.projInfo, PTpinfo.left) = My.Computer.Screen.WorkingArea.Width - 500

            '' ''frmCoord(PTfrm.msInfo, PTpinfo.top) = 125 + 280
            '' ''frmCoord(PTfrm.msInfo, PTpinfo.left) = My.Computer.Screen.WorkingArea.Width - 500

            ' '' '' With listOfWorkSheets(arrWsNames(4))

            ' '' '' Logfile öffnen und ggf. initialisieren
            '' ''Call logfileOpen()


            Try

                ' prüft, ob bereits Excel geöffnet ist 
                'excelObj = GetObject(, "Excel.Application")
                appInstance = CType(GetObject(, "Excel.Application"), Microsoft.Office.Interop.Excel.Application)
            Catch ex As Exception
                Try
                    'excelObj = GetObject(, "Excel.Application")
                    appInstance = CType(CreateObject("Excel.Application"), Microsoft.Office.Interop.Excel.Application)
                Catch ex1 As Exception
                    Call MsgBox("Excel konnte nicht gestartet werden ..." & ex1.Message)
                    Exit Sub
                End Try

            End Try

            ' screeenUpdating auf false setzen, damit Excel nicht aufpoppt
            appInstance.ScreenUpdating = False

            Dim customizationFile As String = "requirements\Project Board Customization.xlsx"
            ' hier muss jetzt das Customization File aufgemacht werden ...
            Try
                xlsCustomization = appInstance.Workbooks.Open(awinPath & customizationFile)
                myCustomizationFile = appInstance.ActiveWorkbook.Name
            Catch ex As Exception
                'appInstance.ScreenUpdating = formerSU
                Throw New ArgumentException("Customization File nicht gefunden - Abbruch")
            End Try

            Dim wsName4 As Excel.Worksheet = CType(appInstance.Worksheets(arrWsNames(4)), _
                                                    Global.Microsoft.Office.Interop.Excel.Worksheet)

            ' '' '' hier muss Datenbank aus Customization-File gelesen werden, damit diese für den Login bekannt ist
            '' ''Try
            '' ''    awinSettings.databaseName = CStr(wsName4.Range("Datenbank").Value).Trim
            '' ''    If awinSettings.databaseName = "" Then
            '' ''        awinSettings.databaseName = "VisboTest"
            '' ''    End If
            '' ''Catch ex As Exception

            '' ''    awinSettings.databaseName = "VisboTest"
            '' ''    'appInstance.ScreenUpdating = formerSU
            '' ''    'Throw New ArgumentException("fehlende Einstellung im Customization-File; DB Name fehlt ... Abbruch " & vbLf & ex.Message)
            '' ''End Try

            If special = "BHTC" Then
                ' keine Datenbank angeschlossen
                ' kein LOGIN erforderlich


                Dim wsName7810 As Excel.Worksheet = CType(appInstance.Worksheets(arrWsNames(7)), _
                                                        Global.Microsoft.Office.Interop.Excel.Worksheet)

                Try
                    ' Aufbauen der Darstellungsklassen  
                    Call aufbauenAppearanceDefinitions(wsName7810)

                    ' Auslesen der BusinessUnit Definitionen
                    Call readBusinessUnitDefinitions(wsName4)

                    ' Auslesen der Phasen Definitionen 
                    Call readPhaseDefinitions(wsName4)

                    ' Auslesen der Meilenstein Definitionen 
                    Call readMilestoneDefinitions(wsName4)

                    ' Auslesen der Rollen Definitionen 
                    Call readRoleDefinitions(wsName4)

                    ' Auslesen der Kosten Definitionen 
                    Call readCostDefinitions(wsName4)

                    ' auslesen der anderen Informationen 
                    Call readOtherDefinitions(wsName4)

                    ' hier muss jetzt das Worksheet Phasen-Mappings aufgemacht werden, das ist in arrwsnames(8) abgelegt 
                    wsName7810 = CType(appInstance.Worksheets(arrWsNames(8)), _
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet)

                    Call readNameMappings(wsName7810, phaseMappings)


                    ' hier muss jetzt das Worksheet Milestone-Mappings aufgemacht werden, das ist in arrwsnames(10) abgelegt 
                    wsName7810 = CType(appInstance.Worksheets(arrWsNames(10)), _
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet)

                    Call readNameMappings(wsName7810, milestoneMappings)

                    ' '' '' '' jetzt muss die Seite mit den Appearance-Shapes kopiert werden 
                    '' '' ''appInstance.EnableEvents = False
                    '' '' ''CType(appInstance.Worksheets(arrWsNames(7)), _
                    '' '' ''Global.Microsoft.Office.Interop.Excel.Worksheet).Copy(After:=projectBoardSheet)

                    ' '' '' '' hier wird die Datei Projekt Tafel Customizations als aktives workbook wieder geschlossen ....
                    '' '' ''appInstance.Workbooks(myCustomizationFile).Close(SaveChanges:=False) ' ur: 6.5.2014 savechanges hinzugefügt
                    '' '' ''appInstance.EnableEvents = True


                    ' '' '' '' jetzt muss die apperanceDefinitions wieder neu aufgebaut werden 
                    '' '' ''appearanceDefinitions.Clear()
                    '' '' ''wsName7810 = CType(appInstance.Worksheets(arrWsNames(7)), _
                    '' '' ''                                        Global.Microsoft.Office.Interop.Excel.Worksheet)
                    '' '' ''Call aufbauenAppearanceDefinitions(wsName7810)


                    ' jetzt werden die ggf vorhandenen detaillierten Ressourcen Kapazitäten ausgelesen 
                    Call readRessourcenDetails()


                    ' '' '' '' jetzt werden die Modul-Vorlagen ausgelesen 
                    '' '' ''Call readVorlagen(True)

                    ' '' '' '' jetzt werden die Projekt-Vorlagen ausgelesen 
                    '' '' ''Call readVorlagen(False)

                    '' '' ''Dim a As Integer = Projektvorlagen.Count
                    '' '' ''Dim b As Integer = ModulVorlagen.Count

                    ' jetzt wird die Projekt-Tafel präpariert - Spaltenbreite und -Höhe
                    ' Beschriftung des Kalenders
                    '' '' ''appInstance.EnableEvents = False
                    '' '' ''Call prepareProjektTafel()


                    '' '' ''projectBoardSheet.Activate()
                    '' '' ''appInstance.EnableEvents = True

                    ' '' '' '' jetzt werden aus der Datenbank die Konstellationen und Dependencies gelesen 
                    '' '' ''Call readInitConstellations()

                Catch ex As Exception

                    appInstance.EnableEvents = True
                    Throw New ArgumentException(ex.Message)
                End Try

                '' '' Logfile wird geschlossen
                ' ''Call logfileSchliessen()


            Else ' es gilt : special <> "BHTC"


                ' '' '' ur: 23.01.2015: Abfragen der Login-Informationen
                '' ''loginErfolgreich = loginProzedur()



                '' ''If Not loginErfolgreich Then
                '' ''    ' Customization-File wird geschlossen
                '' ''    xlsCustomization.Close(SaveChanges:=False)
                '' ''    Call logfileSchreiben("LOGIN fehlerhaft", "", -1)
                '' ''    Call logfileSchliessen()
                '' ''    appInstance.Quit()
                '' ''    Exit Sub
                '' ''Else




                '' ''    Dim wsName7810 As Excel.Worksheet = CType(appInstance.Worksheets(arrWsNames(7)), _
                '' ''                                            Global.Microsoft.Office.Interop.Excel.Worksheet)

                '' ''    Try
                '' ''        ' Aufbauen der Darstellungsklassen  
                '' ''        Call aufbauenAppearanceDefinitions(wsName7810)

                '' ''        ' Auslesen der BusinessUnit Definitionen
                '' ''        Call readBusinessUnitDefinitions(wsName4)

                '' ''        ' Auslesen der Phasen Definitionen 
                '' ''        Call readPhaseDefinitions(wsName4)

                '' ''        ' Auslesen der Meilenstein Definitionen 
                '' ''        Call readMilestoneDefinitions(wsName4)

                '' ''        ' Auslesen der Rollen Definitionen 
                '' ''        Call readRoleDefinitions(wsName4)

                '' ''        ' Auslesen der Kosten Definitionen 
                '' ''        Call readCostDefinitions(wsName4)

                '' ''        ' auslesen der anderen Informationen 
                '' ''        Call readOtherDefinitions(wsName4)

                '' ''        ' hier muss jetzt das Worksheet Phasen-Mappings aufgemacht werden, das ist in arrwsnames(8) abgelegt 
                '' ''        wsName7810 = CType(appInstance.Worksheets(arrWsNames(8)), _
                '' ''                                                Global.Microsoft.Office.Interop.Excel.Worksheet)

                '' ''        Call readNameMappings(wsName7810, phaseMappings)


                '' ''        ' hier muss jetzt das Worksheet Milestone-Mappings aufgemacht werden, das ist in arrwsnames(10) abgelegt 
                '' ''        wsName7810 = CType(appInstance.Worksheets(arrWsNames(10)), _
                '' ''                                                Global.Microsoft.Office.Interop.Excel.Worksheet)

                '' ''        Call readNameMappings(wsName7810, milestoneMappings)

                '' ''        ' jetzt muss die Seite mit den Appearance-Shapes kopiert werden 
                '' ''        appInstance.EnableEvents = False
                '' ''        CType(appInstance.Worksheets(arrWsNames(7)), _
                '' ''        Global.Microsoft.Office.Interop.Excel.Worksheet).Copy(After:=projectBoardSheet)

                '' ''        ' hier wird die Datei Projekt Tafel Customizations als aktives workbook wieder geschlossen ....
                '' ''        appInstance.Workbooks(myCustomizationFile).Close(SaveChanges:=False) ' ur: 6.5.2014 savechanges hinzugefügt
                '' ''        appInstance.EnableEvents = True


                '' ''        ' jetzt muss die apperanceDefinitions wieder neu aufgebaut werden 
                '' ''        appearanceDefinitions.Clear()
                '' ''        wsName7810 = CType(appInstance.Worksheets(arrWsNames(7)), _
                '' ''                                                Global.Microsoft.Office.Interop.Excel.Worksheet)
                '' ''        Call aufbauenAppearanceDefinitions(wsName7810)


                '' ''        ' jetzt werden die ggf vorhandenen detaillierten Ressourcen Kapazitäten ausgelesen 
                '' ''        Call readRessourcenDetails()


                '' ''        ' jetzt werden die Modul-Vorlagen ausgelesen 
                '' ''        Call readVorlagen(True)

                '' ''        ' jetzt werden die Projekt-Vorlagen ausgelesen 
                '' ''        Call readVorlagen(False)

                '' ''        Dim a As Integer = Projektvorlagen.Count
                '' ''        Dim b As Integer = ModulVorlagen.Count

                '' ''        ' jetzt wird die Projekt-Tafel präpariert - Spaltenbreite und -Höhe
                '' ''        ' Beschriftung des Kalenders
                '' ''        appInstance.EnableEvents = False
                '' ''        Call prepareProjektTafel()


                '' ''        projectBoardSheet.Activate()
                '' ''        appInstance.EnableEvents = True

                '' ''        ' jetzt werden aus der Datenbank die Konstellationen und Dependencies gelesen 
                '' ''        Call readInitConstellations()

                '' ''    Catch ex As Exception
                '' ''        appInstance.ScreenUpdating = formerSU
                '' ''        appInstance.EnableEvents = True
                '' ''        Throw New ArgumentException(ex.Message)
                '' ''    End Try

                ' '' '' Logfile wird geschlossen
                '' ''Call logfileSchliessen()


                '' ''End If  ' von "if Login erfolgt"

            End If ' von "if special="BHTC"


        Catch ex As Exception
            Call MsgBox("Fehler beim Laden des VISBO AddIn")
            fehlerBeimLoad = True

        End Try
    End Sub

    ''' <summary>
    ''' liest die Business Unit Definitionen aus der awinsetTypen
    ''' die globale Variable businessUnitDefinitions wird dabei befüllt
    ''' die erste und letzte Zeile des Range wird ignoriert 
    ''' </summary>
    ''' <param name="wsname">Name des Excel Worksheets, das die Infos im aktuellen Workbook enthält</param>
    ''' <remarks></remarks>
    Private Sub readBusinessUnitDefinitions(ByVal wsname As Excel.Worksheet)

        ' hier werden jetzt die Business Unit Informationen ausgelesen 
        businessUnitDefinitions = New SortedList(Of Integer, clsBusinessUnit)

        Try

            With wsname
                '
                ' Business Unit Definitionen auslesen - im bereich awin_BusinessUnit_Definitions
                '
                Dim index As Integer = 1
                Dim tmpBU As clsBusinessUnit

                Dim BURange As Excel.Range = CType(.Range("awin_BusinessUnit_Definitions"), Excel.Range)
                Dim anzZeilen As Integer = BURange.Rows.Count

                For i As Integer = 2 To anzZeilen - 1

                    tmpBU = New clsBusinessUnit

                    Try
                        tmpBU.name = CStr(BURange.Cells(i, 1).value).Trim
                        tmpBU.color = CLng(BURange.Cells(i, 1).Interior.color)

                        If tmpBU.name.Length > 0 Then
                            businessUnitDefinitions.Add(i - 1, tmpBU)
                        End If

                    Catch ex As Exception
                        ' nichts tun ...

                    End Try

                Next

            End With

        Catch ex As Exception
            Throw New ArgumentException("Fehler im Customization-File: BU Definition")
        End Try
        

    End Sub
    
    ''' <summary>
    ''' liest die Phasen Definitionen aus 
    ''' baut die globale Variable PhaseDefinitions auf 
    ''' </summary>
    ''' <param name="wsname">Name des Worksheets, aus dem die Infos ausgelesen werden</param>
    ''' <remarks></remarks>
    Private Sub readPhaseDefinitions(ByVal wsname As Excel.Worksheet)

        Dim hphase As clsPhasenDefinition
        Dim tmpStr As String = ""

        Try

            With wsname

                Dim phaseRange As Excel.Range = .Range("awin_Phasen_Definition")
                Dim anzZeilen As Integer = phaseRange.Rows.Count
                Dim c As Excel.Range

                For iZeile As Integer = 2 To anzZeilen - 1

                    c = CType(phaseRange.Cells(iZeile, 1), Excel.Range)

                    If Not IsNothing(c.Value) Then

                        If CStr(c.Value) <> "" Then
                            tmpStr = CType(c.Value, String)
                            ' das neue ...
                            hphase = New clsPhasenDefinition
                            With hphase
                                '.farbe = CLng(c.Interior.Color)
                                .name = tmpStr.Trim
                                .UID = iZeile - 1

                                ' hat die Phase einen Schwellwert ? 
                                Try
                                    If CInt(c.Offset(0, 1).Value) > 0 Then
                                        .schwellWert = CInt(c.Offset(0, 1).Value)
                                    End If
                                Catch ex As Exception

                                End Try

                                ' ist die Phase eine special Phase ? 
                                Try
                                    If CStr(c.Offset(0, 2).Value).Trim = "LeLe" Then
                                        specialListofPhases.Add(hphase.name, hphase.name)
                                    End If
                                Catch ex As Exception
                                End Try



                                ' hat die Phase eine Abkürzung ? 
                                Dim abbrev As String = ""
                                If Not IsNothing(c.Offset(0, 5).Value) Then
                                    abbrev = CStr(c.Offset(0, 5).Value).Trim
                                End If

                                .shortName = abbrev


                                ' hat die Phase eine Darstellungsklasse ? 
                                Try
                                    Dim darstellungsklasse As String
                                    If Not IsNothing(c.Offset(0, 6).Value) Then

                                        If CStr(c.Offset(0, 6).Value).Trim.Length > 0 Then
                                            darstellungsklasse = CStr(c.Offset(0, 6).Value).Trim
                                            If appearanceDefinitions.ContainsKey(darstellungsklasse) Then
                                                .darstellungsKlasse = darstellungsklasse
                                            Else
                                                .darstellungsKlasse = ""
                                            End If
                                        End If

                                    End If

                                Catch ex As Exception
                                    .darstellungsKlasse = ""
                                End Try



                            End With

                            Try
                                PhaseDefinitions.Add(hphase)
                            Catch ex As Exception

                            End Try


                        End If

                    End If


                Next


            End With

        Catch ex As Exception

            Throw New ArgumentException("Fehler in Customization File: Phasen")

        End Try


    End Sub

    ''' <summary>
    ''' liest die Phasen Definitionen aus 
    ''' </summary>
    ''' <param name="wsname">Name des Worksheets, aus dem die Infos ausgelesen werden</param>
    ''' <remarks></remarks>
    Private Sub readMilestoneDefinitions(ByVal wsname As Excel.Worksheet)

        Dim i As Integer = 0
        Dim hMilestone As clsMeilensteinDefinition
        Dim tmpStr As String


        Try

            With wsname

                Dim milestoneRange As Excel.Range = .Range("awin_Meilenstein_Definition")
                Dim anzZeilen As Integer = milestoneRange.Rows.Count
                Dim c As Excel.Range

                For iZeile As Integer = 2 To anzZeilen - 1

                    c = CType(milestoneRange.Cells(iZeile, 1), Excel.Range)

                    ' hier muss das Aufbauen der MilestoneDefinitions gemacht werden  
                    If Not IsNothing(c.Value) Then

                        If CStr(c.Value) <> "" Then
                            i = i + 1
                            tmpStr = CType(c.Value, String)
                            ' das neue ...
                            hMilestone = New clsMeilensteinDefinition
                            With hMilestone
                                .name = tmpStr.Trim
                                .UID = i

                                ' hat der Milestone einen Schwellwert ? 

                                If IsNothing(c.Offset(0, 1).Value) Then
                                ElseIf IsNumeric(c.Offset(0, 1).Value) Then
                                    If CInt(c.Offset(0, 1).Value) > 0 Then
                                        .schwellWert = CInt(c.Offset(0, 1).Value)
                                    End If
                                End If


                                ' hat der Milestone einen Bezug ? 
                                Dim bezug As String = ""
                                If Not IsNothing(c.Offset(0, 4).Value) Then

                                    bezug = CStr(c.Offset(0, 4).Value).Trim

                                    If PhaseDefinitions.Contains(bezug) Then
                                    Else
                                        bezug = ""
                                    End If

                                End If

                                .belongsTo = bezug

                                ' hat der Milestone eine Abkürzung ? 
                                Dim abbrev As String = ""
                                If Not IsNothing(c.Offset(0, 5).Value) Then
                                    abbrev = CStr(c.Offset(0, 5).Value).Trim
                                End If

                                .shortName = abbrev


                                ' hat der Milestone eine Darstellungsklasse ? 

                                Dim darstellungsklasse As String = ""
                                If Not IsNothing(c.Offset(0, 6).Value) Then

                                    If CStr(c.Offset(0, 6).Value).Trim.Length > 0 Then
                                        darstellungsklasse = CStr(c.Offset(0, 6).Value).Trim
                                        If appearanceDefinitions.ContainsKey(darstellungsklasse) Then
                                            .darstellungsKlasse = darstellungsklasse
                                        Else
                                            .darstellungsKlasse = ""
                                        End If
                                    End If

                                End If



                            End With

                            Try
                                MilestoneDefinitions.Add(hMilestone)
                            Catch ex As Exception

                            End Try


                        End If

                    End If

                Next

            End With

        Catch ex As Exception

            Throw New ArgumentException("Fehler in Customization File: Meilensteine")

        End Try
        

    End Sub

    ''' <summary>
    ''' liest die Rollen Definitionen ein 
    ''' wird in der globalen Variablen RoleDefinitions abgelegt 
    ''' </summary>
    ''' <param name="wsname"></param>
    ''' <remarks></remarks>
    Private Sub readRoleDefinitions(ByVal wsname As Excel.Worksheet)

        '
        ' Rollen Definitionen auslesen - im bereich awin_Rollen_Definition
        '
        Dim index As Integer = 0
        Dim tmpStr As String
        Dim hrole As clsRollenDefinition


        Try


            With wsname

                Dim rolesRange As Excel.Range = .Range("awin_Rollen_Definition")
                Dim anzZeilen As Integer = rolesRange.Rows.Count
                Dim c As Excel.Range

                For i = 2 To anzZeilen - 1
                    c = CType(rolesRange.Cells(i, 1), Excel.Range)

                    If CStr(c.Value) <> "" Then
                        index = index + 1
                        tmpStr = CType(c.Value, String)
                        If index = 1 Then
                            rollenKapaFarbe = c.Offset(0, 1).Interior.Color
                        End If


                        ' jetzt kommt die Rollen Definition 
                        hrole = New clsRollenDefinition
                        Dim cp As Integer
                        With hrole
                            .name = tmpStr.Trim
                            .Startkapa = CDbl(c.Offset(0, 1).Value)
                            .tagessatzIntern = CDbl(c.Offset(0, 2).Value)

                            Try
                                If CDbl(c.Offset(0, 3).Value) = 0.0 Then
                                    .tagessatzExtern = .tagessatzIntern * 1.35
                                Else
                                    .tagessatzExtern = CDbl(c.Offset(0, 3).Value)
                                End If
                            Catch ex As Exception
                                .tagessatzExtern = .tagessatzIntern * 1.35
                            End Try

                            ' Auslesen der zukünftigen Kapazität
                            ' Änderung 29.5.14: von StartofCalendar 240 Monate nach vorne kucken ... 
                            For cp = 1 To 240

                                .kapazitaet(cp) = .Startkapa
                                .externeKapazitaet(cp) = 0.0

                            Next
                            .farbe = c.Interior.Color
                            .UID = index
                        End With

                        '
                        RoleDefinitions.Add(hrole)
                        'hrole = Nothing

                    End If

                Next

            End With

        Catch ex As Exception
            Throw New ArgumentException("Fehler im Customization-File: Rolle")
        End Try

        

    End Sub


    ''' <summary>
    ''' liest die Kosten Definitionen ein 
    ''' wird in der globalen Variablen CostDefinitions abgelegt 
    ''' </summary>
    ''' <param name="wsname"></param>
    ''' <remarks></remarks>
    Private Sub readCostDefinitions(ByVal wsname As Excel.Worksheet)


        Dim index As Integer = 0
        Dim hcost As clsKostenartDefinition
        Dim tmpStr As String


        Try

            With wsname

                Dim costRange As Excel.Range = .Range("awin_Kosten_Definition")
                Dim anzZeilen As Integer = costRange.Rows.Count
                Dim c As Excel.Range

                For i As Integer = 2 To anzZeilen - 1

                    c = CType(costRange.Cells(i, 1), Excel.Range)
                    If CStr(c.Value) <> "" Or index > 0 Then
                        index = index + 1

                        ' jetzt kommt die Kostenarten Definition
                        hcost = New clsKostenartDefinition
                        With hcost
                            If CStr(c.Value) <> "" Then
                                tmpStr = CType(c.Value, String)
                                .name = tmpStr.Trim
                            Else
                                .name = "Personalkosten"
                            End If
                            .farbe = c.Interior.Color
                            .UID = index
                        End With

                        CostDefinitions.Add(hcost)
                    End If

                Next

            End With

        Catch ex As Exception
            Throw New ArgumentException("Fehler in Customization File: Kosten")
        End Try
        

    End Sub


    ''' <summary>
    ''' liest die sonstigen Einstellungen wie Farben, Spaltenbreite, Spaltenhöhe etc aus
    ''' wird in entsprechenden globalen Variablen abgelegt  
    ''' </summary>
    ''' <param name="wsname">Name des Worksheets, aus dem die Infos ausgelesen werden</param>
    ''' <remarks></remarks>
    Private Sub readOtherDefinitions(ByVal wsname As Excel.Worksheet)


        With wsname
            Try
                'showRangeLeft = CInt(.Range("Linker_Rand_Ressourcen_Diagramme").Value)
                'showRangeRight = CInt(.Range("Rechter_Rand_Ressourcen_Diagramme").Value)
                showtimezone_color = .Range("Show_Time_Zone_Color").Interior.Color
                noshowtimezone_color = .Range("NoShow_Time_Zone_Color").Interior.Color
                nrOfDaysMonth = CDbl(.Range("Arbeitstage_pro_Monat").Value)
                farbeInternOP = .Range("Farbe_intern_ohne_Projekte").Interior.Color
                farbeExterne = .Range("Farbe_externe_Ressourcen").Interior.Color
                iProjektFarbe = .Range("Farbe_für_Projekte_ohne_Vorlage").Interior.Color
                iWertFarbe = .Range("Farbe_Ress_Kost_Werte").Interior.Color
                vergleichsfarbe0 = .Range("Vergleichsfarbe1").Interior.Color
                vergleichsfarbe1 = .Range("Vergleichsfarbe2").Interior.Color
                vergleichsfarbe2 = .Range("Vergleichsfarbe3").Interior.Color

                'Dim tmpcolor As Microsoft.Office.Interop.Excel.ColorFormat

                Try
                    awinSettings.SollIstFarbeB = CLng(.Range("Soll_Ist_Farbe_Beauftragung").Interior.Color)
                    awinSettings.SollIstFarbeL = CLng(.Range("Soll_Ist_Farbe_letzte_Freigabe").Interior.Color)
                    awinSettings.SollIstFarbeC = CLng(.Range("Soll_Ist_Farbe_Aktuell").Interior.Color)
                    awinSettings.AmpelGruen = CLng(.Range("AmpelGruen").Interior.Color)
                    'tmpcolor = CType(.Range("AmpelGruen").Interior.Color, Microsoft.Office.Interop.Excel.ColorFormat)
                    awinSettings.AmpelGelb = CLng(.Range("AmpelGelb").Interior.Color)
                    awinSettings.AmpelRot = CLng(.Range("AmpelRot").Interior.Color)
                    awinSettings.AmpelNichtBewertet = CLng(.Range("AmpelNichtBewertet").Interior.Color)
                    awinSettings.glowColor = CLng(.Range("GlowColor").Interior.Color)

                    Try
                        awinSettings.timeSpanColor = CLng(.Range("FarbeZeitraum").Interior.Color)
                        awinSettings.showTimeSpanInPT = CBool(.Range("FarbeZeitraum").Value)
                    Catch ex2 As Exception
                        ' ansonsten wird die Voreinstellung verwendet 
                    End Try


                Catch ex As Exception
                    Throw New ArgumentException("Customization File fehlerhaft - Farben fehlen ... " & vbLf & ex.Message)
                End Try

                Try
                    awinSettings.missingDefinitionColor = CLng(.Range("MissingDefinitionColor").Interior.Color)
                Catch ex As Exception

                End Try

                ergebnisfarbe1 = .Range("Ergebnisfarbe1").Interior.Color
                ergebnisfarbe2 = .Range("Ergebnisfarbe2").Interior.Color
                weightStrategicFit = CDbl(.Range("WeightStrategicFit").Value)
                ' jetzt wird KalenderStart, Zeiteinheit und Datenbank Name ausgelesen 
                awinSettings.kalenderStart = CDate(.Range("Start_Kalender").Value)
                awinSettings.zeitEinheit = CStr(.Range("Zeiteinheit").Value)
                awinSettings.kapaEinheit = CStr(.Range("kapaEinheit").Value)
                awinSettings.offsetEinheit = CStr(.Range("offsetEinheit").Value)
                'ur: 6.08.2015: umgestellt auf Settings in app.config ''awinSettings.databaseName = CStr(.Range("Datenbank").Value)
                awinSettings.EinzelRessExport = CInt(.Range("EinzelRessourcenExport").Value)
                awinSettings.zeilenhoehe1 = CDbl(.Range("Zeilenhoehe1").Value)
                awinSettings.zeilenhoehe2 = CDbl(.Range("Zeilenhoehe2").Value)
                awinSettings.spaltenbreite = CDbl(.Range("Spaltenbreite").Value)
                awinSettings.autoCorrectBedarfe = True
                awinSettings.propAnpassRess = False
                awinSettings.showValuesOfSelected = False
            Catch ex As Exception
                Throw New ArgumentException("fehlende Einstellung im Customization-File ... Abbruch " & vbLf & ex.Message)
            End Try

            ' ist Einstellung für volles Protokoll vorhanden ? 
            Try

                awinSettings.fullProtocol = CBool(.Range("volles_Protokol").Value)
            Catch ex As Exception
                awinSettings.fullProtocol = False
            End Try

            ' Einstellung für addMissingDefinitions
            Try
                awinSettings.addMissingPhaseMilestoneDef = CBool(.Range("addMissingDefinitions").Value)
            Catch ex As Exception
                awinSettings.addMissingPhaseMilestoneDef = False
            End Try

            ' Einstellung für alwaysAcceptTemplate Names 
            Try
                awinSettings.alwaysAcceptTemplateNames = CBool(.Range("alywaysAcceptTemplateDefs").Value)
            Catch ex As Exception
                awinSettings.alwaysAcceptTemplateNames = False
            End Try

            ' Einstellungen, um Duplikate zu eliminieren ; 
            Try
                awinSettings.eliminateDuplicates = CBool(.Range("eliminate_Duplicates").Value)
            Catch ex As Exception
                awinSettings.eliminateDuplicates = True
            End Try

            ' Einstellungen, um unbekannte Namen zu importieren 
            Try
                awinSettings.importUnknownNames = CBool(.Range("importUnknownNames").Value)
            Catch ex As Exception
                awinSettings.importUnknownNames = True
            End Try

            ' Einstellung, um Geschwister-Namen immer eindeutig zu machen
            Try
                awinSettings.createUniqueSiblingNames = CBool(.Range("uniqueSiblingNames").Value)
            Catch ex As Exception
                awinSettings.createUniqueSiblingNames = True
            End Try



            StartofCalendar = awinSettings.kalenderStart
            StartofCalendar = StartofCalendar.ToLocalTime()

            historicDate = StartofCalendar

            ' Import Typ regelt, um welche DateiFormate es sich bei dem Import handelt
            ' 1: Standard
            ' 2: BMW Rplan Export in Excel 
            Try
                awinSettings.importTyp = CInt(.Range("Import_Typ").Value)
            Catch ex As Exception
                awinSettings.importTyp = 1
            End Try

            '
            ' ende Auslesen Einstellungen in Sheet "Einstellungen"
        End With
        

    End Sub

    ''' <summary>
    ''' liest für die definierten Rollen ggf vorhandene detaillierte Ressourcen Kapazitäten ein 
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub readRessourcenDetails()

        ' jetzt werden  für die einzelnen Rollen in dem Directory Ressource Manager Dateien 
        ' die evtl vorhandenen Dateien für die genaue Bestimmung der Kapazität ausgelesen  
        Dim tmpRole As clsRollenDefinition
        Dim tmpRoleDefinitions As New clsRollen
        Dim ix As Integer
        For ix = 1 To RoleDefinitions.Count
            tmpRole = RoleDefinitions.getRoledef(ix)
            ' hier werden die betreffenden Dateien geöffnet und auch wieder geschlossen
            ' wenn es zu Problemen kommen sollte, bleiben die Kapa Werte unverändert ...
            Call readKapaOfRole(tmpRole)
            tmpRoleDefinitions.Add(tmpRole)
        Next

        RoleDefinitions = New clsRollen
        RoleDefinitions = tmpRoleDefinitions

    End Sub

    ''' <summary>
    ''' liest die Projekt- bzw. Modul-Vorlagen ein 
    ''' </summary>
    ''' <param name="isModulVorlage"></param>
    ''' <remarks></remarks>
    Private Sub readVorlagen(ByVal isModulVorlage As Boolean)

        Dim dirName As String
        Dim dateiName As String

        If isModulVorlage Then
            dirName = awinPath & modulVorlagenOrdner
        Else
            dirName = awinPath & projektVorlagenOrdner
        End If

        If My.Computer.FileSystem.DirectoryExists(dirName) Then

            Dim listOfFiles As Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Computer.FileSystem.GetFiles(dirName)

            For i As Integer = 1 To listOfFiles.Count

                dateiName = listOfFiles.Item(i - 1)
                If dateiName.Contains(".xls") Or dateiName.Contains(".xlsx") Then
                    Try

                        appInstance.Workbooks.Open(dateiName)


                        If awinSettings.importTyp = 1 Then



                            Dim projVorlage As New clsProjektvorlage

                            ' Auslesen der Projektvorlage wird wie das Importieren eines Projekts behandelt, nur am Ende in die Liste der Projektvorlagen eingehängt
                            ' Kennzeichen für Projektvorlage ist der 3.Parameter im Aufruf (isTemplate)

                            Call awinImportProjectmitHrchy(Nothing, projVorlage, True, Date.Now)

                            ' ur: 21.05.2015: Vorlagen nun neues Format, mit Hierarchie
                            ' Call awinImportProject(Nothing, projVorlage, True, Date.Now)

                            If isModulVorlage Then
                                ModulVorlagen.Add(projVorlage)
                            Else
                                Projektvorlagen.Add(projVorlage)
                            End If



                        ElseIf awinSettings.importTyp = 2 Then

                            ' hier muss die Datei ausgelesen werden
                            Dim myCollection As New Collection
                            Dim ok As Boolean
                            Dim hproj As clsProjekt = Nothing

                            Call bmwImportProjekteITO15(myCollection, True)

                            ' jetzt muss für jeden Eintrag in ImportProjekte eine Vorlage erstellt werden  
                            For Each pName As String In myCollection

                                ok = True

                                Try

                                    hproj = ImportProjekte.getProject(pName)

                                Catch ex As Exception
                                    Call MsgBox("Projekt " & pName & " ist kein gültiges Projekt ... es wird ignoriert ...")
                                    ok = False
                                End Try

                                If ok Then

                                    ' hier müssen die Werte für die Vorlage übergeben werden.
                                    ' Änderung tk 19.4.15 Übernehmen der Hierarchie 
                                    Dim projVorlage As New clsProjektvorlage
                                    projVorlage.VorlagenName = hproj.name
                                    projVorlage.Schrift = hproj.Schrift
                                    projVorlage.Schriftfarbe = hproj.Schriftfarbe
                                    projVorlage.farbe = hproj.farbe
                                    projVorlage.earliestStart = -6
                                    projVorlage.latestStart = 6
                                    projVorlage.AllPhases = hproj.AllPhases

                                    projVorlage.hierarchy = hproj.hierarchy

                                    If isModulVorlage Then
                                        ModulVorlagen.Add(projVorlage)
                                    Else
                                        Projektvorlagen.Add(projVorlage)
                                    End If

                                End If

                            Next


                        End If
                        ' ur: Test
                        Dim anzphase As Integer = PhaseDefinitions.Count

                        appInstance.ActiveWorkbook.Close(SaveChanges:=True)


                    Catch ex As Exception
                        appInstance.ActiveWorkbook.Close(SaveChanges:=True)
                        Call MsgBox(ex.Message)
                    End Try
                End If


            Next

            Try
                If isModulVorlage Then
                    If ModulVorlagen.Count > 0 Then
                        awinSettings.lastModulTyp = ModulVorlagen.Liste.ElementAt(0).Value.VorlagenName
                        ' Änderung tk 26.11.15 muss doch hier gar nicht gemacht werden .. erst mit Beenden des Wörterbuchs bzw. Beenden der Applikation
                        'Call awinWritePhaseDefinitions()
                        'Call awinWritePhaseMilestoneDefinitions 
                    End If

                Else
                    If Projektvorlagen.Count > 0 Then
                        awinSettings.lastProjektTyp = Projektvorlagen.Liste.ElementAt(0).Value.VorlagenName
                        'Call awinWritePhaseDefinitions()
                        'Call awinWritePhaseMilestoneDefinitions 
                    End If

                End If

            Catch ex As Exception
                awinSettings.lastProjektTyp = ""
            End Try



        Else
            If isModulVorlage Then
                ' nichts tun - kein Problem, wenn es keine Vorlagen gibt 
            Else
                Throw New ArgumentException("der Vorlagen Ordner fehlt:" & vbLf & dirName)
            End If
        End If
        

    End Sub



    ''' <summary>
    ''' gibt für ein gegebenes Projekt die errechnete Farbe und den errechneten Status zurück
    ''' dabei wird das aktuelle Projekt in Relation zur Beauftragung/letzten Freigabe gesetzt
    ''' wenn es noch keine Projekt-Historie gibt, so wird grün und "0" zurückgegeben   
    ''' </summary>
    ''' <param name="hproj"></param>
    ''' das Projekt in Question
    ''' <param name="compareTo">
    ''' =0 : Vergleich mit erstem Projekt-Stand überhaupt
    ''' =1 : Vergleich mit Beauftragung 
    ''' =2 : Vergleich mit letzter Freigabe
    ''' =3 : Vergleich mit letztem Planungs-Stand
    ''' </param>
    ''' <param name="auswahl">
    ''' einer der Werte aus 1=Personalkosten, 2=Sonstige Kosten, 3=Gesamtkosten, 4=Rolle, 5=Kostenart 
    ''' </param>
    ''' <param name="statusValue">
    ''' Rückgabe Paarmeter - ein Wert zwischen 0 und sehr groß; je größer über 1, desto besser im Fortschritt 
    ''' je kleiner unter 1, desto schlechter im Fortschritt
    ''' </param>
    ''' <param name="statusColor">
    ''' Rückgabe Parameter: entweder grün, gelb oder rot
    ''' </param>
    ''' <remarks></remarks>
    Public Sub getStatusColorProject(ByRef hproj As clsProjekt, ByVal compareTo As Integer, ByVal auswahl As Integer, ByVal qualifier As String, _
                                  ByRef statusValue As Double, ByRef statusColor As Long)

        ' ''Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
        ' ''Dim currentValues() As Double
        ' ''Dim formerValues() As Double
        ' ''Dim vglProj As clsProjekt
        ' ''Dim vglName As String, pname As String, variantName As String
        ' ''Dim anzSnapshots As Integer, index As Integer
        ' ''Dim heuteColumn As Integer = getColumnOfDate(Date.Now)
        ' ''Dim cValue As Double, fValue As Double

        ' ''With hproj
        ' ''    pname = .name
        ' ''    variantName = .variantName
        ' ''End With
        ' ''vglName = " "

        ' ''Try
        ' ''    ReDim currentValues(hproj.anzahlRasterElemente - 1)

        ' ''Catch ex As Exception

        ' ''    statusValue = 1.0
        ' ''    statusColor = awinSettings.AmpelGruen
        ' ''    Exit Sub

        ' ''End Try


        ' ''If Not projekthistorie Is Nothing Then
        ' ''    If projekthistorie.Count > 0 Then
        ' ''        vglName = projekthistorie.First.getShapeText
        ' ''    End If
        ' ''Else
        ' ''    projekthistorie = New clsProjektHistorie
        ' ''End If


        ' ''If vglName <> hproj.getShapeText Then
        ' ''    If request.pingMongoDb() Then
        ' ''        ' projekthistorie muss nur dann neu bestimmt werden, wenn sie nicht bereits für dieses Projekt geholt wurde
        ' ''        projekthistorie.liste = request.retrieveProjectHistoryFromDB(projectname:=pname, variantName:=variantName, _
        ' ''                                                           storedEarliest:=StartofCalendar, storedLatest:=Date.Now)
        ' ''        If projekthistorie.Count > 0 Then
        ' ''            projekthistorie.Add(Date.Now, hproj)
        ' ''        End If

        ' ''    Else
        ' ''        Call MsgBox(" Datenbank-Verbindung ist unterbrochen!" & vbLf & " Projekthistorie kann nicht geladen werden")
        ' ''    End If

        ' ''Else
        ' ''    ' es muss nichts gemacht werden - es ist bereits die richtige Historie 
        ' ''End If


        '' '' jetzt sind in der Projekt-Historie die richtigen Snapshots 
        '' '' jetzt muss das Vergleichs-Projekt gesetzt werden 




        ' ''Try
        ' ''    anzSnapshots = projekthistorie.Count
        ' ''Catch ex1 As Exception
        ' ''    anzSnapshots = 0
        ' ''End Try



        ' ''If anzSnapshots > 0 Then

        ' ''    Select Case compareTo

        ' ''        Case 0
        ' ''            ' mit erstem Planungs-Stand vergleichen
        ' ''            vglProj = projekthistorie.ElementAt(0)

        ' ''        Case 1
        ' ''            ' mit Beauftragung vergleichen 
        ' ''            Try
        ' ''                vglProj = projekthistorie.beauftragung
        ' ''            Catch ex As Exception
        ' ''                vglProj = projekthistorie.ElementAt(0)
        ' ''            End Try


        ' ''        Case 2
        ' ''            ' mit letzter Freigabe vergleichen
        ' ''            index = getIndexPrevFreigabe(projekthistorie.liste, anzSnapshots - 1)
        ' ''            vglProj = projekthistorie.ElementAt(index)

        ' ''        Case 3
        ' ''            ' mit letztem Stand vergleichen , das ist das vorletzte Element , da hproj auf der letzten Position ist
        ' ''            vglProj = projekthistorie.ElementAt(anzSnapshots - 2)

        ' ''        Case Else
        ' ''            ' mit erstem Element vergleichen 
        ' ''            vglProj = projekthistorie.ElementAt(0)
        ' ''    End Select


        ' ''    ReDim formerValues(vglProj.anzahlRasterElemente - 1)
        ' ''    Dim hsum As Double = 0.0
        ' ''    Dim vsum As Double = 0.0

        ' ''    Select Case auswahl
        ' ''        Case 1
        ' ''            currentValues = hproj.getAllPersonalKosten
        ' ''            formerValues = vglProj.getAllPersonalKosten

        ' ''        Case 2
        ' ''            currentValues = hproj.getGesamtAndereKosten
        ' ''            formerValues = vglProj.getGesamtAndereKosten

        ' ''        Case 3
        ' ''            currentValues = hproj.getGesamtKostenBedarf
        ' ''            formerValues = vglProj.getGesamtKostenBedarf

        ' ''        Case 4
        ' ''            If RoleDefinitions.Contains(qualifier) Then
        ' ''                currentValues = hproj.getRessourcenBedarf(qualifier)
        ' ''                formerValues = vglProj.getRessourcenBedarf(qualifier)
        ' ''            End If
        ' ''        Case 5
        ' ''            If CostDefinitions.Contains(qualifier) Then
        ' ''                currentValues = hproj.getKostenBedarf(qualifier)
        ' ''                formerValues = vglProj.getKostenBedarf(qualifier)
        ' ''            End If
        ' ''        Case Else
        ' ''            ' wie Gesamtkosten
        ' ''            currentValues = hproj.getGesamtKostenBedarf
        ' ''            formerValues = vglProj.getGesamtKostenBedarf

        ' ''    End Select


        ' ''    ' jetzt muss abgefangen werden, daß in dem Vergleichs-Projekt gar keine Werte dafür da sind 
        ' ''    ' in dem aktuellen Projekt dagegen schon ; oder umgekehrt 
        ' ''    ' es muss natürlich auch abgefangen werden, daß der Wert bei beiden nicht existiert 


        ' ''    If currentValues.Sum <= 0 And formerValues.Sum <= 0 Then
        ' ''        statusValue = 0.0
        ' ''        statusColor = awinSettings.AmpelNichtBewertet
        ' ''        ' beide existieren nicht 

        ' ''    ElseIf currentValues.Sum <= 0 Then
        ' ''        statusValue = 0.0
        ' ''        statusColor = awinSettings.AmpelRot

        ' ''    ElseIf formerValues.Sum <= 0 Then
        ' ''        statusValue = 2.0
        ' ''        statusColor = awinSettings.AmpelGruen

        ' ''    Else
        ' ''        Dim korrFaktor As Double = formerValues.Sum / currentValues.Sum

        ' ''        For h = hproj.Start To heuteColumn - 1
        ' ''            hsum = hsum + currentValues(h - hproj.Start)
        ' ''        Next
        ' ''        cValue = hsum / currentValues.Sum

        ' ''        For v = vglProj.Start To heuteColumn - 1
        ' ''            vsum = vsum + formerValues(v - vglProj.Start)
        ' ''        Next
        ' ''        fValue = vsum / formerValues.Sum

        ' ''        If fValue > 0 Then
        ' ''            statusValue = korrFaktor * cValue / fValue
        ' ''        Else
        ' ''            statusValue = 2
        ' ''        End If

        ' ''        If statusValue >= 1.0 Then
        ' ''            statusColor = awinSettings.AmpelGruen
        ' ''        ElseIf statusValue >= 0.9 Then
        ' ''            statusColor = awinSettings.AmpelGelb
        ' ''        Else
        ' ''            statusColor = awinSettings.AmpelRot
        ' ''        End If

        ' ''    End If

        ' ''Else
        ' ''    statusValue = 1.0
        ' ''    statusColor = awinSettings.AmpelGruen
        ' ''End If

    End Sub




    ''' <summary>
    ''' liest die im Diretory ../ressource manager liegenden detaillierten Kapa files zu den Rollen aus
    ''' und hinterlegt es an entsprechender Stelle im hrole.kapazitaet
    ''' </summary>
    ''' <param name="hrole"></param>
    ''' <remarks></remarks>
    Friend Sub readKapaOfRole(ByRef hrole As clsRollenDefinition)
        Dim kapaFileName As String
        Dim ok As Boolean = True
        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim formerSU As Boolean = appInstance.ScreenUpdating
        Dim summenZeile As Integer, extSummenZeile As Integer
        Dim spalte As Integer = 2
        Dim blattname As String = "Kapazität"
        Dim currentWS As Excel.Worksheet
        Dim index As Integer
        Dim tmpDate As Date
        Dim tmpKapa As Double
        Dim extTmpKapa As Double

        If formerEE Then
            appInstance.EnableEvents = False
        End If

        If formerSU Then
            appInstance.ScreenUpdating = False
        End If

        enableOnUpdate = False

        kapaFileName = awinPath & projektRessOrdner & "\" & hrole.name & " Kapazität.xlsx"


        ' öffnen des Files 
        If My.Computer.FileSystem.FileExists(kapaFileName) Then

            Try
                appInstance.Workbooks.Open(kapaFileName)
                ok = True

                Try

                    currentWS = CType(appInstance.Worksheets(blattname), Global.Microsoft.Office.Interop.Excel.Worksheet)
                    summenZeile = currentWS.Range("intern_sum").Row
                    Try
                        extSummenZeile = currentWS.Range("extern_sum").Row
                    Catch ex As Exception
                        extSummenZeile = 0
                    End Try

                    tmpDate = CDate(CType(currentWS.Cells(1, spalte), Excel.Range).Value)

                    Do While DateDiff(DateInterval.Month, StartofCalendar, tmpDate) > 0 And _
                            spalte < 241
                        index = getColumnOfDate(tmpDate)
                        tmpKapa = CDbl(CType(currentWS.Cells(summenZeile, spalte), Excel.Range).Value)
                        If extSummenZeile > 0 Then
                            extTmpKapa = CDbl(CType(currentWS.Cells(extSummenZeile, spalte), Excel.Range).Value)
                        Else
                            extTmpKapa = 0.0
                        End If

                        If index <= 240 And index > 0 And tmpKapa >= 0 Then
                            hrole.kapazitaet(index) = tmpKapa
                            hrole.externeKapazitaet(index) = extTmpKapa
                        End If

                        spalte = spalte + 1
                        tmpDate = CDate(CType(currentWS.Cells(1, spalte), Excel.Range).Value)
                    Loop

                Catch ex2 As Exception

                End Try

                appInstance.ActiveWorkbook.Close(SaveChanges:=False)
            Catch ex As Exception

            End Try

        End If


        If formerEE Then
            appInstance.EnableEvents = True
        End If

        If formerSU Then
            appInstance.ScreenUpdating = True
        End If

        enableOnUpdate = True

    End Sub

  

    ''' <summary>
    ''' liest die Name-Mapping Definitionen der Phasen bzw Meilensteine ein
    ''' </summary>
    ''' <param name="ws">Worksheet, in dem die Mappings stehen </param>
    ''' <param name="mappings">Klassen-Instanz, die die Mappings aufnimmt</param>
    ''' <remarks></remarks>
    Friend Sub readNameMappings(ByVal ws As Excel.Worksheet, ByRef mappings As clsNameMapping)

        Dim zeile As Integer, spalte As Integer

        With ws

            ' auslesen der Synonyme und Regular Expressions in Spalte 1, beginnend mit Zeile 3
            Dim ok As Boolean = False
            zeile = 3
            spalte = 1
            If Not IsNothing(CType(.Cells(zeile, spalte), Excel.Range).Value) And _
                Not IsNothing(CType(.Cells(zeile, spalte).offset(0, 1), Excel.Range).Value) Then
                If CStr(.Cells(zeile, spalte).Value).Trim.Length > 0 And _
                    CStr(.Cells(zeile, spalte).offset(0, 1).Value).Trim.Length > 0 Then
                    ok = True
                End If
            End If

            Dim syn As String, stdName As String
            Do While ok

                syn = CStr(.Cells(zeile, spalte).Value).Trim
                stdName = CStr(.Cells(zeile, spalte).offset(0, 1).Value).Trim

                Dim regExpression As String = ""
                Dim isRegExpression As Boolean = False

                If syn.StartsWith("[") And syn.EndsWith("]") Then
                    isRegExpression = True
                    For i As Integer = 1 To syn.Length - 2
                        regExpression = regExpression & syn.Chars(i)
                    Next
                End If

                Try
                    If isRegExpression Then
                        mappings.addRegExpressName(regExpression, stdName)
                    Else
                        mappings.addSynonym(syn, stdName)
                    End If


                Catch ex As Exception

                End Try


                zeile = zeile + 1
                ok = False

                If Not IsNothing(CType(.Cells(zeile, spalte), Excel.Range).Value) And _
                Not IsNothing(CType(.Cells(zeile, spalte).offset(0, 1), Excel.Range).Value) Then
                    If CStr(.Cells(zeile, spalte).Value).Trim.Length > 0 And _
                        CStr(.Cells(zeile, spalte).offset(0, 1).Value).Trim.Length > 0 Then
                        ok = True
                    End If
                End If
            Loop


            ' auslesen der Hierarchies Namens in Spalte 4, beginnend mit Zeile 3
            ok = False
            zeile = 3
            spalte = 4
            If Not IsNothing(CType(.Cells(zeile, spalte), Excel.Range).Value) Then
                If CStr(.Cells(zeile, spalte).Value).Trim.Length > 0 Then
                    ok = True
                End If
            End If

            Dim NameToC As String
            Do While ok

                NameToC = CStr(.Cells(zeile, spalte).Value).Trim

                Try

                    mappings.addNameToComplement(NameToC)

                Catch ex As Exception

                End Try


                zeile = zeile + 1
                ok = False

                If Not IsNothing(CType(.Cells(zeile, spalte), Excel.Range).Value) Then
                    If CStr(.Cells(zeile, spalte).Value).Trim.Length > 0 Then
                        ok = True
                    End If
                End If

            Loop

            ' auslesen der To-Ignore-Names in Spalte 6, beginnend mit Zeile 3
            ok = False
            zeile = 3
            spalte = 6
            If Not IsNothing(CType(.Cells(zeile, spalte), Excel.Range).Value) Then
                If CStr(.Cells(zeile, spalte).Value).Trim.Length > 0 Then
                    ok = True
                End If
            End If

            Dim ignoreName As String
            Do While ok

                ignoreName = CStr(.Cells(zeile, spalte).Value).Trim

                Try

                    mappings.addIgnoreName(ignoreName)

                Catch ex As Exception

                End Try


                zeile = zeile + 1
                ok = False

                If Not IsNothing(CType(.Cells(zeile, spalte), Excel.Range).Value) Then
                    If CStr(.Cells(zeile, spalte).Value).Trim.Length > 0 Then
                        ok = True
                    End If
                End If

            Loop

        End With

    End Sub

    ''' <summary>
    ''' baut die Liste der Darstellungsklassen auf 
    ''' übergeben wird das Excel Worksheet 
    ''' </summary>
    ''' <param name="ws"></param>
    ''' <remarks></remarks>
    Friend Sub aufbauenAppearanceDefinitions(ByVal ws As Excel.Worksheet)

        Dim appDefinition As clsAppearance
        Dim errMsg As String = ""

        With ws

            For Each shp As Excel.Shape In .Shapes
                appDefinition = New clsAppearance
                With appDefinition

                    If shp.Title <> "" Then

                        .name = shp.Title
                        If shp.AlternativeText = "1" Then
                            .isMilestone = True
                        Else
                            .isMilestone = False
                        End If
                        .form = shp

                        Try
                            appearanceDefinitions.Add(.name, appDefinition)
                        Catch ex As Exception
                            errMsg = "Mehrfach Definition in den Darstellungsklassen ... " & vbLf & _
                                         "bitte korrigieren"
                            Throw New Exception(errMsg)
                        End Try


                    End If

                End With


            Next

        End With

    End Sub


    ' '' '' '' '' ''' <summary>
    ' '' '' '' '' ''' kopiert eine sortierte Liste , die Strings enthält
    ' '' '' '' '' ''' </summary>
    ' '' '' '' '' ''' <param name="original"></param>
    ' '' '' '' '' ''' <returns></returns>
    ' '' '' '' '' ''' <remarks></remarks>
    ' '' '' '' ''Public Function copyList(ByVal original As SortedList(Of String, String)) As SortedList(Of String, String)
    ' '' '' '' ''    Dim i As Integer
    ' '' '' '' ''    Dim element As String
    ' '' '' '' ''    Dim kopie As New SortedList(Of String, String)

    ' '' '' '' ''    If Not IsNothing(original) Then
    ' '' '' '' ''        For i = 1 To original.Count
    ' '' '' '' ''            element = CStr(original.Item(i))
    ' '' '' '' ''            If Not kopie.ContainsKey(element) Then
    ' '' '' '' ''                kopie.Add(element, element)
    ' '' '' '' ''            End If

    ' '' '' '' ''        Next
    ' '' '' '' ''    End If
    ' '' '' '' ''    copyList = kopie

    ' '' '' '' ''End Function


    ' '' '' '' '' ''' <summary>
    ' '' '' '' '' ''' kopiert eine sortierte Liste , die Strings enthält
    ' '' '' '' '' ''' </summary>
    ' '' '' '' '' ''' <param name="original"></param>
    ' '' '' '' '' ''' <returns></returns>
    ' '' '' '' '' ''' <remarks></remarks>
    ' '' '' '' ''Public Function copyColltoSortedList(ByVal original As Collection) As SortedList(Of String, String)
    ' '' '' '' ''    Dim i As Integer
    ' '' '' '' ''    Dim element As String
    ' '' '' '' ''    Dim kopie As New SortedList(Of String, String)

    ' '' '' '' ''    If Not IsNothing(original) Then
    ' '' '' '' ''        For i = 1 To original.Count
    ' '' '' '' ''            element = CStr(original.Item(i))
    ' '' '' '' ''            If Not kopie.ContainsKey(element) Then
    ' '' '' '' ''                kopie.Add(element, element)
    ' '' '' '' ''            End If

    ' '' '' '' ''        Next
    ' '' '' '' ''    End If
    ' '' '' '' ''    copyColltoSortedList = kopie

    ' '' '' '' ''End Function

    ' '' '' '' '' ''' <summary>
    ' '' '' '' '' ''' kopiert eine sortierte Liste , die Strings enthält in eine Collection mit Strings
    ' '' '' '' '' ''' </summary>
    ' '' '' '' '' ''' <param name="original"></param>
    ' '' '' '' '' ''' <returns></returns>
    ' '' '' '' '' ''' <remarks></remarks>
    ' '' '' '' ''Public Function copySortedListtoColl(ByVal original As SortedList(Of String, String)) As Collection
    ' '' '' '' ''    Dim i As Integer
    ' '' '' '' ''    Dim element As String
    ' '' '' '' ''    Dim kopie As New Collection

    ' '' '' '' ''    If Not IsNothing(original) Then
    ' '' '' '' ''        For Each kvp As KeyValuePair(Of String, String) In original
    ' '' '' '' ''            element = kvp.Value
    ' '' '' '' ''            If Not kopie.Contains(element) Then
    ' '' '' '' ''                kopie.Add(element, element)
    ' '' '' '' ''            End If

    ' '' '' '' ''        Next
    ' '' '' '' ''    End If
    ' '' '' '' ''    copySortedListtoColl = kopie

    ' '' '' '' ''End Function


    ' '' '' '' ''Public Sub XMLExportReportProfil(ByVal profil As clsReport)



    ' '' '' '' ''    Dim xmlfilename As String = awinPath & ReportProfileOrdner & "\" & profil.name & ".xml"

    ' '' '' '' ''    Try

    ' '' '' '' ''        Dim serializer = New DataContractSerializer(GetType(clsReport))

    ' '' '' '' ''        ' ''Dim xmlstring As String
    ' '' '' '' ''        ' ''Dim sw As New StringWriter()
    ' '' '' '' ''        ' ''Dim writer As New XmlTextWriter(sw)
    ' '' '' '' ''        ' ''writer.Formatting = Formatting.Indented
    ' '' '' '' ''        ' ''serializer.WriteObject(writer, profil)
    ' '' '' '' ''        ' ''writer.Flush()
    ' '' '' '' ''        ' ''xmlstring = sw.ToString()

    ' '' '' '' ''        ' XML-Datei Öffnen
    ' '' '' '' ''        ' A FileStream is needed to write the XML document.

    ' '' '' '' ''        Dim file As New FileStream(xmlfilename, FileMode.Create)
    ' '' '' '' ''        serializer.WriteObject(file, profil)
    ' '' '' '' ''        file.Close()
    ' '' '' '' ''    Catch ex As Exception

    ' '' '' '' ''        Call MsgBox("Beim Schreiben der XML-Datei '" & xmlfilename & "' ist ein Fehler aufgetreten !")

    ' '' '' '' ''    End Try

    ' '' '' '' ''End Sub

    ' '' '' '' ''Public Function XMLImportReportProfil(ByVal profilName As String) As clsReport

    ' '' '' '' ''    Dim profil As New clsReport

    ' '' '' '' ''    Dim serializer = New DataContractSerializer(GetType(clsReport))
    ' '' '' '' ''    Dim xmlfilename As String = awinPath & ReportProfileOrdner & "\" & profilName & ".xml"
    ' '' '' '' ''    Try

    ' '' '' '' ''        ' XML-Datei Öffnen
    ' '' '' '' ''        ' A FileStream is needed to read the XML document.
    ' '' '' '' ''        Dim file As New FileStream(xmlfilename, FileMode.Open)
    ' '' '' '' ''        profil = serializer.ReadObject(file)
    ' '' '' '' ''        file.Close()

    ' '' '' '' ''        XMLImportReportProfil = profil

    ' '' '' '' ''    Catch ex As Exception

    ' '' '' '' ''        Call MsgBox("Beim Lesen der XML-Datei '" & xmlfilename & "' ist ein Fehler aufgetreten !")
    ' '' '' '' ''        XMLImportReportProfil = Nothing
    ' '' '' '' ''    End Try

    ' '' '' '' ''End Function

    ' '' ''Public Sub xmltestwrite1()


    ' '' ''    Dim overview As New clsXMLtest
    ' '' ''    overview.title = "Neues vom XMLTest"
    ' '' ''    overview.reportCalendarVon = Date.Now
    ' '' ''    overview.reportCalendarBis = Date.Now.AddMonths(12)
    ' '' ''    overview.reportIsMpp = True
    ' '' ''    overview.reportPPTTemplate = "templateName"
    ' '' ''    overview.reportVon = Date.Now.AddYears(-2)
    ' '' ''    Dim hstr As String = ""
    ' '' ''    For i = 1 To 5
    ' '' ''        If i = 1 Then
    ' '' ''            hstr = hstr & "name" & i.ToString
    ' '' ''        End If
    ' '' ''        hstr = hstr & ";" & "name" & i.ToString
    ' '' ''    Next
    ' '' ''    'overview.reportPhase = hstr
    ' '' ''    Dim writer As New System.Xml.Serialization.XmlSerializer(GetType(clsXMLtest))
    ' '' ''    Dim file As New System.IO.StreamWriter("\\KOYTEK-NAS\backup\Projekt-Tafel Folder\BHTC\requirements\ReportProfile\xmltest1.xml")
    ' '' ''    writer.Serialize(file, overview)
    ' '' ''    file.Close()

    ' '' ''End Sub

    ' '' ''Public Sub xmltestwrite2()

    ' '' ''    Dim overview As New clsXMLtest
    ' '' ''    overview.title = "Neues vom XMLTest"
    ' '' ''    overview.reportCalendarVon = Date.Now
    ' '' ''    overview.reportCalendarBis = Date.Now.AddMonths(12)
    ' '' ''    overview.reportIsMpp = True
    ' '' ''    overview.reportPPTTemplate = "templateName"
    ' '' ''    overview.reportVon = Date.Now.AddYears(-2)

    ' '' ''    ' ''Dim hstr As String = ""
    ' '' ''    ' ''For i = 1 To 5
    ' '' ''    ' ''    If i = 1 Then
    ' '' ''    ' ''        hstr = hstr & "name" & i.ToString
    ' '' ''    ' ''    End If
    ' '' ''    ' ''    hstr = hstr & ";" & "name" & i.ToString
    ' '' ''    ' ''Next
    ' '' ''    ' ''overview.reportPhase = hstr

    ' '' ''    overview.reportPhasen = New SortedList(Of String, String)
    ' '' ''    For i = 5 To 1 Step -1
    ' '' ''        overview.reportPhasen.Add("name" & i.ToString, "name" & i.ToString)
    ' '' ''    Next

    ' '' ''    Dim serializer = New DataContractSerializer(GetType(clsXMLtest))
    ' '' ''    Dim xmlstring As String
    ' '' ''    Dim sw As New StringWriter()
    ' '' ''    Dim writer As New XmlTextWriter(sw)
    ' '' ''    writer.Formatting = Formatting.Indented
    ' '' ''    serializer.WriteObject(writer, overview)
    ' '' ''    writer.Flush()
    ' '' ''    xmlstring = sw.ToString()
    ' '' ''    ' XML-Datei Öffnen
    ' '' ''    ' A FileStream is needed to read the XML document.

    ' '' ''    Dim file As New FileStream("\\KOYTEK-NAS\backup\Projekt-Tafel Folder\BHTC\requirements\ReportProfile\xmltest2.xml", FileMode.Create)
    ' '' ''    serializer.WriteObject(file, overview)
    ' '' ''    file.Close()

    ' '' ''End Sub

    ' '' ''Public Sub xmltestread2()

    ' '' ''    Dim overview As New clsXMLtest

    ' '' ''    Dim serializer = New DataContractSerializer(GetType(clsXMLtest))

    ' '' ''    ' XML-Datei Öffnen
    ' '' ''    ' A FileStream is needed to read the XML document.
    ' '' ''    Dim file As New FileStream("\\KOYTEK-NAS\backup\Projekt-Tafel Folder\BHTC\requirements\ReportProfile\xmltest2.xml", FileMode.Open)

    ' '' ''    overview = serializer.ReadObject(file)
    ' '' ''    file.Close()

    ' '' ''End Sub

End Module
