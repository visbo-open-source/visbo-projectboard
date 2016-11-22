Imports ProjectBoardDefinitions
Imports System.Globalization
Imports System.Collections.Generic
Imports System.Math
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Core
Imports pptNS = Microsoft.Office.Interop.PowerPoint
Imports System.Xml.Serialization



Public Module Module1


    ' in Modul 1 sollten jetzt alle Konstanten und Einstellungen in einer Klasse zusammengefasst werden
    ' awinSettings: für StartOfCalendar, linker Rand, rechter Rand, ...
    ' Laufzeit Parameter;

    'login - Informationen
    Public dbUsername As String = ""
    Public dbPasswort As String = ""
    Public loginErfolgreich As Boolean = False
    Public noDB As Boolean = True

    Public myWindowsName As String

    Public awinSettings As New clsawinSettings
    Public visboZustaende As New clsVisboZustaende
    Public magicBoardCmdBar As New clsCommandBarEvents
    Public anzahlCalls As Integer = 0
    Public iProjektFarbe As Object
    Public iWertFarbe As Object
    'Public HoehePrcChart As Double


    Public myProjektTafel As String = ""
    Public myCustomizationFile As String
    Public myLogfile As String

    ' gibt an, in welchem Modus sich aktuell die Projekt-Tafe befindet 
    Public currentProjektTafelModus As Integer

    'Definition der Klasse für die ReportMessages ( müssen in awinSettypen gelesen werden aus xml-File)
    Public repMessages As clsReportMessages
   
    
    'Definitionen zum Schreiben eines Logfiles
    Public xlsLogfile As Excel.Workbook = Nothing
    Public logmessage As String = ""
    Public anzFehler As Long = 0

    Public vergleichsfarbe0 As Object
    Public vergleichsfarbe1 As Object
    Public vergleichsfarbe2 As Object
    Public ergebnisfarbe1 As Object
    Public ergebnisfarbe2 As Object

    ' diese Variable steuert, ob die Ereignis-Routine Cmdbar.onupdate durchlaufen wird oder gleich zu beginn wieder verlassen wird
    ' wird immer dann auf false gesetzt , wenn in eigenen Routinen Projekte gesetzt, gelöscht oder ins Show/Noshow gestellt werden 
    Public enableOnUpdate As Boolean = True


    '' MongoDB ist gestartet mongoDBaktiv = true; MongoDB ist unterbrochen mongoDBaktiv=false
    'Public mongoDBaktiv = False

    Public Projektvorlagen As New clsProjektvorlagen
    Public ModulVorlagen As New clsProjektvorlagen
    Public ShowProjekte As New clsProjekte
    Public noShowProjekte As New clsProjekte
    Public selectedProjekte As New clsProjekte
    'Public AlleProjekte As New SortedList(Of String, clsProjekt)
    Public AlleProjekte As New clsProjekteAlle

    Public ImportProjekte As New clsProjekteAlle
    Public projectConstellations As New clsConstellations
    Public currentConstellation As String = "" ' hier wird mitgeführt, was die aktuelle Projekt-Konstellation ist 
    Public allDependencies As New clsDependencies
    Public projectboardShapes As New clsProjektShapes


    ' hier werden die Mapping Informationen abgelegt 
    Public phaseMappings As New clsNameMapping
    Public milestoneMappings As New clsNameMapping


    ' hier wird die Projekt Historie eines Projektes aufgenommen 
    Public projekthistorie As New clsProjektHistorie
    Public specialListofPhases As New Collection

    Public feierTage As New SortedSet(Of Date)

    Public timeMachineIsOn As Boolean = False


    Public PfChartBubbleNames() As String

    Public appearanceDefinitions As New SortedList(Of String, clsAppearance)
    Public RoleDefinitions As New clsRollen
    Public PhaseDefinitions As New clsPhasen
    Public MilestoneDefinitions As New clsMeilensteine


    Public CostDefinitions As New clsKostenarten
    ' Welche Business-Units gibt es ? 
    Public businessUnitDefinitions As New SortedList(Of Integer, clsBusinessUnit)

    ' welche CustomFields gibt es ? 
    Public customFieldDefinitions As New clsCustomFieldDefinitions

    ' wird benötigt, um aufzusammeln und auszugeben, welche Phasen -, Meilenstein Namen  im CustomizationFile noch nicht enthalten sind. 
    Public missingPhaseDefinitions As New clsPhasen
    Public missingMilestoneDefinitions As New clsMeilensteine
    Public missingRoleDefinitions As New clsRollen
    Public missingCostDefinitions As New clsKostenarten

    ' diese Collection nimmt alle Filter Definitionen auf 
    Public filterDefinitions As New clsFilterDefinitions
    Public selFilterDefinitions As New clsFilterDefinitions

    Public DiagramList As New clsDiagramme
    Public awinButtonEvents As New clsAwinEvents




    ' damit ist das Formular Milestone / Status / Phase überall verfügbar
    Public formMilestone As New frmMilestoneInformation
    Public formStatus As New frmStatusInformation
    Public formPhase As New frmPhaseInformation



    ' variable gibt an, zu welchem Objekt-Rolle (Rolle, Kostenart, Ergebnis, ..)  der Röntgen Blick gezeigt wird 
    Public roentgenBlick As New clsBestFitObject

    ' diese beiden folgenden Variablen steuern im Sheet "Ressourcen", welcher Bereich in den Diagrammen angezeigt werden soll
    Public showRangeLeft As Integer
    Public showRangeRight As Integer

    ' diese beiden Variablen nehmen die Farben auf für Showtimezone bzw. Noshowtimezone
    Public showtimezone_color As Object, noshowtimezone_color As Object


    ' maxScreenHeight, maxScreenWidth gibt die maximale Höhe/Breite des Bildschirms in Punkten an 
    Public maxScreenHeight As Double, maxScreenWidth As Double
    Public boxWidth As Double = 19.3, boxHeight As Double, topOfMagicBoard As Double
    Public screen_correct As Double = 0.26
    Public miniWidth As Double = 126 ' wird aber noch in Abhängigkeit von maxscreenwidth gesetzt 
    Public miniHeight As Double = 70 ' wird aber noch in abhängigkeit von maxscreenheight gesetzt

    ' diese Konstante legt den Namen für das Root Element , 1. Phase eines Projektes fest 
    ' das muss mit der calcHryElemKey(".", False) übereinstimmen 
    Public Const rootPhaseName As String = "0§.§"

    Public visboFarbeBlau As Integer = RGB(69, 140, 203)
    Public visboFarbeOrange As Integer = RGB(247, 148, 30)

    ' ur:04.05.2016: da "0§.§" kann in MOngoDB 3.0 nicht in einer sortierten Liste verarbeitet werden (ergibt BsonSerializationException)
    ' also wir rootPhaseName in rootPhaseNameDB geändert nur zum Speichern in DB. Beim Lesen umgekehrt.
    Public Const rootPhaseNameDB As String = "0"

    ' ur:29.06.2016: da "." kann in MOngoDB 3.0 nicht in einer sortierten Liste verarbeitet werden (ergibt BsonSerializationException)
    ' also wird "." = punktName durch "~|°" = punktNameDB  nur zum Speichern in DB ersetzt. Beim Lesen umgekehrt.
    Public Const punktName As String = "."
    Public Const punktNameDB As String = "~|°"

    Public Const minColumns As Integer = 2

    ' diese Konstante legt die Einrücktiefe fest. Das wird benötigt beim Exportieren von Projekte in ein File, ebenso beim Importieren von Datei
    Public Const einrückTiefe As Integer = 2

    ' diese Konstanten werden benötigt, um die Diagramme gemäß des gewählten Zeitraums richtig zu positionieren
    '' ''Public Const summentitel1 As String = "Prognose Ergebniskennzahl"
    '' ''Public Const summentitel2 As String = "strategischer Fit, Risiko & Marge"
    '' ''Public Const summentitel3 As String = "Personal-Kosten intern/extern"
    '' ''Public Const summentitel4 As String = "Personal Kosten Struktur"
    '' ''Public Const summentitel5 As String = "Ergebnis Verbesserungs-Potentiale"
    '' ''Public Const summentitel6 As String = "Bisherige Ziel-Erreichung"
    '' ''Public Const summentitel7 As String = "Prognose zukünftige Ziel-Erreichung"
    '' ''Public Const summentitel8 As String = "Bisherige & zukünftige Ziel-Erreichung"
    '' ''Public Const summentitel9 As String = "Auslastungs-Übersicht"
    '' ''Public Const summentitel10 As String = "Details zur Über-Auslastung"
    '' ''Public Const summentitel11 As String = "Details zur Unter-Auslastung"

    ' diese Variablen werden benötigt, um die Diagramme gemäß des gewählten Zeitraums richtig zu positionieren
    Public summentitel1 As String
    Public summentitel2 As String
    Public summentitel3 As String
    Public summentitel4 As String
    Public summentitel5 As String
    Public summentitel6 As String
    Public summentitel7 As String
    Public summentitel8 As String
    Public summentitel9 As String
    Public summentitel10 As String
    Public summentitel11 As String

   
    Public Const maxProjektdauer As Integer = 60

    ' welche Art von CustomFields gibt es 
    ' kann später ggf erweitert werden auf StrArray, DblArray, etc
    ' muss dann auch in clsProjektVorlage und clsCustomField angepasst werden  
    Public Enum ptCustomFields
        Str = 0
        Dbl = 1
        bool = 2
    End Enum

    Public Enum ptModus
        graficboard = 0
        massEditRessCost = 1
    End Enum

    ' die NAmen für die RPLAN Spaltenüberschriften in Rplan Excel Exports 
    Public Enum ptRplanNamen
        Name = 0
        Anfang = 1
        Ende = 2
        Beschreibung = 3
        Vorgangsklasse = 4
        Produktlinie = 5
        Protocol = 6
        Dauer = 7
    End Enum


    Public Enum PTbubble
        strategicFit = 0
        depencencies = 1
        marge = 2
    End Enum

    Public Enum PTpsel
        alle = -1
        laufend = 0
        lfundab = 1
        abgeschlossen = 2
    End Enum

    ' Enumeration Portfolio Diagramm Kennung 
    Public Enum PTpfdk
        Phasen = 0
        Rollen = 1
        Kosten = 2
        ZieleV = 3
        ZieleF = 4
        FitRisiko = 5
        Auslastung = 6
        UeberAuslastung = 7
        Unterauslastung = 8
        ErgebnisWasserfall = 9
        ComplexRisiko = 10
        ZeitRisiko = 11
        Meilenstein = 12
        AmpelFarbe = 13
        ProjektFarbe = 14
        FitRisikoVol = 15
        Dependencies = 16
        betterWorseL = 17 ' es wird mit dem letzten Stand verglichen
        betterWorseB = 18 ' es wird mit dem Beauftragunsg-Stand verglichen
        Budget = 19
        FitRisikoDependency = 20
    End Enum

    ' immer darauf achten daß die identischen Begriffe PTpfdk und PTprdk auch die gleichen Nummern haben 
    Public Enum PTprdk
        PersonalBalken = 0
        PersonalPie = 1
        KostenBalken = 2
        KostenPie = 3
        Phasen = 4
        StrategieRisiko = 5
        Ergebnis = 6
        ComplexRisiko = 10
        ZeitRisiko = 11
        FitRisikoVol = 15
        Dependencies = 16
    End Enum

    ' projektL bezeichnet die Projekt-Linie , die auch vom Typ mixed ist 
    ' darüber können Abhängigkeites-Connectoren dann auch von Dependency Konnektoren unterschieden werden 
    ' Enumertaion, um in Onupdate, etc. den Typ des Shapes feststellen zu können 
    Public Enum PTshty
        projektN = 0
        projektC = 1
        projektE = 2
        projektL = 3
        phaseN = 4
        phaseE = 5
        phase1 = 6
        milestoneN = 7
        milestoneE = 8
        status = 9
        dependency = 10
        beschriftung = 11
    End Enum

    ' Enumeration History Change Criteria: um anzugeben, welche Veränderung man in der History eines Projektes sucht 

    Public Enum PThcc
        none = 0
        perscost = 1
        othercost = 2
        budget = 3
        ergebnis = 4
        fitrisk = 5
        resultdates = 6
        projektampel = 7
        resultampel = 8
        phasen = 9
        startdatum = 10
        deliverables = 11
        customfields = 12
        projecttype = 13
        endedatum = 14
        persbedarf = 15
        rolle = 16
        kostenart = 17
    End Enum

    ''' <summary>
    ''' betimmt bei den combined Rollen, ob nach allen SubRoles, den Platzhaltern und den Real Rollen aufgelöst werden soll 
    ''' nur nach den Platzhaltern bzw real Rollen 
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum PTcbr
        all = 0
        placeholders = 1
        realRoles = 2
    End Enum

    Public Enum PThis
        current = 0
        vorlage = 1
        beauftragung = 2
        letzterStand = 3
        ersterStand = 4
    End Enum

    ' Enumeration für die Farbe 
    Public Enum PTfarbe
        none = 0
        green = 1
        yellow = 2
        red = 3
    End Enum


    Public Enum PTdpndncy
        none = 0
        schwach = 1
        stark = 3
    End Enum

    Public Enum PTdpndncyType
        none = 0
        inhalt = 1
    End Enum

    Public Enum PTmenue
        visualisieren = 0
        leistbarkeitsAnalyse = 1
        multiprojektReport = 2
        filterdefinieren = 3
        einzelprojektReport = 4
        excelExport = 5
        vorlageErstellen = 6
        rplan = 7
        meilensteinTrendanalyse = 8
        filterAuswahl = 9
        reportBHTC = 10
        sessionFilterDefinieren = 11
    End Enum
    Public Enum PTlicense
        swimlanes = 0
       
    End Enum

    Public Enum PTpptAnnotationType
        text = 0
        datum = 1
    End Enum

    ''' <summary>
    ''' kann überall do verwendet werden, wo es wichtig ist, die CallerApp zu unterscheiden, 
    ''' also wurde ein bestimmtes Formular aus der Multiprojekt Tafel aufgerufen, aus dem Project Add-In oder aus dem Powerpoint Add-In 
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum PTCallerApp
        projektTafel = 0
        projectAddIn = 1
        pptAddIn = 2
    End Enum


    ' wird in awinSetTypen dimensioniert und gesetzt 
    Public portfolioDiagrammtitel() As String

    ' nimmt die Namen der im Zuge der Optimierung automatisch generierten Szenarios auf
    Public autoSzenarioNamen(3) As String


    ' dieser array nimmt die Koordinaten der Formulare auf 
    ' die Koordinaten werden in der Reihenfolge gespeichert: top, left, width, height 
    Public frmCoord(21, 3) As Double

    ' Enumeration Formulare - muss in Korrelation sein mit frmCoord: Dim von frmCoord muss der Anzahl Elemente entsprechen
    Public Enum PTfrm
        timeMachine = 0
        editRess = 1
        noshowBack = 2
        loadC = 3
        storeC = 4
        changeProj = 5
        eingabeProj = 6
        projInfo = 7
        msInfo = 8
        ziele = 9
        auslastung0 = 10
        auslastung1 = 11
        auslastung2 = 12
        zeitraum = 13
        report = 14
        prcChart = 15
        listselP = 16
        listSelR = 17
        listSelM = 18
        phaseInfo = 19
        createVariant = 20
        listInfo = 21
    End Enum

    Public Enum PTpinfo
        top = 0
        left = 1
        width = 2
        height = 3
    End Enum

    ' Sprachen für die ReportMessages
    Public Enum PTSprache
        deutsch = 0
        englisch = 1
        französisch = 2
        spanisch = 3
    End Enum

    ' wird in der Treeview für Laden, Löschen, Aktivieren von TreeView Formularen benötigt 
    Public Enum PTTvActions
        delFromDB = 0
        delFromSession = 1
        loadPVS = 2
        activateV = 3
        definePortfolioDB = 4
        definePortfolioSE = 5
        loadPV = 6
        deleteV = 7
        chgInSession = 8
    End Enum

    ''' <summary>
    ''' alle Bezeichner, die sowohl lesend wie schreibend sind , stehen am Anfang; 
    ''' dann kommen die, die nur lesend sind ... 
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum PTImpExp
        visbo = 0
        rplan = 1
        msproject = 2
        simpleScen = 3
        modulScen = 4
        massenEdit = 5
        addElements = 6
        rplanrxf = 7
    End Enum

    ' SoftwareKomponenten für die Lizensierung
    Public Enum PTSWKomp
        ProjectAdmin = 0
        Swimlanes2 = 1
        SWkomp2 = 2
        SWkomp3 = 3
        SWkomp4 = 4
        Premium = 5
    End Enum

    ' wird in Customization File gesetzt - dies hier ist nur die Default Einstellung 
    ' soll so früh gesetzt sein, damit 
    Public StartofCalendar As Date = #1/1/2000#

    Public weightStrategicFit As Double

    '
    '
    ' Lizenzkomponente kann sein:
    ' ProjectAdmin
    ' Swimlanes2
    Public LizenzKomponenten(5) As String '
    '
    ' Projektstatus kann sein:
    ' beendet
    ' geplant
    ' beauftragt
    ' abgeschlossen
    Public ProjektStatus() As String = {"geplant", "beauftragt", "beauftragt, Änderung noch nicht freigegeben", "beendet", "abgeschlossen"}


    '
    ' aktuell angewendetes ReportProfil
    '
    Public currentReportProfil As New clsReportAll


    '
    'ReportSprache kann sein:
    '   deutsch
    '   englisch
    '   französisch
    '   spanisch
    '
    Public ReportLang() As CultureInfo = {New CultureInfo("de-DE"), _
                                         New CultureInfo("en-US"), _
                                         New CultureInfo("fr-FR"), _
                                         New CultureInfo("es-ES")}
    ' aktuell verwendete Sprache
    '
    Public repCult As CultureInfo

    '
    '
    ' Diagramm-Typ kann sein:
    ' Phase
    ' Rolle
    ' Kostenart
    ' Summe
    ' portfolio

    ' Variable nimmt die Namen der Diagramm-Typen auf 
    Public DiagrammTypen(6) As String

    ' Variable nimmt die Namen der Windows auf  
    Public windowNames(5) As String

    ' Variable nimmt die Namen der Ergebnis Charts auf  
    Public ergebnisChartName(3) As String

    ' diese Variabe nimmt die Farbe der Kapa-Linie an
    Public rollenKapaFarbe As Object

    ' diese Variable nimmt die Farbe der internen Ressourcen, ohne Projekte an auf
    Public farbeInternOP As Object

    ' diese Variable nimmt die Farbe der externen Ressourcen auf
    Public farbeExterne As Object

    ' Variable nimmt die Namen der Worksheets für Portfolio und Ressourcen auf
    Public arrWsNames(0 To 20) As String

    ' variable nimmt auf, wieviel Tage ein Monat hat
    Public nrOfDaysMonth As Double

    ' so werden in Visual Basic die Worksheets der aktuell geladenen Excel Applikation zugänglich gemacht   
    'Public appInstance As _Application
    Public appInstance As Microsoft.Office.Interop.Excel.Application

    Public pptApp As Microsoft.Office.Interop.PowerPoint.Application


    ' nimmt den Pfad Namen auf - also wo liegen Customization File und Projekt-Details
    Public globalPath As String
    Public awinPath As String
    Public importOrdnerNames() As String
    Public exportOrdnerNames() As String
    Public reportOrdnerName As String

    'Public projektFilesOrdner As String = "ProjectFiles"
    'Public rplanimportFilesOrdner As String = "RPLANImport"
    'Public exportFilesOrdner As String = "Export Dateien"

    Public excelExportVorlage As String = "export Vorlage.xlsx"
    Public requirementsOrdner As String = "requirements\"
    Public licFileName As String = requirementsOrdner & "License.xml"
    Public repMsgFileName As String = "ReportTexte"
    Public logFileName As String = requirementsOrdner & "logFile.xlsx"                               ' für Fehlermeldung aus Import und Export
    Public customizationFile As String = requirementsOrdner & "Project Board Customization.xlsx" ' Projekt Tafel Customization.xlsx
    Public cockpitsFile As String = requirementsOrdner & "Project Board Cockpits.xlsx"
    Public projektVorlagenOrdner As String = requirementsOrdner & "ProjectTemplates"
    Public modulVorlagenOrdner As String = requirementsOrdner & "ModuleTemplates"
    Public projektAustausch As String = requirementsOrdner & "Projekt-Steckbrief.xlsx"
    Public projektRessOrdner As String = requirementsOrdner & "Ressource Manager"
    Public RepProjectVorOrdner As String = requirementsOrdner & "ReportTemplatesProject"
    Public RepPortfolioVorOrdner As String = requirementsOrdner & "ReportTemplatesPortfolio"
    Public ReportProfileOrdner As String = requirementsOrdner & "ReportProfile"
    Public demoModusHistory As Boolean = False
    Public historicDate As Date

    Public FirstX As Double = -1.0
    Public FirstY As Double = -1.0
    Public LastX As Double = -1.0
    Public LastY As Double = -1.0
    Public firstPress As Boolean = True

    Public fehlerBeimLoad As Boolean = False










    ''' <summary>
    ''' setzt enableEvents, enableOnUpdate auf true
    ''' </summary>
    ''' <remarks></remarks>
    Sub projektTafelInit()

        With appInstance
            .EnableEvents = True
            If .ScreenUpdating = False Then
                'Call MsgBox ("Screen Update !")
                .ScreenUpdating = True
            End If
        End With


    End Sub


    ''' <summary>
    ''' eingefügt, um eine Warteschleife relisieren zu können ... 
    ''' </summary>
    ''' <param name="dwMilliseconds"></param>
    ''' <remarks></remarks>
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

    'Sub awinLoescheProjekt(pname As String)
    '    '
    '    'Prozedur löscht in Ws Ressourcen alle zeilen, die den Projektnamen enthalten
    '    '
    '    '
    '    'Dim zeile As Integer, endpunkt As Integer
    '    Dim hproj As clsProjekt

    '    Dim tfz As Integer, tfs As Integer
    '    Dim key As String



    '    ' prüfen, ob es in der ShowProjektListe ist ...
    '    If ShowProjekte.contains(pname) Then

    '        ' Shape wird gelöscht - ausserdem wird der Verweis in hproj auf das Shape gelöscht 
    '        Call clearProjektinPlantafel(pname)


    '        Try
    '            hproj = ShowProjekte.getProject(pname)
    '            key = calcProjektKey(hproj)
    '            'Try
    '            '    DeletedProjekte.Add(hproj)
    '            'Catch ex As Exception
    '            '    ' nichts tun, dann wurde das eben schon mal gelöscht ..
    '            'End Try

    '        Catch ex As Exception
    '            Call MsgBox(" Fehler in Delete " & pname & " , Modul: awinLoescheProjekt")
    '            Exit Sub
    '        End Try



    '        With hproj
    '            tfz = .tfZeile
    '            tfs = .tfspalte
    '        End With


    '        ShowProjekte.Remove(pname)
    '        AlleProjekte.Remove(key)


    '        'Dim abstand As Integer ' eigentlich nur Dummy Variable, wird aber in Tabelle2 benötigt ...
    '        'Call awinClkReset(abstand)

    '        ' ein Projekt wurde gelöscht bzw aus Showprojekte entfernt  - typus = 3
    '        Call awinNeuZeichnenDiagramme(3)



    '    Else
    '        Call MsgBox("Projekt " & pname & " wurde nicht gefunden")
    '    End If




    'End Sub

    '
    ' prüft , ob übergebenes Diagramm ein Ergebnis Diagramm ist - in index steht ggf als Ergebnis die entsprechende Nummer; 0 wenn es kein Ergebnis Diagramm ist
    '
    Function istErgebnisDiagramm(ByRef chtobj As ChartObject, ByRef index As Integer) As Boolean
        Dim e As Integer
        Dim found As Boolean
        Dim anzErgebnisArten As Integer = 2
        'Dim chtTitle As String


        e = 1
        index = 0
        found = False

        'Try
        '    chtTitle = chtobj.Chart.ChartTitle.Text
        'Catch ex As Exception
        '    chtTitle = " "
        'End Try

        'While Not found And e <= anzErgebnisArten
        '    If chtTitle Like ergebnisChartName(e - 1) & "*" Then
        '        found = True
        '    Else
        '        e = e + 1
        '    End If
        'End While

        'If found Then
        '    index = e
        'End If

        istErgebnisDiagramm = found

    End Function

    '
    ' prüft , ob übergebenes Diagramm ein Rollen Diagramm ist - in R steht ggf als Ergebnis die entsprechende Rollen-Nummer; 0 wenn es kein Rollen Diagramm ist
    '
    Function istRollenDiagramm(ByRef chtobj As ChartObject) As Boolean

        Dim found As Boolean
        Dim chtobjName As String
        Dim tmpStr(20) As String




        found = False
        chtobjName = chtobj.Name

        Try

            tmpStr = chtobjName.Split(New Char() {CChar("#")}, 20)
            If tmpStr(0) = "pf" And tmpStr.Length >= 2 Then

                If CInt(tmpStr(1)) = PTpfdk.Rollen Then

                    found = True

                End If

            End If

        Catch ex As Exception
        End Try


        istRollenDiagramm = found

    End Function

    '
    ' prüft , ob übergebenes Diagramm ein Cockpit Diagramm ist
    '
    Function istCockpitDiagramm(ByRef chtobj As ChartObject) As Boolean
        Dim ergebnis As Boolean = True

        ' Änderung 31.7 es gibt keine Cockpit Diagramme mehr, deswegen wird immer falsch zurückgegeben 
        'Dim Sc As Microsoft.Office.Interop.Excel.SeriesCollection

        ' Cockpit Diagramme 
        'Sc = chtobj.Chart.SeriesCollection

        'With chtobj
        '    If .Chart.HasAxis(Excel.XlAxisType.xlValue) = False And (Sc.Item(1).ChartType = Excel.XlChartType.xlColumnClustered Or _
        '                                                             Sc.Item(1).ChartType = Excel.XlChartType.xlColumnStacked) And .Width < miniWidth * 1.05 Then
        '        ergebnis = True
        '    ElseIf .Chart.HasLegend = False And Sc.Item(1).ChartType = Excel.XlChartType.xlPie Then
        '        ergebnis = True
        '    Else
        '        ergebnis = False
        '    End If

        'End With

        istCockpitDiagramm = ergebnis



    End Function


    '
    ' prüft , ob übergebenes Diagramm ein Summen Diagramm ist - in rwert steht 1, wenn Rollen Summe, 2, wenn Kosten-Summe
    '
    Function istSummenDiagramm(ByRef chtobj As Excel.ChartObject, ByRef rwert As Integer) As Boolean

        Dim found As Boolean
        Dim chtobjName As String
        Dim tmpStr(20) As String
        Dim vglValue As Integer


        rwert = 0
        found = False


        chtobjName = chtobj.Name

        Try

            tmpStr = chtobjName.Split(New Char() {CChar("#")}, 20)
            If tmpStr(0) = "pf" And tmpStr.Length >= 2 Then


                vglValue = CInt(tmpStr(1))
                If (vglValue >= 3 And vglValue <= 11) Or _
                    (vglValue >= 15 And vglValue <= 19) Then
                    found = True
                    rwert = vglValue
                End If


            End If

        Catch ex As Exception
        End Try

        istSummenDiagramm = found

    End Function

    '
    ' prüft , ob übergebenes Diagramm ein Kosten Diagramm ist - in kostenart steht ggf als Ergebnis die entsprechende Kostenart-Nummer; 0 wenn es kein Kostenart Diagramm ist
    '
    Function istKostenartDiagramm(ByRef chtobj As ChartObject) As Boolean


        Dim found As Boolean
        Dim chtobjName As String
        Dim tmpStr(20) As String


        found = False


        chtobjName = chtobj.Name

        Try

            tmpStr = chtobjName.Split(New Char() {CChar("#")}, 20)
            If tmpStr(0) = "pf" And tmpStr.Length >= 2 Then

                If CInt(tmpStr(1)) = PTpfdk.Kosten Then

                    found = True

                End If

            End If

        Catch ex As Exception
        End Try

        istKostenartDiagramm = found

    End Function

    '
    ' prüft , ob übergebenes Diagramm ein Phasen Diagramm ist - in phasenart steht ggf als Ergebnis die entsprechende Phasen-Nummer; 0 wenn es kein Phasen Diagramm ist
    '
    Function istPhasenDiagramm(ByRef chtobj As ChartObject) As Boolean


        Dim found As Boolean
        Dim chtobjName As String
        Dim tmpStr(20) As String

        found = False


        chtobjName = chtobj.Name

        Try

            tmpStr = chtobjName.Split(New Char() {CChar("#")}, 20)
            If tmpStr(0) = "pf" And tmpStr.Length >= 2 Then

                If CInt(tmpStr(1)) = PTpfdk.Phasen Then

                    found = True

                End If

            End If

        Catch ex As Exception
        End Try

        istPhasenDiagramm = found

    End Function

    '
    ' prüft , ob übergebenes Diagramm ein Meilenstein Diagramm ist - in phasenart steht ggf als Ergebnis die entsprechende Phasen-Nummer; 0 wenn es kein Phasen Diagramm ist
    '
    Function istMileStoneDiagramm(ByRef chtobj As ChartObject) As Boolean


        Dim found As Boolean
        Dim chtobjName As String
        Dim tmpStr(20) As String

        
        found = False


        chtobjName = chtobj.Name

        Try

            tmpStr = chtobjName.Split(New Char() {CChar("#")}, 20)
            If tmpStr(0) = "pf" And tmpStr.Length >= 2 Then

                If CInt(tmpStr(1)) = PTpfdk.Meilenstein Then
                    found = True

                End If

            End If

        Catch ex As Exception
        End Try

        istMileStoneDiagramm = found

    End Function


    '
    ' prüft , ob übergebenes Diagramm ein Rollen Diagramm ist - in R steht ggf als Ergebnis die entsprechende Rollen-Nummer; 0 wenn es kein Rollen Diagramm ist
    '
    Function istPortfolioDiagramm(ByVal chtobj As ChartObject, ByVal portfolio As Integer) As Boolean

        Dim found As Boolean = False

        Dim chtobjName As String
        Dim tmpStr(20) As String


        chtobjName = chtobj.Name

        Try

            tmpStr = chtobjName.Split(New Char() {CChar("#")}, 20)
            If tmpStr(0) = "pf" And tmpStr.Length >= 2 Then

                If CInt(tmpStr(1)) = PTpfdk.FitRisiko Or _
                    CInt(tmpStr(1)) = PTpfdk.FitRisikoVol Or _
                    CInt(tmpStr(1)) = PTpfdk.ComplexRisiko Or _
                    CInt(tmpStr(1)) = PTpfdk.Dependencies Or _
                    CInt(tmpStr(1)) = PTpfdk.ZeitRisiko Then

                    found = True

                End If

            End If

        Catch ex As Exception
        End Try

        istPortfolioDiagramm = found

    End Function

    '
    '
    '
    'Sub awinProjektDefinitionen(ByVal index As Integer)

    '    Dim k As Integer, m As Integer, r As Integer, pnr As Integer
    '    Dim wsnr As Integer
    '    Dim anfang As Integer, ende As Integer
    '    Dim temp_name As String
    '    Dim phaseName As String
    '    Dim chk_phase As Boolean
    '    Dim Zelle As Range
    '    Dim FarbeAktuell As Object
    '    Dim Xwerte() As Double

    '    Dim crole As clsRolle
    '    Dim cphase As New clsPhase
    '    Dim ccost As clsKostenart
    '    'Dim hproj As New clsProjekt
    '    Dim hpv As New clsProjektvorlage
    '    'Dim tstproj As New clsProjekt


    '    If index = 1 Then
    '        wsnr = 5
    '    ElseIf index = 2 Then
    '        wsnr = 6
    '    Else
    '        MsgBox("Fehler in awinProjektDefinitionen !")
    '        Exit Sub
    '    End If

    '    For Each kvp As KeyValuePair(Of String, clsProjektvorlage) In Projektvorlagen.Liste

    '        hpv = kvp.Value

    '        ' hier wird die Farbe des aktuellen Projektes bestimmt ...
    '        FarbeAktuell = hpv.farbe
    '        ' erst sollen die Phasen geprüft werden, dann die Rollen
    '        chk_phase = True

    '        'If index = 1 Then
    '        '    temp_name = hpv.RessourcenDefinitionsBereich
    '        'Else
    '        '    temp_name = hpv.KostenDefinitionsBereich
    '        'End If


    '        pnr = 1

    '        ' hier wird der Bereich ausgelesen - es muss darauf geachtet werden, daß der Bereich lediglich die erste Spalte umfasst, weil das die Anzahl der Schleifen-Durchläufe steuert;
    '        ' für jede Zeile wird entweder die erste Spalte (Phasen-Namen) oder die zweite Spalte (Rollen Name) ausgelesen
    '        ' die Variable chk_phase steuert, ob die erste Spalte (enthält Phasen Namen) oder die zweite Spalte der Zeile (enthält Rollen Namen) ausgelesen wird

    '        If temp_name <> "" Then

    '            For Each Zelle In appInstance.Worksheets(arrWsNames(wsnr)).Range(temp_name)

    '                Select Case chk_phase
    '                    Case True
    '                        ' hier wird die Phasen Information ausgelesen
    '                        If index = 1 Then
    '                            cphase = New clsPhase
    '                            If Len(Zelle.Value) > 0 Then
    '                                phaseName = Zelle.Value

    '                                ' Auslesen der Phasen Dauer
    '                                anfang = 1
    '                                While Zelle.Offset(0, anfang + 1).Interior.Color <> FarbeAktuell
    '                                    anfang = anfang + 1
    '                                End While

    '                                ende = anfang + 1
    '                                While Zelle.Offset(0, ende + 1).Interior.Color = FarbeAktuell
    '                                    ende = ende + 1
    '                                End While
    '                                ende = ende - 1

    '                                chk_phase = False

    '                                With cphase
    '                                    .name = phaseName
    '                                    .relStart = anfang
    '                                    .relEnde = ende
    '                                    .Offset = 0
    '                                End With

    '                            End If

    '                        Else

    '                            chk_phase = False
    '                            cphase = hpv.getPhase(pnr)
    '                            With cphase
    '                                phaseName = .name
    '                                anfang = .relStart
    '                                ende = .relEnde
    '                            End With

    '                        End If



    '                    Case False


    '                        ' hier wird die Rollen bzw Kosten Information ausgelesen

    '                        If Len(Zelle.Offset(0, 1).Value) > 0 Then
    '                            If index = 1 Then
    '                                ' es handelt sich um die Ressourcen Definition
    '                                '
    '                                Try
    '                                    r = RoleDefinitions.getRoledef(Zelle.Offset(0, 1).Value).UID

    '                                    ReDim Xwerte(ende - anfang)
    '                                    For m = anfang To ende
    '                                        Xwerte(m - anfang) = Zelle.Offset(0, m + 1).Value
    '                                    Next m

    '                                    crole = New clsRolle(ende - anfang)
    '                                    With crole
    '                                        .RollenTyp = r
    '                                        .Xwerte = Xwerte
    '                                    End With

    '                                    With cphase
    '                                        .AddRole(crole)
    '                                    End With
    '                                Catch ex As Exception
    '                                    Call MsgBox("kein gültiger Ressourcen-Name: " & _
    '                                                 Zelle.Offset(0, 1).Value)
    '                                End Try



    '                            Else
    '                                ' es handelt sich um die Kostenart Definition
    '                                '
    '                                Try
    '                                    k = CostDefinitions.getCostdef(Zelle.Offset(0, 1).Value).UID

    '                                    ReDim Xwerte(ende - anfang)
    '                                    For m = anfang To ende
    '                                        Xwerte(m - anfang) = Zelle.Offset(0, m + 1).Value
    '                                    Next m

    '                                    ccost = New clsKostenart(ende - anfang)
    '                                    With ccost
    '                                        .KostenTyp = k
    '                                        .Xwerte = Xwerte
    '                                    End With


    '                                    'get Phase pnr
    '                                    With cphase
    '                                        .AddCost(ccost)
    '                                    End With
    '                                Catch ex As Exception
    '                                    Call MsgBox("kein gültiger Name für Kostenart: " & _
    '                                                 Zelle.Offset(0, 1).Value)
    '                                End Try


    '                            End If

    '                        Else
    '                            chk_phase = True

    '                            If index = 1 Then

    '                                hpv.AddPhase(cphase)

    '                            End If

    '                            pnr = pnr + 1

    '                        End If


    '                End Select

    '            Next Zelle
    '        End If


    '    Next kvp
    '    ' End With


    '    ' für Debuggen ...
    '    'If index = 2 Then

    '    'For Each kvp As KeyValuePair(Of String, clsProjekt) In Projektvorlagen.Liste
    '    '    tstproj = kvp.Value
    '    '    For p = 1 To tstproj.CountPhases
    '    '        With tstproj.getPhase(p)
    '    '            For r = 1 To .CountRoles
    '    '                Dim tstrole As New clsRolle
    '    '                Dim chksum As Double
    '    '                tstrole = .getRole(r)
    '    '                chksum = tstrole.summe
    '    '            Next r
    '    '            For k = 1 To .CountCosts
    '    '                Dim tstcost As New clsKostenart
    '    '                Dim chksum As Double
    '    '                tstcost = .getCost(k)
    '    '                chksum = tstcost.summe
    '    '            Next k
    '    '        End With

    '    '    Next p
    '    'Next kvp
    '    'End If


    'End Sub

    'Sub awinClkReset(abstand As Integer)

    '    abstand = 0


    'End Sub



    Sub awinRightClickinPortfolioAendern()
        Dim myBar As CommandBar
        Dim myitem As CommandBarButton
        'Dim myitem As CommandBarControl
        Dim i As Integer, endofsearch As Integer
        Dim found As Boolean
        Dim awinevent As clsEventsPfCharts

        found = False
        i = 1

        With appInstance.CommandBars
            endofsearch = .Count

            While i <= endofsearch And Not found
                If .Item(i).Name = "awinRightClickinPortfolio" Then
                    found = True
                Else
                    i = i + 1
                End If
            End While
        End With

        If found Then
            Exit Sub
        End If

        'CommandBars.Item.Name
        myBar = appInstance.CommandBars.Add(Name:="awinRightClickinPortfolio", Position:=MsoBarPosition.msoBarPopup, Temporary:=True)


        ' Add a menu item
        myitem = CType(myBar.Controls.Add(Type:=MsoControlType.msoControlButton), Microsoft.Office.Core.CommandBarButton)
        With myitem
            .Caption = "Umbenennen"
            .Tag = "Umbenennen"
            '.OnAction = "awinRenameProject"
        End With
        'awinevent = New clsAwinEvent
        'awinevent.Button3Events = myitem
        awinevent = New clsEventsPfCharts
        awinevent.PfChartRightClick = myitem
        awinButtonEvents.Add(awinevent)


        ' Add a menu item
        myitem = CType(myBar.Controls.Add(Type:=MsoControlType.msoControlButton), Microsoft.Office.Core.CommandBarButton)
        With myitem
            .Caption = "Löschen"
            .Tag = "Loesche aus Portfolio"
            '.OnAction = "awinDeleteChartorProject"
        End With
        'awinevent = New clsAwinEvent
        'awinevent.Button3Events = myitem
        awinevent = New clsEventsPfCharts
        awinevent.PfChartRightClick = myitem
        awinButtonEvents.Add(awinevent)

        ' Add a menu item
        myitem = CType(myBar.Controls.Add(Type:=MsoControlType.msoControlButton), Microsoft.Office.Core.CommandBarButton)
        With myitem
            .Caption = "Show / Noshow"
            .Tag = "Show / Noshow"
            '.OnAction = "awinShowNoShowProject"
        End With
        'awinevent = New clsAwinEvent
        'awinevent.Button3Events = myitem
        awinevent = New clsEventsPfCharts
        awinevent.PfChartRightClick = myitem
        awinButtonEvents.Add(awinevent)

        ' Add a menu item
        myitem = CType(myBar.Controls.Add(Type:=MsoControlType.msoControlButton), Microsoft.Office.Core.CommandBarButton)
        With myitem
            .Caption = "Bearbeiten Projekt-Attribute"
            .Tag = "Bearbeiten Projekt-Attribute"
            '.OnAction = "awinEditDataProject"
        End With
        'awinevent = New clsAwinEvent
        'awinevent.Button3Events = myitem
        awinevent = New clsEventsPfCharts
        awinevent.PfChartRightClick = myitem
        awinButtonEvents.Add(awinevent)

        ' Add a menu item
        myitem = CType(myBar.Controls.Add(Type:=MsoControlType.msoControlButton), Microsoft.Office.Core.CommandBarButton)
        With myitem
            .Caption = "Beauftragen"
            .Tag = "Beauftragen"
            '.OnAction = "awinBeauftrageProject"
        End With
        'awinevent = New clsAwinEvent
        'awinevent.Button3Events = myitem
        awinevent = New clsEventsPfCharts
        awinevent.PfChartRightClick = myitem
        awinButtonEvents.Add(awinevent)

    End Sub

    ''' <summary>
    ''' aktiviert die Right Clicks in den Charts 
    ''' </summary>
    ''' <remarks></remarks>
    Sub awinRightClickinPRCCharts()
        Dim myBar As CommandBar
        Dim myitem As CommandBarButton
        Dim i As Integer, endofsearch As Integer
        Dim found As Boolean
        'Dim awinevent As clsAwinEvents
        Dim awinevent As clsEventsPrcCharts

        found = False
        i = 1

        With appInstance.CommandBars
            endofsearch = .Count

            While i <= endofsearch And Not found
                If .Item(i).Name = "awinRightClickinPRCChart" Then
                    found = True
                Else
                    i = i + 1
                End If
            End While
        End With

        If found Then
            Exit Sub
        End If

        'CommandBars.Item.Name
        myBar = appInstance.CommandBars.Add(Name:="awinRightClickinPRCChart", Position:=MsoBarPosition.msoBarPopup, Temporary:=True)


        ' Add a menu item
        myitem = CType(myBar.Controls.Add(Type:=MsoControlType.msoControlButton), Microsoft.Office.Core.CommandBarButton)
        With myitem
            .Caption = "Löschen"
            .Tag = "Löschen"
            '.OnAction = "awinDeleteChartorProject"
        End With

        'awinevent = New clsAwinEvent
        'awinevent.Button4Events = myitem
        awinevent = New clsEventsPrcCharts
        awinevent.PrcChartRightClick = myitem
        awinButtonEvents.Add(awinevent)

        ' Add a menu item
        myitem = CType(myBar.Controls.Add(Type:=MsoControlType.msoControlButton), Microsoft.Office.Core.CommandBarButton)
        With myitem
            .Caption = "Röntgenblick ein/aus"
            .Tag = "Bedarf anzeigen"
            '.OnAction = "awinShowNeedsOfProjects"
        End With
        'awinevent = New clsAwinEvent
        'awinevent.Button4Events = myitem
        awinevent = New clsEventsPrcCharts
        awinevent.PrcChartRightClick = myitem
        awinButtonEvents.Add(awinevent)

        ' Add a menu item
        myitem = CType(myBar.Controls.Add(Type:=MsoControlType.msoControlButton), Microsoft.Office.Core.CommandBarButton)
        With myitem
            .Caption = "nach Freiheitsgraden optimieren"
            .Tag = "Optimieren"
            '.OnAction = "awinOptimizeStartOfProjects"
        End With
        'awinevent = New clsAwinEvent
        'awinevent.Button4Events = myitem
        awinevent = New clsEventsPrcCharts
        awinevent.PrcChartRightClick = myitem
        awinButtonEvents.Add(awinevent)

        ' ergänzt am 2.11.2014
        ' Add a menu item
        myitem = CType(myBar.Controls.Add(Type:=MsoControlType.msoControlButton), Microsoft.Office.Core.CommandBarButton)
        With myitem
            .Caption = "nach Varianten optimieren"
            .Tag = "Varianten optimieren"
            '.OnAction = "awinOptimizeStartOfProjects"
        End With
        'awinevent = New clsAwinEvent
        'awinevent.Button4Events = myitem
        awinevent = New clsEventsPrcCharts
        awinevent.PrcChartRightClick = myitem
        awinButtonEvents.Add(awinevent)

    End Sub

    Sub awinKontextReset()

        Try
            appInstance.CommandBars("awinRightClickinPortfolio").Delete()
        Catch ex As Exception

        End Try

        Try
            appInstance.CommandBars("awinRightClickinPRCChart").Delete()
        Catch ex As Exception

        End Try


        ' die Short Cut Menues aus Excel wieder alle aktivieren ...
        'Dim cbar As CommandBar

        'For Each cbar In appInstance.CommandBars

        '    cbar.Enabled = True
        '    'Try
        '    '    cbar.Reset()
        '    'Catch ex As Exception

        '    'End Try

        'Next


    End Sub



    '

    '
    ''' <summary>
    ''' gibt die Überdeckung zurück zwischen den beiden Zeiträumen definiert durch showRangeLeft /showRangeRight und anfang / ende
    ''' </summary>
    ''' <param name="anfang">Anfang Zeitraum 2</param>
    ''' <param name="ende">Ende Zeitraum 2</param>
    ''' <param name="ixZeitraum">gibt an , in welchem Monat des Zeitraums die Überdeckung anfängt: 0 = 1. Monat</param>
    ''' <param name="ix">gibt an, in welchem Monat des durch Anfang / ende definierten Zeitraums die Überdeckung anfängt</param>
    ''' <param name="anzahl">enthält die Breite der Überdeckung</param>
    ''' <remarks></remarks>
    Sub awinIntersectZeitraum(anfang As Integer, ende As Integer, _
                                    ByRef ixZeitraum As Integer, ByRef ix As Integer, ByRef anzahl As Integer)



        If istBereichInTimezone(anfang, ende) Then
            If anfang <= showRangeLeft Then
                ixZeitraum = 0
                ix = showRangeLeft - anfang
                If ende >= showRangeRight Then
                    anzahl = showRangeRight - showRangeLeft + 1
                Else
                    anzahl = ende - showRangeLeft + 1
                End If
            Else
                ixZeitraum = anfang - showRangeLeft
                ix = 0
                If ende >= showRangeRight Then
                    anzahl = showRangeRight - anfang + 1
                Else
                    anzahl = ende - anfang + 1
                End If
            End If
        Else
            anzahl = 0
        End If


    End Sub

    '
    ' löscht alle Cockpit Charts
    '
    Sub awinLoescheCockpitCharts()
        Dim i As Integer
        Dim chtobj As Excel.ChartObject

        With CType(appInstance.Workbooks.Item("Projectboard.xlsx").Worksheets(arrWsNames(3)), Excel.Worksheet)

            For Each chtobj In CType(.ChartObjects, Excel.ChartObjects)
                If istCockpitDiagramm(chtobj) Then
                    chtobj.Delete()
                End If
            Next chtobj

            i = 1

            While i <= DiagramList.Count
                If DiagramList.getDiagramm(i).isCockpitChart Then
                    DiagramList.Remove(i)
                Else
                    i = i + 1
                End If
            End While

        End With


    End Sub

    ''' <summary>
    ''' löscht alle Cockpit Charts, die vom Typ DiagrammTypen(prctyp) sind)
    ''' </summary>
    ''' <param name="prctyp"></param>
    ''' <remarks></remarks>
    Sub awinLoescheCockpitCharts(ByVal prctyp As Integer)
        Dim i As Integer
        Dim chtobj As Excel.ChartObject
        Dim chtTitle As String

        ' finde alle Charts, die Cockpit Chart sind und vom Typ her diagrammtypen(prctyp)

        With appInstance.Workbooks.Item("Projectboard.xlsx").Worksheets(arrWsNames(3))
            Dim found As Boolean
            For Each chtobj In CType(.ChartObjects, Excel.ChartObjects)
                Try
                    chtTitle = chtobj.Chart.ChartTitle.Text
                Catch ex As Exception
                    chtTitle = " "
                End Try

                If istCockpitDiagramm(chtobj) Then
                    found = False
                    i = 1
                    While i <= DiagramList.Count And Not found
                        'If (chtTitle Like (DiagramList.getDiagramm(i).DiagrammTitel & "*")) And _
                        If (chtTitle = DiagramList.getDiagramm(i).DiagrammTitel) And _
                                        (DiagramList.getDiagramm(i).isCockpitChart = True) And _
                                        (DiagramList.getDiagramm(i).diagrammTyp = DiagrammTypen(prctyp)) Then
                            DiagramList.Remove(i)
                            chtobj.Delete()
                            found = True
                        Else
                            i = i + 1
                        End If
                    End While
                End If
            Next chtobj

        End With


    End Sub

    Sub awinLoescheChartsAtPosition(ByVal left As Double)

        Dim chtobj As Excel.ChartObject
        Dim tstLeft As Double
        Dim tmpArray() As String

        ' finde alle Charts, die bei left platziert sind ... 



        With appInstance.Workbooks.Item("Projectboard.xlsx").Worksheets(arrWsNames(3))

            For Each chtobj In CType(.ChartObjects, Excel.ChartObjects)

                tmpArray = chtobj.Name.Split(New Char() {CType("#", Char)}, 5)

                Try
                    tstLeft = chtobj.Left
                Catch ex As Exception
                    tstLeft = -10
                End Try

                Try
                    If System.Math.Abs(tstLeft - left) < 5 And tmpArray(0) = "pf" Then
                        chtobj.Delete()
                    End If
                Catch ex As Exception

                End Try

            Next chtobj

        End With

    End Sub
    Function TypOfCockpitChart(ByRef chtobj As ChartObject) As Integer
        Dim chtTitle As String
        Dim found As Boolean
        Dim i As Integer


        Dim ergebnis As Integer = -1

        Try
            chtTitle = chtobj.Chart.ChartTitle.Text
        Catch ex As Exception
            chtTitle = " "
        End Try



        found = False
        i = 1
        While i <= DiagramList.Count And Not found
            'If (chtTitle Like (DiagramList.getDiagramm(i).DiagrammTitel & "*")) And _
            If (chtTitle = DiagramList.getDiagramm(i).DiagrammTitel) And _
                            (DiagramList.getDiagramm(i).isCockpitChart) Then
                With DiagramList.getDiagramm(i)
                    Select Case .diagrammTyp
                        Case DiagrammTypen(0)
                            ergebnis = 0
                        Case DiagrammTypen(1)
                            ergebnis = 1
                        Case DiagrammTypen(2)
                            ergebnis = 2
                        Case DiagrammTypen(3)
                            ergebnis = 3
                        Case DiagrammTypen(4)
                            ergebnis = 4
                    End Select
                End With
                found = True
            Else
                i = i + 1
            End If
        End While

        TypOfCockpitChart = ergebnis

    End Function

    '
    ' istinStringCollection(hproj.name, TypeCollection)
    '
    Function istinStringCollection(ByRef suchbegriff As String, ByRef myCollection As Collection) As Boolean
        Dim i As Integer
        Dim found As Boolean

        found = False
        i = 1
        While i <= myCollection.Count And Not found
            If myCollection.Item(i) = suchbegriff Then
                found = True
            Else
                i = i + 1
            End If
        End While

        istinStringCollection = found
    End Function

    ''' <summary>
    ''' Selektion aller Objekte in der Liste "selectedProjekte"
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Sub awinSelect()

        Dim worksheetShapes As Excel.Shapes
        Dim hproj As clsProjekt
        Dim shapegruppe As Excel.ShapeRange
        Dim shpElement As Excel.Shape
        Dim shpArray() As String
        Dim i As Integer = 0

        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False

        ' Selektierte Projekte als selektiert kennzeichnen in der ProjektTafel

        If selectedProjekte.Count > 0 Then
            worksheetShapes = CType(appInstance.Workbooks.Item("Projectboard.xlsx").Worksheets(arrWsNames(3)), Excel.Worksheet).Shapes
            ReDim shpArray(selectedProjekte.Count - 1)

            For Each kvp In selectedProjekte.Liste

                hproj = kvp.Value
                i = i + 1
                Try
                    shpElement = CType(appInstance.Workbooks.Item("Projectboard.xlsx").Worksheets(arrWsNames(3)), Excel.Worksheet).Shapes.Item(hproj.name)
                    shpArray(i - 1) = shpElement.Name

                Catch ex As Exception

                End Try

            Next
            shapegruppe = worksheetShapes.Range(shpArray)
            shapegruppe.Select()
        End If



        appInstance.EnableEvents = formerEE

    End Sub
    ''' <summary>
    ''' De-Selektion aller Objekte durch Selektion einer Zelle in Zeile 2 in der Mitte des aktuell gezeigten Fensters 
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Sub awinDeSelect()
        Dim srow As Integer = 1
        Dim hziel As Integer
        Dim vziel As Integer


        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False

        ' Selektierte Projekte auf Null setzen 

        If selectedProjekte.Count > 0 Then
            selectedProjekte.Clear()
            If awinSettings.showValuesOfSelected Then
                Call awinNeuZeichnenDiagramme(8)
            End If

        End If



        '
        ' das folgende selektiert die Zelle in der Mitte des aktuell gezeigten Fensters
        ' das verhindert, daß sich plötzlich der Fenster Ausschnitt verändert
        '
        Try
            With appInstance.ActiveWindow
                hziel = CInt((.VisibleRange.Left + .VisibleRange.Width / 2) / boxWidth)
                vziel = CInt((.VisibleRange.Top + .VisibleRange.Height / 2) / boxHeight)
                If vziel < 2 Then
                    vziel = 2
                End If
            End With

            With appInstance.ActiveSheet
                '.Cells(2, hziel).Select()
                .Cells(vziel, hziel).Select()
            End With
        Catch ex As Exception

            With appInstance.ActiveSheet
                .Cells(2, 20).Select()
            End With

        End Try



        appInstance.EnableEvents = formerEE

    End Sub

    Public Function magicBoardZeileIstFrei(ByVal zeile As Integer) As Boolean
        Dim istfrei = True
        Dim ix As Integer = 1
        Dim anzahlP As Integer = ShowProjekte.Count
        Dim tmpCollection As New Collection

        If zeile >= 2 Then

            For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

                With kvp.Value
                    If zeile >= .tfZeile And zeile < .tfZeile + kvp.Value.calcNeededLines(tmpCollection, tmpCollection, awinSettings.drawphases Or kvp.Value.extendedView, False) Then
                        istfrei = False
                        Exit For
                    End If
                End With

            Next

        Else

            istfrei = False

        End If

        magicBoardZeileIstFrei = istfrei
    End Function



    Public Function magicBoardIstFrei(ByVal mycollection As Collection, ByVal pname As String, ByVal zeile As Integer, _
                                      ByVal spalte As Integer, ByVal laenge As Integer, ByVal anzahlZeilen As Integer) As Boolean
        Dim istfrei = True
        Dim ix As Integer = 1
        Dim anzahlP As Integer = ShowProjekte.Count


        For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

            If pname <> kvp.Key And Not mycollection.Contains(kvp.Key) And kvp.Value.shpUID <> "" Then
                With kvp.Value
                    If .tfZeile >= zeile And .tfZeile <= zeile + anzahlZeilen - 1 Then
                        If spalte <= .tfspalte Then
                            If spalte + laenge - 1 >= .tfspalte Then
                                istfrei = False
                                Exit For
                            End If
                        ElseIf spalte <= .tfspalte + .anzahlRasterElemente - 1 Then
                            istfrei = False
                            Exit For
                        End If
                    End If
                End With
            End If

        Next
        magicBoardIstFrei = istfrei
    End Function

    Public Function findeMagicBoardPosition(ByVal mycollection As Collection, ByVal pname As String, ByVal zeile As Integer, ByVal spalte As Integer, ByVal laenge As Integer) As Integer
        Dim lookDown As Boolean = True
        Dim tryoben As Integer, tryunten As Integer
        Dim anzahlzeilen As Integer
        Dim tmpCollection As New Collection


        Try
            Dim hproj As clsProjekt = ShowProjekte.getProject(pname)
            anzahlzeilen = hproj.calcNeededLines(tmpCollection, tmpCollection, hproj.extendedView Or awinSettings.drawphases, False)

            ' Konsistenzbedingung prüfen ... 
            If zeile < 2 Then
                zeile = 2
            End If

            'If mycollection.Count = 0 Then
            '    mycollection.Add(pname, pname)
            'End If

            If Not magicBoardIstFrei(mycollection, pname, zeile, spalte, laenge, anzahlzeilen) Then
                tryoben = zeile - 1
                tryunten = zeile + 1

                ' jetzt ggf eine neue Position für das Shape suchen - dabei iterierend unten bzw oben suchen
                zeile = tryunten
                lookDown = True

                While Not magicBoardIstFrei(mycollection, pname, zeile, spalte, laenge, anzahlzeilen)
                    'lookDown = Not lookDown
                    If lookDown Then
                        tryunten = tryunten + 1
                        zeile = tryunten
                    Else
                        tryoben = tryoben - 1
                        If tryoben < 2 Then
                            tryunten = tryunten + 1
                            zeile = tryunten
                        Else
                            zeile = tryoben
                        End If
                    End If
                End While
            End If
        Catch ex As Exception

        End Try


        findeMagicBoardPosition = zeile

    End Function

    Sub awinSubtest()

        Call MsgBox("del gedrückt ...")
    End Sub

   


    ''' <summary>
    ''' löscht die Symbole - je nach auswahl 
    ''' </summary>
    ''' <param name="auswahl">
    ''' 0=alle
    ''' 1=nur meilensteine
    ''' 2=nur Status
    ''' 3=nur Phasen</param>
    ''' <remarks></remarks>
    Public Sub awinDeleteProjectChildShapes(ByVal auswahl As Integer)

        Dim worksheetShapes As Excel.Shapes
        Dim shpElement As Excel.Shape
        Dim shapeType As Integer

        Dim typCollection As New Collection

        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim formereO As Boolean = enableOnUpdate
        appInstance.EnableEvents = False
        enableOnUpdate = False


        Select Case auswahl
            Case 0
                formMilestone.Visible = False
                formStatus.Visible = False
                formPhase.Visible = False

                typCollection.Add(CInt(PTshty.milestoneN).ToString, CInt(PTshty.milestoneN).ToString)
                typCollection.Add(CInt(PTshty.phaseN).ToString, CInt(PTshty.phaseN).ToString)

            Case 1
                formMilestone.Visible = False
                typCollection.Add(CInt(PTshty.milestoneN).ToString, CInt(PTshty.milestoneN).ToString)

            Case 2
                formStatus.Visible = False

            Case 3
                formPhase.Visible = False
                typCollection.Add(CInt(PTshty.phaseN).ToString, CInt(PTshty.phaseN).ToString)

            Case Else
                appInstance.EnableEvents = formerEE
                enableOnUpdate = formereO
                Exit Sub
        End Select

        Try
            worksheetShapes = CType(appInstance.Workbooks.Item("Projectboard.xlsx").Worksheets(arrWsNames(3)), Excel.Worksheet).Shapes



            For Each shpElement In worksheetShapes

                shapeType = kindOfShape(shpElement)

                ' neu 

                If isProjectType(shapeType) And shpElement.AutoShapeType = MsoAutoShapeType.msoShapeMixed Then

                    projectboardShapes.removeChildsOfType(shpElement, typCollection)

                ElseIf shapeType = PTshty.status Then
                    projectboardShapes.remove(shpElement)
                End If

                ' Ende neu 

            Next
        Catch ex As Exception

        End Try

        appInstance.EnableEvents = formerEE
        enableOnUpdate = formereO

    End Sub

    ''' <summary>
    ''' löscht die Beschriftungen in der Projekt-Tafel 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub deleteBeschriftungen(Optional ByVal pName As String = "")
        ' jetzt werden die Aktionen gemacht 
        Dim worksheetShapes As Excel.Shapes
        Dim shpElement As Excel.Shape

        Dim descriptionShapeName As String = "Description#" & pName

        enableOnUpdate = False


        Try
            worksheetShapes = CType(appInstance.Workbooks.Item("Projectboard.xlsx").Worksheets(arrWsNames(3)), Excel.Worksheet).Shapes

            If pName = "" Then
                For Each shpElement In worksheetShapes

                    If shpElement.AlternativeText = CInt(PTshty.beschriftung).ToString Then
                        shpElement.Delete()
                    End If

                Next
            Else
                Try
                    shpElement = worksheetShapes.Item(descriptionShapeName)
                    If Not IsNothing(shpElement) Then
                        shpElement.Delete()
                    End If
                Catch ex As Exception

                End Try


            End If

        Catch ex As Exception
            Call MsgBox(ex.Message)
        End Try



        enableOnUpdate = True
    End Sub


    ''' <summary>
    ''' löscht zu dem angegebenen Shape die Child Shapes Milestone, Phase oder Status 
    ''' </summary>
    ''' <param name="pShape"></param>
    ''' <param name="auswahl">
    ''' 0=alle
    ''' 1=nur meilensteine
    ''' 2=nur Status
    ''' 3=nur Phasen</param>
    ''' <remarks></remarks>
    Public Sub awinDeleteProjectChildShapes(ByVal pShape As Excel.Shape, ByVal auswahl As Integer)

        Dim shapeType As Integer
        Dim typCollection As New Collection

        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim formereO As Boolean = enableOnUpdate
        appInstance.EnableEvents = False
        enableOnUpdate = False


        Select Case auswahl
            Case 0

                typCollection.Add(CInt(PTshty.milestoneN).ToString, CInt(PTshty.milestoneN).ToString)
                typCollection.Add(CInt(PTshty.phaseN).ToString, CInt(PTshty.phaseN).ToString)

            Case 1

                typCollection.Add(CInt(PTshty.milestoneN).ToString, CInt(PTshty.milestoneN).ToString)

            Case 2

            Case 3

                typCollection.Add(CInt(PTshty.phaseN).ToString, CInt(PTshty.phaseN).ToString)

            Case Else
                appInstance.EnableEvents = formerEE
                enableOnUpdate = formereO
                Exit Sub
        End Select

        Try
            

            shapeType = kindOfShape(pShape)

                ' neu 

            If isProjectType(shapeType) And pShape.AutoShapeType = MsoAutoShapeType.msoShapeMixed Then

                projectboardShapes.removeChildsOfType(pShape, typCollection)

            ElseIf shapeType = PTshty.status Then
                projectboardShapes.remove(pShape)
            End If

                ' Ende neu 


        Catch ex As Exception

        End Try

        appInstance.EnableEvents = formerEE
        enableOnUpdate = formereO

    End Sub


    ''' <summary>
    ''' Sub berechnet die neuen Werte so, daß die Charakterisitik der Werte möglichst erhalten bleibt 
    ''' Übergeben wird die neue Länge - es wird dann entschieden, welche Charakteristik am ehesten zutrifft - danach werden die Werte neu bestimmt
    ''' newlength ist die echte länge, also z.Bsp steht 2 für 2 Monate 
    ''' changeProp gibt an, ob die Werte proportional zur Verkürzung / Verlängerung geändert werden sollen 
    ''' oder ob die Gesamt Summe konstant bleibt und einfach neu verteilt wird 
    ''' </summary>
    ''' <param name="newLength"></param>
    ''' <param name="bedarf">der bisherige Array mit den Werten</param>
    ''' <param name="changeProp">
    ''' true: es soll proportional verändert werden 
    ''' false: Gesamt Summe bleibt konstant - wird nur anders aufgeteilt
    ''' </param>
    ''' <remarks></remarks>

    Public Function adjustArrayLength(ByVal newLength As Integer, ByVal bedarf() As Double, ByVal changeProp As Boolean) As Double()
        Dim oldLength As Integer
        Dim oldSum As Double, newSum As Double
        Dim avg As Double
        Dim min As Double, max As Double

        Dim newValues() As Double
        Dim typus As Integer

        Dim ix As Integer
        


        Try
            ReDim newValues(newLength - 1)
            oldLength = bedarf.Length
            avg = bedarf.Sum / oldLength
            min = bedarf.Min
            max = bedarf.Max
        Catch ex As Exception
            Throw New ArgumentException("Fehler bei Adjust Array Length ...")
        End Try




        If newLength = oldLength Then
            ' wenn keine Änderung vorzunehmen ist ... 

            newValues = bedarf

        Else

            oldSum = bedarf.Sum

            If changeProp Then
                ' ändere proportional 
                newSum = newLength / oldLength * oldSum
            Else
                ' behalte die Werte
                newSum = oldSum
            End If


            typus = definecharacteristics(bedarf)


            Dim ixi As Integer

            Select Case typus
                Case 1

                    ' aufsteigend von klein zu groß 
                    ' es wird der neue Array einfach von hinten her aufgefüllt 
                    ix = 0
                    ixi = newLength - 1
                    Do While ix <= newSum

                        newValues(ixi) = newValues(ixi) + 1
                        If ixi = 0 Then
                            ixi = newLength - 1
                        Else
                            ixi = ixi - 1
                        End If

                        ix = ix + 1

                    Loop

                    If ix < newSum Then
                        newValues(newLength - 1) = newValues(newLength - 1) + newSum - ix
                    End If


                Case 2
                    ' gleich bzw Buckel Funktion - aktuell wie aufsteigend, aber beginnend in der Mitte  
                    ix = 0
                    ixi = CInt(newLength / 2)
                    Do While ix <= newSum

                        newValues(ixi) = newValues(ixi) + 1
                        If ixi = 0 Then
                            ixi = newLength - 1
                        Else
                            ixi = ixi - 1
                        End If

                        ix = ix + 1

                    Loop

                    If ix < newSum Then
                        newValues(CInt(newLength / 2)) = newValues(CInt(newLength / 2)) + newSum - ix
                    End If


                Case 3
                    ' absteigend von groß zu klein
                    Do While ix <= newSum

                        newValues(ixi) = newValues(ixi) + 1
                        If ixi = newLength - 1 Then
                            ixi = 0
                        Else
                            ixi = ixi + 1
                        End If

                        ix = ix + 1

                    Loop

                    If ix < newSum Then
                        newValues(0) = newValues(0) + newSum - ix
                    End If

            End Select


        End If

        adjustArrayLength = newValues

    End Function


    ''' <summary>
    ''' bestimmt die Charakteristik des Verlaufs: 
    ''' 1-minimum vorne, max hinten -  steigender Verlauf
    ''' 2-Max in der Mitte bzw. einigermaßen konstanter Verlauf
    ''' 3-max vorne, min hinten -  fallender Verlauf
    ''' </summary>
    Public Function definecharacteristics(ByVal Bedarf() As Double) As Integer

        Dim min As Double
        Dim max As Double
        Dim avg As Double

        Dim bereich As Integer
        Dim i As Integer
        Dim minvorne As Boolean = False, minhinten As Boolean = False, _
            maxvorne As Boolean = False, maxhinten As Boolean = False


        ' Festsetzungen 
        Try
            min = Bedarf.Min
            max = Bedarf.Max
            avg = Bedarf.Sum / Bedarf.Length
            bereich = CInt(Bedarf.Length / 4)
        Catch ex As Exception
            Throw New ArgumentException("Fehler ... Bedarf kein Arraey von zahlen ? ")
        End Try


        For i = 0 To bereich
            If Bedarf(i) = min Then
                minvorne = True
            ElseIf Bedarf(i) = max Then
                maxvorne = True
            End If
        Next i

        For i = Bedarf.Length - (bereich + 1) To Bedarf.Length - 1
            If Bedarf(i) = min Then
                minhinten = True
            ElseIf Bedarf(i) = max Then
                maxhinten = True
            End If
        Next

        If minvorne And maxhinten Then
            definecharacteristics = 1
        ElseIf maxvorne And minhinten Then
            definecharacteristics = 3
        Else
            definecharacteristics = 2
        End If

    End Function

    ''' <summary>
    ''' Funktion berechnet die Dauer in Tagen des Zeitraums, der durch startDatum und endeDatum aufgespannt wird 
    ''' Wenn StartDatum = EndeDatum: Dauer = 1
    ''' Wenn StartDatum nach dem EndeDatum liegt, wird eine negative Dauer ausgegegeben   
    ''' </summary>
    ''' <param name="startDatum"></param>
    ''' <param name="endeDatum"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function calcDauerIndays(ByVal startDatum As Date, ByVal endeDatum As Date) As Integer

        If startDatum.Date > endeDatum.Date Then
            calcDauerIndays = CInt(DateDiff(DateInterval.Day, startDatum.Date, endeDatum.Date) - 1)
        Else
            calcDauerIndays = CInt(DateDiff(DateInterval.Day, startDatum.Date, endeDatum.Date) + 1)
        End If

    End Function

    ''' <summary>
    ''' Funktion berechnet die Dauer in Tagen des Zeitraums, der durch StartDatum und Dauer in Monaten aufgespannt wird 
    ''' wenn isRelative=false, dann steht rasterMonat für die absolute Spalte der Projekt-Tafel, in der das Projekt endet
    ''' </summary>
    ''' <param name="startDatum"></param>
    ''' <param name="rasterMonat"></param>
    ''' <param name="isRelative"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function calcDauerIndays(ByVal startDatum As Date, ByVal rasterMonat As Integer, ByVal isRelative As Boolean) As Integer
        Dim endeDatum As Date

        If isRelative Then
            If rasterMonat >= 0 Then
                endeDatum = StartofCalendar.AddMonths(getColumnOfDate(startDatum) - 1 + rasterMonat).AddDays(-1)
            Else
                endeDatum = StartofCalendar.AddMonths(getColumnOfDate(startDatum) - 1 + rasterMonat)
            End If

        Else

            If rasterMonat >= 0 Then
                endeDatum = StartofCalendar.AddMonths(rasterMonat).AddDays(-1)
            Else
                endeDatum = StartofCalendar.AddMonths(getColumnOfDate(startDatum) - 1 + rasterMonat)
            End If

        End If

        If startDatum.Date > endeDatum.Date Then
            calcDauerIndays = CInt(DateDiff(DateInterval.Day, startDatum, endeDatum) - 1)
        Else
            calcDauerIndays = CInt(DateDiff(DateInterval.Day, startDatum, endeDatum) + 1)
        End If


    End Function

    Public Function calcDatum(ByVal datum As Date, ByVal dauerInDays As Integer) As Date

        If dauerInDays > 0 Then
            calcDatum = datum.AddDays(dauerInDays - 1)
        ElseIf dauerInDays < 0 Then
            calcDatum = datum.AddDays(dauerInDays + 1)
        Else
            Throw New Exception("Dauer von Null ist unzulässig ..")
        End If

    End Function

    ''' <summary>
    ''' gibt einen String zurück, der den dem level entsprechenden Indent an Leerzeichen enthält  
    ''' bei level = -1 wird "???" als String zurückgegeben 
    ''' </summary>
    ''' <param name="level"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function erzeugeIndent(ByVal level As Integer) As String
        Dim indentDelta As String = "   "
        Dim tmpStr As String = ""

        If level = -1 Then
            tmpStr = "???"
        Else
            For i As Integer = 1 To level
                tmpStr = tmpStr & indentDelta
            Next
        End If

        erzeugeIndent = tmpStr

    End Function

    ''' <summary>
    ''' berechnet den "ersten" Namen, der in der sortedList der Hierarchie auftreten würde 
    ''' </summary>
    ''' <param name="elemName"></param>
    ''' <param name="isMilestone"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function calcHryElemKey(ByVal elemName As String, ByVal isMilestone As Boolean, Optional ByVal lfdNr As Integer = 0) As String

        Dim elemKey As String
        Dim elemTyp As String

        If isMilestone Then
            elemTyp = "1"
        Else
            elemTyp = "0"
        End If

        If lfdNr <= 1 Then
            elemKey = elemTyp & "§" & elemName & "§"
        Else
            elemKey = elemTyp & "§" & elemName & "§" & lfdNr.ToString("000#")
        End If


        calcHryElemKey = elemKey


    End Function

    ''' <summary>
    ''' berechnet den Namen, der in selectedphases bzw. selectedMilestones reinkommt, bestehend aus: 
    ''' Breadcrumb und elemName; Breadcrumb und die einzelnen Stufen des Breadcrumbs sind getrennt durch #
    ''' </summary>
    ''' <param name="elemName"></param>
    ''' <param name="breadcrumb"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function calcHryFullname(ByVal elemName As String, ByVal breadcrumb As String) As String

        If breadcrumb = "" Then
            calcHryFullname = elemName
        Else
            calcHryFullname = breadcrumb & "#" & elemName
        End If

    End Function

    ''' <summary>
    ''' bestimmt den eindeutigen Namen des Shapes für einen Meilenstin oder eine Phase 
    ''' </summary>
    ''' <param name="hproj"></param>
    ''' <param name="elemID"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function calcPPTShapeName(ByVal hproj As clsProjekt, elemID As String) As String

        Dim tmpName As String = elemID
        If Not IsNothing(hproj) Then
            tmpName = "(" & hproj.name & "#" & hproj.variantName & ")" & elemID
        End If

        calcPPTShapeName = tmpName

    End Function

    ''' <summary>
    ''' gibt den Elem-Name und Breadcrumb als einzelne Strings zurück
    ''' </summary>
    ''' <param name="fullname"></param>
    ''' <param name="elemName"></param>
    ''' <param name="breadcrumb"></param>
    ''' <remarks></remarks>
    Public Sub splitHryFullnameTo2(ByVal fullname As String, ByRef elemName As String, ByRef breadcrumb As String)
        Dim tmpstr() As String
        Dim tmpBC As String = ""
        Dim anzahl As Integer

        tmpstr = fullname.Split(New Char() {CChar("#")}, 20)
        anzahl = tmpstr.Length
        If tmpstr.Length = 1 Then
            elemName = tmpstr(0)
        ElseIf tmpstr.Length > 1 Then
            elemName = tmpstr(anzahl - 1)
            For i As Integer = 0 To anzahl - 2
                If i = 0 Then
                    tmpBC = tmpstr(i)
                Else
                    tmpBC = tmpBC & "#" & tmpstr(i)
                End If
            Next
        Else
            elemName = "?"
        End If
        breadcrumb = tmpBC

    End Sub

    ''' <summary>
    ''' zerteilt einen String, der folgendes Format hat: breadcrumb#elemName#lfdnr in seine Bestandteile 
    ''' </summary>
    ''' <param name="fullname"></param>
    ''' <param name="elemName"></param>
    ''' <param name="breadcrumb"></param>
    ''' <param name="lfdNr"></param>
    ''' <remarks></remarks>
    Public Sub splitBreadCrumbFullnameTo3(ByVal fullname As String, ByRef elemName As String, ByRef breadcrumb As String, ByRef lfdNr As Integer)
        Dim tmpstr() As String
        Dim tmpBC As String = ""
        Dim anzahl As Integer

        tmpstr = fullname.Split(New Char() {CChar("#")}, 20)
        anzahl = tmpstr.Length
        If tmpstr.Length = 1 Then
            elemName = tmpstr(0)
            breadcrumb = ""
            lfdNr = 1
        ElseIf tmpstr.Length > 1 Then
            lfdNr = CInt(tmpstr(anzahl - 1))
            For i As Integer = 0 To anzahl - 2
                If i = 0 Then
                    tmpBC = tmpstr(i)
                Else
                    tmpBC = tmpBC & "#" & tmpstr(i)
                End If
            Next
            Call splitHryFullnameTo2(tmpBC, elemName, breadcrumb)
        Else
            elemName = "?"
            breadcrumb = ""
            lfdNr = 0
        End If

    End Sub

    ''' <summary>
    ''' gibt true zurück, wenn es sich bei der ElemID um die ID eines Meilensteins handelt 
    ''' </summary>
    ''' <param name="elemID"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function elemIDIstMeilenstein(ByVal elemID As String) As Boolean
        elemIDIstMeilenstein = elemID.StartsWith("1§")
    End Function

    ''' <summary>
    ''' extrahiert den Elem-Namen aus der ElemID 
    ''' ElemID=Typ§ElemName§lfd-Nr 
    ''' </summary>
    ''' <param name="ElemID"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function elemNameOfElemID(ByVal elemID As String) As String
        Dim tmpStr() As String

        tmpStr = elemID.Split(New Char() {CChar("§")}, 5)
        If tmpStr.Length = 3 Then
            elemNameOfElemID = tmpStr(1)
        ElseIf tmpStr.Length = 1 Then
            elemNameOfElemID = elemID
        Else
            elemNameOfElemID = "?"
        End If


    End Function


    Public Function istElemID(ByVal itemName As String) As Boolean

        Dim tmpStr() As String

        tmpStr = itemName.Split(New Char() {CChar("§")}, 5)
        If tmpStr.Length = 3 Then
            If tmpStr(0) = "1" Or tmpStr(0) = "0" Then
                istElemID = True
            Else
                istElemID = False
            End If
        Else
            istElemID = False
        End If

    End Function

    ''' <summary>
    ''' extrahiert die lfdNr aus der ElemID 
    ''' </summary>
    ''' <param name="elemID"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function lfdNrOfElemID(ByVal elemID As String) As Integer
        Dim tmpStr() As String

        tmpStr = elemID.Split(New Char() {CChar("§")}, 5)
        If tmpStr.Length = 3 Then
            Try
                If tmpStr(2) = "" Then
                    lfdNrOfElemID = 1
                Else
                    lfdNrOfElemID = CInt(tmpStr(2))
                End If
            Catch ex As Exception
                lfdNrOfElemID = 1
            End Try

        Else

            lfdNrOfElemID = 1

        End If

    End Function

    ''' <summary>
    ''' erzeugt die monatlichen Budget Werte für ein Projekt
    ''' berechnet aus dem Wert für Erloes, verteilt nach einem Schlüssel, der sich aus Marge und Kostenbedarf ergibt 
    ''' </summary>
    ''' <param name="hproj"></param>
    ''' <remarks></remarks>

    Public Sub awinCreateBudgetWerte(ByRef hproj As clsProjekt)


        Dim costValues() As Double, budgetValues() As Double
        Dim curBudget As Double, avgbudget As Double

        ' Ergänzung am 26.5.14: wenn hproj in den Längen der Bedarfe Arrays nicht konsistent ist: 
        ' anpassen 
        If Not hproj.isConsistent Then
            Call hproj.syncXWertePhases()
        End If

        costValues = hproj.getGesamtKostenBedarf
        ReDim budgetValues(costValues.Length - 1)

        curBudget = hproj.Erloes
        avgbudget = curBudget / costValues.Length

        If curBudget > 0 Then
            If costValues.Sum > 0 Then
                Dim pMarge As Double = hproj.ProjectMarge
                For i = 0 To costValues.Length - 1
                    budgetValues(i) = costValues(i) * (1 + pMarge)
                Next
            Else
                For i = 0 To costValues.Length - 1
                    budgetValues(i) = avgbudget
                Next
            End If
        End If


        hproj.budgetWerte = budgetValues


    End Sub

    ''' <summary>
    ''' aktualisiert die Budget werte , wobei die Charakteristik erhalten bleibt 
    ''' Vorbedingung ist, daß das bisherige Budget > 0 Null ist 
    ''' </summary>
    ''' <param name="hproj"></param>
    ''' <param name="newBudget">Gesamt Wert des neuen Budgets</param>
    ''' <remarks></remarks>
    Public Sub awinUpdateBudgetWerte(ByRef hproj As clsProjekt, ByVal newBudget As Double)



        Dim curValues() As Double, budgetValues() As Double
        Dim oldBudget As Double
        Dim faktor As Double

        curValues = hproj.budgetWerte
        ReDim budgetValues(curValues.Length - 1)
        oldBudget = curValues.Sum

        If oldBudget = 0 Then
            Throw New Exception("altes Budget darf beim Update nicht Null sein")
        Else
            If newBudget <= 0 Then
                ' budgetvalues ist bereits auf Null gesetzt  
            Else
                faktor = newBudget / oldBudget
                For i = 0 To curValues.Length - 1
                    budgetValues(i) = curValues(i) * faktor
                Next
            End If

        End If

        hproj.budgetWerte = budgetValues

    End Sub

    ''' <summary>
    ''' bereichnet zu einer gegebenen Y-Koordinate (Top) die dazugehörige Zeile in der Projekt-Tafel
    ''' </summary>
    ''' <param name="YCoord"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function calcYCoordToZeile(ByVal YCoord As Double) As Integer
        Dim tmpValue As Integer
        'Dim chkValue As Integer

        'chkValue = 1 + CInt((YCoord - topOfMagicBoard) / boxHeight)
        tmpValue = 1 + CInt(Truncate((YCoord - topOfMagicBoard) / boxHeight))

        'If chkValue <> tmpValue Then
        '    Call MsgBox("Fehler in calcYCoordToZeile")
        'End If

        calcYCoordToZeile = tmpValue

    End Function

    ''' <summary>
    ''' berechnet zu einer gegeb. X-Koordinate (Left) die dazugehörige Spalte in der Projekt-Tafel
    ''' </summary>
    ''' <param name="XCoord"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function calcXCoordToSpalte(ByVal XCoord As Double) As Integer
        Dim tmpValue As Integer

        tmpValue = CInt(System.Math.Truncate(XCoord / boxWidth) + 1)

        calcXCoordToSpalte = tmpValue


    End Function

    Public Function calcZeileToYCoord(ByVal zeile As Integer) As Double
        Dim tmpvalue As Double

        tmpvalue = topOfMagicBoard + (zeile - 1) * boxHeight
        calcZeileToYCoord = tmpvalue

    End Function

    ''' <summary>
    ''' berechnet, wieviel Tage vom startofCalendar bis zur angegebenen Koordinate sind 
    ''' </summary>
    ''' <param name="XCoord"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function calcXCoordToTage(ByVal XCoord As Double) As Integer
        Dim tmpValue As Integer
        tmpValue = CInt(365 * XCoord / (12 * boxWidth))

        calcXCoordToTage = tmpValue

    End Function

    ''' <summary>
    ''' berechnet wieviel Tage es vom Refdatum zu dem durch XCoord angegebenem Datum ist 
    ''' </summary>
    ''' <param name="refDate"></param>
    ''' <param name="XCoord"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function calcXCoordToTage(ByVal refDate As Date, ByVal XCoord As Double) As Integer
        Dim tmpValue As Integer
        tmpValue = CInt(365 * XCoord / (12 * boxWidth)) - CInt(DateDiff(DateInterval.Day, StartofCalendar, refDate))

        calcXCoordToTage = tmpValue

    End Function


    ''' <summary>
    ''' berechnet das Datum, das der angegebenen X-Koordinate auf der Projekt-Tafel entspricht 
    ''' </summary>
    ''' <param name="XCoord"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function calcXCoordToDate(ByVal XCoord As Double) As Date

        Dim tmpValue As Integer
        tmpValue = CInt(365 * XCoord / (12 * boxWidth))

        calcXCoordToDate = StartofCalendar.AddDays(tmpValue)


    End Function

    Public Function calcDateToXCoord(ByVal datum As Date) As Double

        Dim tmpValue As Double
        Dim anzahlTage As Integer
        anzahlTage = CInt(DateDiff(DateInterval.Day, StartofCalendar, datum))
        If anzahlTage < 0 Then
            Throw New ArgumentException("Datum kann nicht vor Start des Kalenders liegen")
        End If

        tmpValue = anzahlTage * 12 * boxWidth / 365

        calcDateToXCoord = tmpValue


    End Function

    ''' <summary>
    ''' bestimmt den Prozentsatz der Überdeckung der beiden durch Start- und End-Datum angegebenen Phasen 
    ''' </summary>
    ''' <param name="startDate1">StartDatum Phase 1</param>
    ''' <param name="endDate1">Ende Datum Phase 1</param>
    ''' <param name="startDate2">Startdatum Phase 2</param>
    ''' <param name="enddate2">Ende Datum Phase 2</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function calcPhaseUeberdeckung(ByVal startDate1 As Date, endDate1 As Date, _
                                              ByVal startDate2 As Date, ByVal enddate2 As Date) As Double

        Dim duration1 As Long = DateDiff(DateInterval.Day, startDate1, endDate1) + 1
        Dim duration2 As Long = DateDiff(DateInterval.Day, startDate2, enddate2) + 1
        Dim ueberdeckungsStart As Date, ueberdeckungsEnde As Date
        Dim ueberdeckungsduration As Long
        Dim ergebnis As Double


        If DateDiff(DateInterval.Day, endDate1, startDate2) > 0 Or _
            DateDiff(DateInterval.Day, enddate2, startDate1) > 0 Then
            ' es gibt gar keine Überdeckung ...
            ergebnis = 0.0
        Else

            If DateDiff(DateInterval.Day, startDate1, startDate2) >= 0 Then
                ueberdeckungsStart = startDate2
            Else
                ueberdeckungsStart = startDate1
            End If

            If DateDiff(DateInterval.Day, endDate1, enddate2) >= 0 Then
                ueberdeckungsEnde = endDate1
            Else
                ueberdeckungsEnde = enddate2
            End If

            ueberdeckungsduration = DateDiff(DateInterval.Day, ueberdeckungsStart, ueberdeckungsEnde) + 1
            'ergebnis = System.Math.Max(ueberdeckungsduration / duration1, ueberdeckungsduration / duration2)
            ergebnis = System.Math.Min(ueberdeckungsduration / duration1, ueberdeckungsduration / duration2)

        End If

        calcPhaseUeberdeckung = ergebnis


    End Function
    '
    ' Funktion prüft, ob der angegebene Name bereits Element der Projektliste ist
    '
    Function inProjektliste(strName As String) As Boolean

        Dim found As Boolean = False
        Dim foundinDatabase As Boolean = False
        Dim key As String = calcProjektKey(strName, "")
        'Dim request As New Request(awinSettings.databaseName)

        If Len(strName) < 2 Then
            ' ProjektName soll mehr als 1 Zeichen haben
            found = True
        ElseIf AlleProjekte.Containskey(key) Then
            found = True
            'ElseIf request.pingMongoDb() Then

            '    found = request.projectNameAlreadyExists(strName, "", Date.Now)
            'Else
            '    Call MsgBox("Datenbank-Verbindung ist unterbrochen!")
            '    found = False
        End If

        inProjektliste = found

    End Function

    ''' <summary>
    ''' übersetzt den im Import File angegebenen String auf die Standard Darstellungsklassen 
    ''' wenn nicht gemappt werden kann, wird "" zurückgegeben 
    ''' </summary>
    ''' <param name="completeText">die Bezeichnung für die Darstellungsklasse aus RPLAN</param>
    ''' <returns>den Standard Namen</returns>
    ''' <remarks></remarks>
    Public Function mapToAppearance(ByVal completeText As String, isMilestone As Boolean) As String
        Dim ergebnis As String = ""
        Dim found As Boolean = False
        Dim index As Integer = 0
        Dim anzElements As Integer = appearanceDefinitions.Count


        Do While Not found And index <= anzElements - 1

            If completeText.Contains(appearanceDefinitions.ElementAt(index).Key.Trim) And _
                    isMilestone = appearanceDefinitions.ElementAt(index).Value.isMilestone Then
                found = True
                ergebnis = appearanceDefinitions.ElementAt(index).Key
            Else
                index = index + 1
            End If

        Loop

        mapToAppearance = ergebnis

    End Function

    '' '' ''' <summary>
    '' '' ''' speichert den letzten Filter und setzt die temporären Collections wieder zurück 
    '' '' ''' </summary>
    '' '' ''' <remarks></remarks>
    '' ''Public Sub storeFilter(ByVal fName As String, ByVal menuOption As Integer, _
    '' ''                                          ByVal fBU As Collection, ByVal fTyp As Collection, _
    '' ''                                          ByVal fPhase As Collection, ByVal fMilestone As Collection, _
    '' ''                                          ByVal fRole As Collection, ByVal fCost As Collection, _
    '' ''                                          ByVal calledFromHry As Boolean)

    '' ''    Dim lastFilter As clsFilter


    '' ''    If calledFromHry Then
    '' ''        Dim nameLastFilter As clsFilter = filterDefinitions.retrieveFilter("Last")

    '' ''        If Not IsNothing(nameLastFilter) Then
    '' ''            With nameLastFilter
    '' ''                lastFilter = New clsFilter(fName, .BUs, .Typs, fPhase, fMilestone, .Roles, .Costs)
    '' ''            End With
    '' ''        Else
    '' ''            lastFilter = New clsFilter(fName, fBU, fTyp, _
    '' ''                              fPhase, fMilestone, _
    '' ''                             fRole, fCost)
    '' ''        End If


    '' ''    Else
    '' ''        lastFilter = New clsFilter(fName, fBU, fTyp, _
    '' ''                              fPhase, fMilestone, _
    '' ''                             fRole, fCost)
    '' ''    End If

    '' ''    If menuOption = PTmenue.filterdefinieren Then

    '' ''        filterDefinitions.storeFilter(fName, lastFilter)
    '' ''        Dim request As New Request(awinSettings.databaseName, dbUsername, dbPasswort)


    '' ''    Else
    '' ''        selFilterDefinitions.storeFilter(fName, lastFilter)
    '' ''    End If


    '' ''End Sub

    

    ''' <summary>
    ''' besetzt die Selection Collections mit den Werten des Filters mit Namen fName
    ''' </summary>
    ''' <param name="fName"></param>
    ''' <param name="selectedBUs"></param>
    ''' <param name="selectedTyps"></param>
    ''' <param name="selectedPhases"></param>
    ''' <param name="selectedMilestones"></param>
    ''' <param name="selectedRoles"></param>
    ''' <param name="selectedCosts"></param>
    ''' <remarks></remarks>
    Public Sub retrieveSelections(ByVal fName As String, ByVal menuOption As Integer, _
                                       ByRef selectedBUs As Collection, ByRef selectedTyps As Collection, _
                                       ByRef selectedPhases As Collection, ByRef selectedMilestones As Collection, _
                                       ByRef selectedRoles As Collection, ByRef selectedCosts As Collection)

        Dim lastFilter As clsFilter

        If menuOption = PTmenue.filterdefinieren Or _
            menuOption = PTmenue.sessionFilterDefinieren Or _
            menuOption = PTmenue.filterAuswahl Then
            lastFilter = filterDefinitions.retrieveFilter(fName)
        Else
            lastFilter = selFilterDefinitions.retrieveFilter(fName)

            ' ur: 30.07.2015: wenn kein selFilterDefinitions existiert, so soll auch keine Voreinstellung angezeigt werden
            ' ''If IsNothing(lastFilter) Then
            ' ''    lastFilter = filterDefinitions.retrieveFilter(fName)
            ' ''End If
        End If


        If Not IsNothing(lastFilter) Then

            'selectedBUs = lastFilter.BUs
            selectedBUs = copyCollection(lastFilter.BUs)
            selectedTyps = copyCollection(lastFilter.Typs)
            selectedPhases = copyCollection(lastFilter.Phases)
            selectedMilestones = copyCollection(lastFilter.Milestones)
            selectedRoles = copyCollection(lastFilter.Roles)
            selectedCosts = copyCollection(lastFilter.Costs)

        Else
            selectedBUs = New Collection
            selectedTyps = New Collection
            selectedPhases = New Collection
            selectedMilestones = New Collection
            selectedRoles = New Collection
            selectedCosts = New Collection
        End If

    End Sub

    ''' <summary>
    ''' kennzeichnet ein Powerpoint Slide als ein Slide, das Smart Elements enthält 
    ''' fügt die Kennzeichnung "SMART" mit type an, StartofCalendar, CreationDate und Datenbank Infos 
    ''' </summary>
    ''' <param name="pptSlide"></param>
    ''' <remarks></remarks>
    Public Sub addSmartPPTSlideInfo(ByRef pptSlide As PowerPoint.Slide, _
                                    ByVal type As String, _
                                    ByVal calendarLeft As Date, _
                                    ByVal calendarRight As Date)

        If Not IsNothing(pptSlide) Then
            With pptSlide

                If Not IsNothing(type) Then
                    .Tags.Add("SMART", type)
                    .Tags.Add("SOC", StartofCalendar.ToShortDateString)
                    .Tags.Add("CRD", Date.Now.ToString)
                    .Tags.Add("CALL", calendarLeft.ToShortDateString)
                    .Tags.Add("CALR", calendarRight.ToShortDateString)

                    If Not noDB Then
                        If awinSettings.databaseURL.Length > 0 Then
                            .Tags.Add("DBURL", awinSettings.databaseURL)
                        End If
                        If awinSettings.databaseName.Length > 0 Then
                            .Tags.Add("DBNAME", awinSettings.databaseName)
                        End If

                    End If
                End If


            End With
        End If


    End Sub


    ''' <summary>
    ''' fügt an ein Powerpoint Shape Informationen über Tags an, die vom PPT Add-In SmartPPT ausgelesen werden können
    ''' </summary>
    ''' <param name="pptShape"></param>
    ''' <param name="fullBreadCrumb"></param>
    ''' <param name="classifiedName"></param>
    ''' <param name="shortName"></param>
    ''' <param name="originalName"></param>
    ''' <param name="startDate"></param>
    ''' <param name="endDate"></param>
    ''' <param name="ampelColor"></param>
    ''' <param name="ampelErlaeuterung"></param>
    ''' <remarks></remarks>
    Public Sub addSmartPPTShapeInfo(ByRef pptShape As PowerPoint.Shape, _
                                          ByVal fullBreadCrumb As String, ByVal classifiedName As String, ByVal shortName As String, ByVal originalName As String, _
                                          ByVal startDate As Date, ByVal endDate As Date, _
                                          ByVal ampelColor As Integer, ByVal ampelErlaeuterung As String, _
                                          ByVal lieferumfaenge As String)

        Dim nullDate As Date = Nothing

        If Not IsNothing(pptShape) Then
            With pptShape


                If Not IsNothing(fullBreadCrumb) Then
                    .Tags.Add("BC", fullBreadCrumb)
                End If

                If Not IsNothing(classifiedName) Then
                    .Tags.Add("CN", classifiedName)
                End If

                If Not IsNothing(shortName) Then
                    If shortName <> classifiedName And shortName <> "" Then
                        .Tags.Add("SN", shortName)
                    End If
                End If

                If Not IsNothing(originalName) Then
                    If originalName <> classifiedName And originalName <> "" Then
                        .Tags.Add("ON", originalName)
                    End If
                End If

                If Not IsNothing(startDate) Then
                    If Not startDate = nullDate Then
                        .Tags.Add("SD", startDate.ToShortDateString)
                    End If
                End If

                If Not IsNothing(endDate) Then
                    If Not endDate = nullDate Then
                        .Tags.Add("ED", endDate.ToShortDateString)
                    End If

                End If

                If Not IsNothing(ampelColor) Then
                    If ampelColor >= 0 And ampelColor <= 3 Then
                        .Tags.Add("AC", ampelColor.ToString)
                    Else
                        .Tags.Add("AC", "0")
                    End If

                    If Not IsNothing(ampelErlaeuterung) Then
                        If ampelErlaeuterung.Length > 0 Then
                            .Tags.Add("AE", ampelErlaeuterung)
                        End If

                    End If

                End If

                If Not IsNothing(lieferumfaenge) Then
                    If lieferumfaenge.Length > 0 Then
                        .Tags.Add("LU", lieferumfaenge)
                    End If

                End If

            End With
        End If



    End Sub



    Public Sub PPTstarten()
        Try
            ' prüft, ob bereits Powerpoint geöffnet ist 
            pptApp = CType(GetObject(, "PowerPoint.Application"), pptNS.Application)
        Catch ex As Exception
            Try
                pptApp = CType(CreateObject("PowerPoint.Application"), pptNS.Application)
            Catch ex1 As Exception
                Throw New ArgumentException("Powerpoint konnte nicht gestartet werden ...", ex1.Message)
                'Exit Sub
            End Try

        End Try
    End Sub
End Module
