Imports ProjectBoardDefinitions
Imports System.Globalization
Imports System.Collections.Generic
Imports System.Math
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Core
Imports pptNS = Microsoft.Office.Interop.PowerPoint
Imports System.Xml.Serialization



Public Module Module1


    ' in Modul 1 sollten jetzt alle Konstanten und Einstellungen in einer Klasse zusammengefasst werden
    ' awinSettings: für StartOfCalendar, linker Rand, rechter Rand, ...
    ' Laufzeit Parameter;

    ' das Objekt, das später die Instanz-Variable Request aufnimmt 
    Public databaseAcc As Object = Nothing

    Public iDkey As String = ""

    'login - Informationen
    Public dbUsername As String = ""
    Public dbPasswort As String = ""
    Public loginErfolgreich As Boolean = False
    Public noDB As Boolean = True

    'Name des VisboClient
    Public visboClient As String = "VISBO Projectboard / "

    'Cache - Infos
    Public cacheUpdateDelay As Long

    ' tk 4.12.18 
    Public dbUserID As String = ""
    ' hier sind für den eingeloggten Nutzer  aktuell gewählte  alle 
    Public customUserRoles As New clsCustomUserRoles

    ' wird verwendet um Informationen verschlüsselt zu schreiben 
    Public visboCryptoKey As String = "Berge2007QuebecKanada&2010SeilmitZeltThomasUtePhilippDenise060162130790141090050715&@Tecoplan@IPEQ@Visbo"

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
    Public myLogfile As String = ""

    ' gibt an, in welchem Modus sich aktuell die Projekt-Tafe befindet 
    Public currentProjektTafelModus As Integer

    'Definition der Klasse für die ReportMessages ( müssen in awinSettypen gelesen werden aus xml-File)
    Public repMessages As clsReportMessages

    'Definition der Klasse für die ReportMessages ( müssen in awinSettypen gelesen werden aus xml-File)
    Public menuMessages As clsReportMessages

    
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


    ' hier werden die gültigen Organisationen mitgemerkt ... 
    Public validOrganisations As New clsOrganisations


    Public Projektvorlagen As New clsProjektvorlagen
    Public ModulVorlagen As New clsProjektvorlagen
    Public ShowProjekte As New clsProjekte
    ' noShowProjekte am 21.3 rausgenommen 
    ''Public noShowProjekte As New clsProjekte
    Public selectedProjekte As New clsProjekte
    'Public AlleProjekte As New SortedList(Of String, clsProjekt)
    Public AlleProjekte As New clsProjekteAlle

    ' ist das Pendant zu AlleProjekte, nimmt nur Summary Projekte auf 
    Public AlleProjektSummaries As New clsProjekteAlle
    ' ist das Pendant zu ShowProjekte, gibt an welche Summary Projekte geladen sind  
    Public ShowProjekteSummaries As New clsProjekte

    ' der DBCache der von allen Projekten angelegt wird, die im Mass-Edit bearbeitet werden 
    ' evtl wird das später mal erweitert auf alleProjekte, die geladen sind und in der DB existieren
    ' damit liesse sich die Zeit deutlich reduzieren , wenn es um den Vergleich aktueller Stand / DB Stand geht 
    Public sessionCacheProjekte As New clsProjekteAlle

    ' die globale Variable für die Write Protections
    Public writeProtections As New clsWriteProtections

    Public ImportProjekte As New clsProjekteAlle
    Public projectConstellations As New clsConstellations
    ' die currentSessionConstellation ist das Abbild der aktuellen Session 
    Public currentSessionConstellation As New clsConstellation
    ' die beforeFilterConstellation ist das Abbild der aktuellen Session vor einer Filteraktion
    Public beforeFilterConstellation As New clsConstellation
    Public currentConstellationName As String = "" ' hier wird mitgeführt, was die aktuelle Projekt-Konstellation ist 
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
    Public RoleHierarchy As New clsroleHrchy
    Public PhaseDefinitions As New clsPhasen
    Public MilestoneDefinitions As New clsMeilensteine


    Public CostDefinitions As New clsKostenarten
    ' Welche Business-Units gibt es ? 
    Public businessUnitDefinitions As New SortedList(Of Integer, clsBusinessUnit)

    ' welche CustomFields gibt es ? 
    Public customFieldDefinitions As New clsCustomFieldDefinitions

    ' was ist meine CustomUser Role, die ich für die aktuelle Slide brauche ? 
    Public myCustomUserRole As New clsCustomUserRole


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
    ' tk 9.4.19 das soll nicht mehr gezeigt werden ... 
    'Public formMilestone As New frmMilestoneInformation
    'Public formStatus As New frmStatusInformation
    'Public formPhase As New frmPhaseInformation
    Public formProjectInfo1 As frmProjectInfo1 = Nothing



    ' diese beiden folgenden Variablen steuern im Sheet "Ressourcen", welcher Bereich in den Diagrammen angezeigt werden soll
    Public showRangeLeft As Integer
    Public showRangeRight As Integer

    ' diese beiden Variablen nehmen die Farben auf für Showtimezone bzw. Noshowtimezone
    Public showtimezone_color As Object, noshowtimezone_color As Object, calendarFontColor As Object


    ' maxScreenHeight, maxScreenWidth gibt die maximale Höhe/Breite des Bildschirms in Punkten an 
    Public maxScreenHeight As Double, maxScreenWidth As Double
    Public boxWidth As Double = 19.3, boxHeight As Double, topOfMagicBoard As Double
    Public screen_correct As Double = 0.26
    Public chartWidth As Double = 140 ' wird aber noch in Abhängigkeit von maxscreenwidth gesetzt 
    Public chartHeight As Double = 120 ' wird aber noch in abhängigkeit von maxscreenheight gesetzt

    ' dieser Array dient zur Aufnahme der Spaltenbreiten, Schriftgrösse für MassEditRC (0), massEditTE (1), massEditAT (2)
    Public massColFontValues(2, 100) As Double

    ' diese Konstante legt den Namen für das Root Element , 1. Phase eines Projektes fest 
    ' das muss mit der calcHryElemKey(".", False) übereinstimmen 
    Public Const rootPhaseName As String = "0§.§"

    ' diese Konstante bestimmt, welchen Varianten Namen Portfolios bzw. Programme bekommen 
    'Public Const portfolioVName As String = ""
    ' diese Konstante bestimmt, wie die Variante heissen soll, die die Ist-Daten - zumindest temporär - aufnimmt 
    'Public Const istDatenVName As String = "ActualData"
    ' diese Konstante wird benutzt, wenn keine Variante angegeben wurde, d.h. meistens das alle Variante relevant sind.
    Public Const noVariantName As String = "-9999999"

    ' diese Konstante wird verwendet, um den VisboImportTyp zu erkennen
    Public Const visboImportKennung = "VisboImportTyp"

    Public visboFarbeBlau As Integer = RGB(69, 140, 203)
    Public visboFarbeOrange As Integer = RGB(247, 148, 30)
    Public visboFarbeNone As Integer = RGB(127, 127, 127)
    Public visboFarbeGreen As Integer = RGB(0, 176, 80)
    Public visboFarbeYellow As Integer = RGB(255, 197, 13)
    Public visboFarbeRed As Integer = RGB(255, 0, 0)

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

    Public Enum ptVariantFixNames
        pfv = 0 ' für den Portfolio Manager, für die Vorgaben reserviert
        acd = 1 ' ggf später für ActualData 
    End Enum

    Public settingTypes() = {"customfields", "customroles", "organisation", "customization", "phasemilestonedefs", "reportprofiles", "ppttemplates", "generalexcelcharts"}

    Public Enum ptSettingTypes
        customfields = 0
        customroles = 1
        organisation = 2
        customization = 3
        phasemilestonedefs = 4
        reportprofiles = 5
        ppttemplates = 6
        generalexcelcharts = 7
    End Enum

    Public customUserRoleBezeichner() As String = {"Organiations-Admin", "Portfolio", "Ressourcen", "Projektleiter", "All"}

    ' hier wird geregelt, wer denn welche Menu-Punkte sehen darf
    Public customUserRoleAllowance(,) As String = Nothing

    ''' <summary>
    ''' Werte-Bereich: {0=Admin, 1=PortfolioMgr; 2=RessourcenManager; 3=Projektleiter
    ''' </summary>
    Public Enum ptCustomUserRoles
        OrgaAdmin = 0
        PortfolioManager = 1
        RessourceManager = 2
        ProjektLeitung = 3
        Alles = 4
        InternalViewer = 5
        ExternalViewer = 6
        TeamManager = 7
    End Enum

    ''' <summary>
    ''' definiert, welche Import-Methode angewendet werden soll ; angelegt 30.11.18 by tk
    ''' </summary>
    Public Enum ptVisboImportTypen
        visboSimple = 0
        visboProjectbrief = 1
        visboMassCreation = 2
        visboRXF = 3
        visboExcelBMW = 4
        allianzMassImport1 = 5
        allianzMassImport2 = 6
        allianzTeamRessZuordnung = 7
        allianzIstDaten = 8
        visboMassRessourcenEdit = 9
        visboMPP = 10
    End Enum

    Public Enum ptImportSettings
        attributeNames1 = 0
        attributeNamesCol1 = 1
        roleCostNames1 = 2
        roleCostNamesCol1 = 3
        customFieldNames1 = 4
        customFieldsNamesCol1 = 5
        ' Werte für Import Typ 2
        attributeNames2 = 6
        attributeNamesCol2 = 7
        roleCostNames2 = 8
        roleCostNamesCol2 = 9
        customFieldNames2 = 10
        customFieldsNamesCol2 = 11
        ' Werte für Import Typ 3
        attributeNames3 = 12
        attributeNamesCol3 = 13
        roleCostNames3 = 14
        roleCostNamesCol3 = 15
        customFieldNames3 = 16
        customFieldsNamesCol3 = 17

    End Enum

    Public Enum ptReportBigTypes
        charts = 0
        tables = 1
        components = 2
        planelements = 4
    End Enum

    'Public Enum ptReportTables
    '    prMilestones = 0
    '    pfMilestones = 1
    'End Enum

    Public Enum ptReportComponents
        prAmpel = 0
        prStand = 1
        prName = 2
        prCustomField = 3
        prAmpelText = 4
        prDescription = 5
        prBusinessUnit = 6
        prLaufzeit = 7
        prVerantwortlich = 8
        prRisks = 9
        pfStand = 11
        prSymRisks = 12
        prSymTrafficLight = 13
        prSymDescription = 14
        prCard = 15
        prCardinvisible = 16
        prSymFinance = 17
        prSymSchedules = 18
        prSymTeam = 19
        prSymProject = 20
        pfName = 21
    End Enum

    ' wenn diese Enum erweitert wird, inbedingt im clsProjekt .projecttype Property den Wertebereich anpassen ...
    ' in mongoDBaccess wird statisch auf "0" für projekt abgefragt ..
    Public Enum ptPRPFType
        project = 0
        portfolio = 1
        projectTemplate = 2
        all = 3
    End Enum

    ''' <summary>
    ''' kann verwendet werden, um die Typen zu kennzeichnen
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum ptElementTypen
        phases = 0
        roles = 1
        costs = 2
        portfolio = 3
        ergebnis = 4
        milestones = 5
        mta = 6
        rolesAndCost = 7
    End Enum


    ' gibt an, nach welchem Sortierkriterium die _sortList aufgebaut wurde 
    ' 0: alphabetisch nach Name
    ' 1: custom tfzeile 
    ' 2: custom Liste
    ' 3: BU, ProjektStart, Name
    ' 4: Formel: strategic Fit* 100 - risk*90 + 100*Marge + korrFaktor
    Public Enum ptSortCriteria
        alphabet = 0
        customTF = 1
        customListe = 2
        strategyProfitLossRisk = 3
        customFields12 = 4
        buStartName = 5
        formel = 6
        strategyRiskProfitLoss = 7
    End Enum

    Public Enum ptTables
        none = 0
        repCharts = 1
        MPT = 3
        cstSettings = 4
        meRC = 5
        meTE = 6
        cstPmappings = 8
        meAT = 9
        cstMmappings = 10
        meCharts = 11
        mptPfCharts = 12
        mptPrCharts = 13
        cstMissingDefs = 15
    End Enum

    Public Enum ptWriteProtectionType
        project = 0
        scenario = 1
    End Enum

    Public Enum ptSzenarioConsider
        all = 0
        show = 1
        noshow = 2
    End Enum

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
        massEditTermine = 2
        massEditAttribute = 3
    End Enum

    ' die NAmen für die RPLAN Spaltenüberschriften in Rplan Excel Exports 
    Public Enum ptPlanNamen
        Name = 0
        Anfang = 1
        Ende = 2
        Beschreibung = 3
        Vorgangsklasse = 4
        BusinessUnit = 5
        Protocol = 6
        Dauer = 7
        Abkuerzung = 8
        Verantwortlich = 9
        percentDone = 10
        TrafficLight = 11
        TLExplanation = 12
        DocUrl = 13
        Deliv = 14
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


    ''' <summary>
    ''' eine Auflistung der Charttypen (Balken, Curve, Pie, etc
    ''' </summary>
    Public Enum PTChartTypen
        Balken = 0
        ZweiBalken = 1
        CurveCumul = 2
        Pie = 3
        Bubble = 4
        Waterfall = 5
    End Enum

    Public Enum PTVergleichsArt
        beauftragung = 0
        planungsstand = 1
    End Enum

    Public Enum PTVergleichsTyp
        erster = 0
        letzter = 1
        standVom = 2
    End Enum

    Public Enum PTEinheiten
        personentage = 0
        euro = 1
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
        PhaseCategories = 21
        MilestoneCategories = 22
    End Enum

    ' Enumeration Projekt Diagramm Kennungen 
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
        MilestoneTrendanalysis = 17
        SollIstPersonalkosten = 18
        SollIstSonstKosten = 19
        SollIstGesamtkosten = 20
        SollIstPersonalkostenC = 21
        SollIstSonstKostenC = 22
        SollIstGesamtkostenC = 23
        SollIstRolleC = 24
        SollIstKostenartC = 25
        PersonalBalken2 = 26
        KostenBalken2 = 27
        SollIstPersonalkosten2 = 28
        SollIstSonstKosten2 = 29
        SollIstGesamtkosten2 = 30
        SollIstPersonalkostenC2 = 31
        SollIstSonstKostenC2 = 32
        SollIstGesamtkostenC2 = 33
        SollIstRolleC2 = 34
        SollIstKostenartC2 = 35
        ProjektbedarfsChart = 36
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
        onlymyteammembers = 3
    End Enum

    Public Enum PThis
        current = 0
        vorlage = 1
        beauftragung = 2
        letzterStand = 3
        ersterStand = 4
        letzteBeauftragung = 5
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
        setWriteProtection = 12
        unsetWriteProtection = 13
        reportMultiprojektTafel = 14
    End Enum

    Public Enum PTProtectedMode
        noProtection = 0
        byOther = 1
        byMe = 2
    End Enum

    Public Enum PTLeadOrDependency
        none = 0
        lead = 1
        dependent = 2
    End Enum

    Public Enum PTTreeNodeTyp
        project = 0
        pVariant = 1
        timestamp = 2
    End Enum
    Public Enum PTlicense
        swimlanes = 0

    End Enum

    Public Enum PTpptAnnotationType
        text = 0
        datum = 1
        calloutAmpel = 2
        calloutLU = 3
        calloutRC = 4
        calloutMV = 5
    End Enum

    Public Enum PTpptTableTypes
        prZiele = 0
        prBudgetCostAPVCV = 1
        prMilestoneAPVCV = 2
    End Enum
    Public Enum PTpptTableCellType
        name = 0
        lfdNr = 1
        ampelColor = 2
        ampelText = 3
        lieferumfang = 4
        datum = 5
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

    Public Enum PTItemType
        vorlage = 0
        projekt = 1
        nameList = 2
        categoryList = 3
        portfolio = 4
    End Enum

    ''' <summary>
    ''' Aufzaehlung der Windows, wird in projectboardwindows(x) verwendet 
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum PTwindows
        mpt = 0
        mptpr = 1
        mptpf = 2
        meChart = 3
        massEdit = 4
    End Enum

    ''' <summary>
    ''' Aufzählung der Views, wird in projectboardViews(x) verwendet 
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum PTview
        mpt = 0
        mptpr = 1
        mptprpf = 2
        meOnly = 3
        meChart = 4
    End Enum

    ' wird in awinSetTypen dimensioniert und gesetzt 
    Public portfolioDiagrammtitel() As String

    ' nimmt die Namen der im Zuge der Optimierung automatisch generierten Szenarios auf
    Public autoSzenarioNamen(3) As String


    ' dieser array nimmt die Koordinaten der Formulare auf 
    ' die Koordinaten werden in der Reihenfolge gespeichert: top, left, width, height 
    Public frmCoord(23, 3) As Double

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
        projInfoPL = 22
        rolecostME = 23
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
        loadPV = 6
        deleteV = 7
        chgInSession = 8
        delAllExceptFromDB = 9
        setWriteProtection = 10
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
        batchlists = 3
        modulScen = 4
        massenEdit = 5
        addElements = 6
        rplanrxf = 7
        scenariodefs = 8
        Orga = 9
        Kapas = 10
        customUserRoles = 11
        actualData = 12
        offlineData = 13
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

    Public Enum PTProjektStati
        geplant = 0
        beauftragt = 1
        ChangeRequest = 2
        abgebrochen = 3
        abgeschlossen = 4
        ' die beiden folgenden nicht mehr verwenden ! 
        geplanteVorgabe = 5
        beauftragteVorgabe = 6
    End Enum



    ' wird in Customization File gesetzt - dies hier ist nur die Default Einstellung 
    ' soll so früh gesetzt sein, damit 
    Public StartofCalendar As Date = #1/1/2015#

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
    ' Vorgabe, geplant
    ' Vorgabe beauftragt
    Public ProjektStatus() As String = {"geplant", "beauftragt", "beauftragt, Änderung noch nicht freigegeben", "beendet", "abgeschlossen", "geplanteVorgabe", "beauftragteVorgabe"}


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
    Public ReportLang() As CultureInfo = {New CultureInfo("de-DE"),
                                         New CultureInfo("en-US"),
                                         New CultureInfo("fr-FR"),
                                         New CultureInfo("es-ES")}
    ' aktuell verwendete Sprache für reports
    '
    Public repCult As CultureInfo

    ' aktuell verwendete Sprache für Menu Strukturen 
    Public menuCult As CultureInfo
    '
    '
    ' Diagramm-Typ kann sein:
    ' Phase
    ' Rolle
    ' Kostenart
    ' Summe
    ' portfolio

    ' Variable nimmt die Namen der Diagramm-Typen auf 
    Public DiagrammTypen(8) As String

    ' Variable nimmt die Namen der Windows auf  
    Public windowNames(4) As String

    ' nimmt alle Excel.Window Definitionen auf 
    Public projectboardWindows(4) As Excel.Window

    ' Variable nimmt die View Namen auf ; eine View ist eine Zusammenstellung von Windows
    Public projectboardViews(4) As Excel.CustomView

    ' Variable nimmt die Namen der Ergebnis Charts auf  
    Public ergebnisChartName(3) As String

    ' tk 25.8.17 Nonsense, wird nicht gebraucht 
    ' diese Variabe nimmt die Farbe der Kapa-Linie an
    'Public rollenKapaFarbe As Object

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
    Public awinsetTypen_Performed As Boolean = False


    Private Declare Function OpenClipboard& Lib "user32" (ByVal hwnd As Long)
    Private Declare Function EmptyClipboard Lib "user32" () As Long
    Private Declare Function CloseClipboard& Lib "user32" ()


    Public Sub ClearClipboard()
        OpenClipboard(0&)
        EmptyClipboard()
        CloseClipboard()
    End Sub




    ''' <summary>
    ''' setzt EnableEvents, ScreenUpdating auf true
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub projektTafelInit()

        With appInstance
            .EnableEvents = True
            If .ScreenUpdating = False Then
                .ScreenUpdating = True
            End If
        End With


    End Sub


    ''' <summary>
    ''' aktiviert, wenn visible, das Multiprojekt Window ...
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub activateProjectBoard()
        Try

            If Not IsNothing(projectboardWindows(PTwindows.mpt)) Then

                If projectboardWindows(PTwindows.mpt).Visible = True Then
                    projectboardWindows(PTwindows.mpt).Activate()
                End If

            End If

        Catch ex As Exception

        End Try

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

        Dim found As Boolean
        Dim chtobjName As String
        Dim tmpStr(20) As String


        found = False


        chtobjName = chtobj.Name

        Try

            tmpStr = chtobjName.Split(New Char() {CChar("#")}, 20)
            If tmpStr(0) = "pf" And tmpStr.Length >= 2 Then

                If CInt(tmpStr(1)) = PTpfdk.UeberAuslastung Then
                    found = True
                    index = PTpfdk.UeberAuslastung

                ElseIf CInt(tmpStr(1)) = PTpfdk.Unterauslastung Then
                    found = True
                    index = PTpfdk.Unterauslastung

                ElseIf CInt(tmpStr(1)) = PTpfdk.Auslastung Then
                    found = True
                    index = PTpfdk.Auslastung

                ElseIf CInt(tmpStr(1)) = PTpfdk.ErgebnisWasserfall Then
                    found = True
                    index = PTpfdk.ErgebnisWasserfall

                Else
                    found = False
                End If

            End If


        Catch ex As Exception
        End Try


        istErgebnisDiagramm = found

    End Function



    ''' <summary>
    ''' liefert eine Liste an Namen der Projekte zurück, die als "selektiert" gelten 
    ''' Projekte können selektiert werden durch explizites Selektieren des Project-Shapes, alle Projekte, die zu einem selektierten Chart beitragen und/oder 
    ''' alle Projekte, die markiert sind ... 
    ''' 
    ''' </summary>
    ''' <param name="takeAllIFNothingWasSelected"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getProjectSelectionList(ByVal takeAllIFNothingWasSelected As Boolean) As Collection
        Dim tmpCollection As Collection = New Collection
        Dim msg As String = ""

        Dim chtobjName As String = ""


        ' Exit, wenn nicht im PRojekt-Tafel-Modus 
        If visboZustaende.projectBoardMode <> ptModus.graficboard Then
            ' leere Menge zurückgegeben 
        Else

            Try
                If ShowProjekte.Count > 0 Then

                    ' alle selektierten Projekte aufnehmen ... 
                    If selectedProjekte.Count > 0 Then
                        For Each kvp As KeyValuePair(Of String, clsProjekt) In selectedProjekte.Liste

                            If kvp.Value.projectType = ptPRPFType.project Or
                                myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager Then

                                If Not tmpCollection.Contains(kvp.Key) Then
                                    ' nur aufnehmen, wenn das Projekt überhaupt im Timeframe liegt ... 
                                    If kvp.Value.isWithinTimeFrame(showRangeLeft, showRangeRight) Then
                                        tmpCollection.Add(kvp.Key, kvp.Key)
                                    End If
                                End If

                            End If


                        Next
                    Else
                        ' soll nur ausgewertet werden, wenn keine einzelnen Projekte selektiert waren 
                        ' jetzt soll geprüft werden, ob irgendwelche Projekte markiert sind, die sollen auch alle übernommen werden 
                        For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

                            If kvp.Value.marker = True And (kvp.Value.projectType = ptPRPFType.project Or
                                                            myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager) Then

                                If Not tmpCollection.Contains(kvp.Key) Then
                                    ' nur aufnehmen, wenn das Projekt überhaupt im Timeframe liegt ... 
                                    If kvp.Value.isWithinTimeFrame(showRangeLeft, showRangeRight) Then
                                        tmpCollection.Add(kvp.Key, kvp.Key)
                                    End If

                                End If
                            End If

                        Next

                    End If

                    If tmpCollection.Count = 0 And takeAllIFNothingWasSelected Then

                        ' Portfolio Manager darf Summary Projekte bearbeiten
                        If myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager Then
                            ' jetzt alle Projekte aufnehmen, die in der TimeFrame liegen 
                            tmpCollection = ShowProjekte.withinTimeFrame(PTpsel.alle, showRangeLeft, showRangeRight)
                        Else
                            ' es dürfen keine Summary Projekte enthalten sein ...
                            If ShowProjekte.containsAnySummaryProject Then
                                If awinSettings.englishLanguage Then
                                    Call MsgBox("no summary projects allowed in this context ... please select projects only. ")
                                Else
                                    Call MsgBox("Summary Projekte nicht zugelassen ... bitte nur einfache Projekte auswählen.")
                                End If
                            Else
                                ' jetzt alle Projekte aufnehmen, die in der TimeFrame liegen 
                                tmpCollection = ShowProjekte.withinTimeFrame(PTpsel.alle, showRangeLeft, showRangeRight)
                            End If
                        End If

                    End If


                Else
                    If awinSettings.englishLanguage Then
                        msg = "no active projects - please load or activate projects first"
                    Else
                        msg = "keine aktiven Projekte - bitte zuerst Projekte laden bzw. aktivieren"
                    End If
                    Call MsgBox(msg)
                End If
            Catch ex As Exception

                Call MsgBox(ex.Message)
            End Try

        End If

        'testActiveWindow = CStr(appInstance.ActiveWindow.Caption)

        getProjectSelectionList = tmpCollection

    End Function

    ''' <summary>
    ''' markiert alle Projekte, die zu dem Chart beitragen 
    ''' wird aus einem Chart Event heraus aus aufgerufen, d.h das Chart existiert 
    ''' </summary>
    ''' <param name="chtObj"></param>
    ''' <remarks></remarks>
    Public Sub markProjectsOFChart(ByVal chtObj As Excel.ChartObject)

        ' jetzt bestimmen, welches Projekt zu diesem Chart beitägt 
        Dim found As Boolean = False
        Dim myCollection As Collection = New Collection
        Dim foundDiagram As clsDiagramm = Nothing
        Dim index As Integer = -1
        Dim tmpCollection As New Collection
        Dim diagrammType As Integer = -1
        Dim currentFilter As New clsFilter

        ' EnableEvents ausschalten ...
        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False

        ' jetzt muss das Chart selber noch markiert werden ...
        Try
            Dim currentSheetName As String = arrWsNames(ptTables.mptPfCharts)
            With CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(currentSheetName), Excel.Worksheet)

                Dim curShape As Excel.Shape = .Shapes.Item(chtObj.Name)
                With curShape.Fill
                    .Visible = MsoTriState.msoTrue
                    .ForeColor.RGB = CInt(awinSettings.glowColor)
                    .Transparency = 0.4
                End With

            End With
        Catch ex As Exception

        End Try

        Try
            foundDiagram = DiagramList.getDiagramm(chtObj.Name)
            If Not IsNothing(foundDiagram) Then
                myCollection = foundDiagram.gsCollection
                If foundDiagram.diagrammTyp = DiagrammTypen(0) Then
                    diagrammType = PTpfdk.Phasen
                ElseIf foundDiagram.diagrammTyp = DiagrammTypen(5) Then
                    diagrammType = PTpfdk.Meilenstein
                ElseIf foundDiagram.diagrammTyp = DiagrammTypen(7) Then
                    diagrammType = PTpfdk.PhaseCategories
                ElseIf foundDiagram.diagrammTyp = DiagrammTypen(8) Then
                    diagrammType = PTpfdk.MilestoneCategories
                End If
                found = True
            End If

        Catch ex As Exception
            myCollection = New Collection
        End Try

        Dim showPhasesMilestones As Boolean = False
        ' es ist ein Rollen Diagramm mit der Menge von angegebenen Rollen in myCollection 

        If found Then
            Dim weitermachen As Boolean = False
            If istRollenDiagramm(chtObj) Then

                ' tk 9.9.18 jetzt sollen alle Kinder- und Kindes-Kinder Rollen gekennzeichnet werden 
                ' es sollen jetzt Sammelrollen durch alle ihre BasicRoles ersetzt werden ... 
                Dim substituteCollection As New Collection

                For Each roleName As String In myCollection
                    If Not substituteCollection.Contains(roleName) Then
                        substituteCollection.Add(roleName, roleName)

                        Dim subRoleIDs As SortedList(Of Integer, Double) = RoleDefinitions.getSubRoleIDsOf(roleName)
                        For Each roleKvP As KeyValuePair(Of Integer, Double) In subRoleIDs
                            Dim childName As String = RoleDefinitions.getRoleDefByID(roleKvP.Key).name
                            If Not substituteCollection.Contains(childName) Then
                                substituteCollection.Add(childName, childName)
                            End If
                        Next
                    End If

                Next

                'currentFilter = New clsFilter("temp", Nothing, Nothing, Nothing, Nothing,
                '                                                myCollection, Nothing)
                currentFilter = New clsFilter("temp", Nothing, Nothing, Nothing, Nothing,
                                                                substituteCollection, Nothing)
                weitermachen = True


            ElseIf istKostenartDiagramm(chtObj) Then
                currentFilter = New clsFilter("temp", Nothing, Nothing, Nothing, Nothing,
                                                                Nothing, myCollection)
                weitermachen = True

            ElseIf istPhasenDiagramm(chtObj) Then
                currentFilter = New clsFilter("temp", Nothing, Nothing, myCollection, Nothing,
                                                               Nothing, Nothing)
                weitermachen = True
                showPhasesMilestones = True

            ElseIf istMileStoneDiagramm(chtObj) Then
                currentFilter = New clsFilter("temp", Nothing, Nothing, Nothing, myCollection,
                                                               Nothing, Nothing)
                weitermachen = True
                showPhasesMilestones = True

            ElseIf istErgebnisDiagramm(chtObj, index) Then

                If index = PTpfdk.UeberAuslastung Or
                    index = PTpfdk.Unterauslastung Then

                    currentFilter = New clsFilter("temp", Nothing, Nothing, Nothing, Nothing,
                                                                    myCollection, Nothing)
                    weitermachen = True
                End If

            End If
            If weitermachen Then

                For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste
                    If currentFilter.doesNotBlock(kvp.Value) Then

                        If kvp.Value.marker = False Then
                            ' dann muss es neu markiert und angezeigt werden 
                            Try
                                Dim idCollection As Collection = Nothing
                                Dim msNames As New Collection
                                Dim phNames As New Collection

                                If showPhasesMilestones Then
                                    If diagrammType = PTpfdk.Meilenstein Or
                                        diagrammType = PTpfdk.MilestoneCategories Then

                                        msNames = myCollection

                                    ElseIf diagrammType = PTpfdk.Phasen Or
                                        diagrammType = PTpfdk.PhaseCategories Then

                                        phNames = myCollection

                                    End If
                                End If

                                kvp.Value.marker = True
                                Dim tmpC As New Collection
                                Call ZeichneProjektinPlanTafel(tmpC, kvp.Value.name, kvp.Value.tfZeile, phNames, msNames)

                            Catch ex As Exception
                                Call MsgBox(ex.Message)
                            End Try
                        End If

                    End If
                Next
            End If
        End If

        ' jetzt muss ggf das BubbleChart Strategie/Risiko neu gezeichnet werden 
        Call awinNeuZeichnenDiagramme(99)

        appInstance.EnableEvents = formerEE

    End Sub


    ''' <summary>
    ''' markiert alle Projekte, wenn die bereits markiert waren, wird atleastOne = false zurückgegeben 
    ''' </summary>
    ''' <param name="atleastOne"></param>
    ''' <remarks></remarks>
    Public Sub markAllProjects(ByRef atleastOne As Boolean)
        Try

            For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

                If kvp.Value.marker = False Then
                    kvp.Value.marker = True
                    atleastOne = True
                    Dim tmpCollection As New Collection
                    Call ZeichneProjektinPlanTafel(tmpCollection, kvp.Value.name, kvp.Value.tfZeile, tmpCollection, tmpCollection)
                End If

            Next

        Catch ex As Exception
            'Call MsgBox(ex.Message)
        End Try

    End Sub

    ''' <summary>
    ''' im Falle Rolle=PortfolioMgr: aggregiert alle Rollen der einzelnen Timestamps in der  Projekthistorie zu den entsprechenden Mutter-Rollen
    ''' </summary>
    ''' <param name="pHist"></param>
    ''' <returns></returns>
    Public Function prepProjectsForRoles(ByVal pHist As clsProjektHistorie) As clsProjektHistorie

        Dim tmpResult As New clsProjektHistorie

        For Each kvp As KeyValuePair(Of Date, clsProjekt) In pHist.liste

            Dim newProj As clsProjekt = prepProjectForRoles(kvp.Value)
            tmpResult.Add(kvp.Key, newProj)

        Next

        For Each kvp As KeyValuePair(Of Date, clsProjekt) In pHist.pfvListe

            Dim newProj As clsProjekt = prepProjectForRoles(kvp.Value)
            tmpResult.AddPfv(newProj)

        Next

        prepProjectsForRoles = tmpResult
    End Function

    Public Function prepProjectsForRoles(ByVal pList As SortedList(Of String, clsProjekt)) As SortedList(Of String, clsProjekt)

        Dim tmpResult As New SortedList(Of String, clsProjekt)

        For Each kvp As KeyValuePair(Of String, clsProjekt) In pList
            Dim newProj As clsProjekt = prepProjectForRoles(kvp.Value)
            tmpResult.Add(kvp.Key, newProj)
        Next

        prepProjectsForRoles = tmpResult

    End Function

    ''' <summary>
    ''' wenn myCustomUserRole = Portfolio Mgr: Ressourcen Zuordnungen müssen aggregiert werden 
    ''' </summary>
    ''' <param name="hproj"></param>
    ''' <returns></returns>
    Public Function prepProjectForRoles(ByVal hproj As clsProjekt) As clsProjekt

        Dim tmpResult As clsProjekt = hproj

        If Not IsNothing(hproj) Then
            ' wenn customUserRole = Portfolio 
            If myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager Then
                Dim idArray() As Integer = myCustomUserRole.getAggregationRoleIDs
                If Not IsNothing(idArray) Then
                    If idArray.Length >= 1 Then
                        tmpResult = hproj.aggregateForPortfolioMgr(idArray)
                    End If
                End If

            End If
        End If

        prepProjectForRoles = tmpResult

    End Function

    ''' <summary>
    ''' prüft, ob es sich um eine Aggregations-Rolle handelt, nur bei Portfolio Mgr relevant;
    ''' in diesem Fall kann in der Hierarchie nicht weiter runtergegangen werden
    ''' </summary>
    ''' <param name="role"></param>
    ''' <returns></returns>
    Public Function isAggregationRole(ByVal role As clsRollenDefinition) As Boolean
        Dim tmpResult As Boolean = False

        If Not IsNothing(role) Then

            If myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager Then
                ' nur dann muss mehr geprüft werden 
                Dim idArray() As Integer = myCustomUserRole.getAggregationRoleIDs

                If Not IsNothing(idArray) Then
                    tmpResult = idArray.Contains(role.UID)
                End If
                'ElseIf myCustomUserRole.customUserRole = ptCustomUserRoles.RessourceManager Then
                '    ' dieser ElseIF Zweig wird aktuell nur für den Allianz Proof of Concept benötigt 
                '    tmpResult = role.isTeam
            ElseIf (myCustomUserRole.customUserRole = ptCustomUserRoles.RessourceManager Or myCustomUserRole.customUserRole = ptCustomUserRoles.TeamManager) Then

                ' tk 2.7.19 das wurde für Allianz Demo eingeführt ; das muss ggf unterchieden werden , ob Normal-Modus oder allianz Modus 
                ' Andernfalls wäre es nicht mehr möglich, einzelnen Team-Membern etwas zuzuweisen
                tmpResult = role.isTeam

            End If
        End If

        isAggregationRole = tmpResult
    End Function

    ''' <summary>
    ''' wenn ein Portfolio Manager speichert, muss der Variant-NAme auf pfv gesetzt werden
    ''' er kann nur Vorgaben speichern, niemals Planungen , also VariantName=""
    ''' entsprechend muss das AlleProjekte upgedated werden
    ''' </summary>
    ''' <param name="projekt"></param>
    Public Sub changeVariantNameAccordingUserRole(ByRef projekt As clsProjekt)
        Dim oldVariantName As String = projekt.variantName
        Dim oldPkey As String = calcProjektKey(projekt)

        Call projekt.setVariantNameAccordingUserRole()
        ' jetzt muss in AlleProjekte ggf der Schlüssel ausgetauscht werden 

        Dim newPkey As String = calcProjektKey(projekt)
        If newPkey <> oldPkey Then
            ' es ergab sich eine Änderung 

            If AlleProjekte.Containskey(oldPkey) Then
                AlleProjekte.Remove(oldPkey)
            End If

            If AlleProjekte.Containskey(newPkey) Then
                AlleProjekte.Remove(newPkey)
            End If

            ' jetzt aufnehmen 
            AlleProjekte.Add(projekt)
        End If
    End Sub

    ''' <summary>
    ''' setzt die Markierungen alle Projekte zurück ...
    ''' wenn die alle schon unmarkiert waren, wird false zurückgegeben, andernfalls true
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub unMarkAllProjects(ByRef atleastOne As Boolean)
        Try

            For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

                If kvp.Value.marker = True Then
                    kvp.Value.marker = False
                    atleastOne = True
                    Dim tmpCollection As New Collection
                    Call ZeichneProjektinPlanTafel(tmpCollection, kvp.Value.name, kvp.Value.tfZeile, tmpCollection, tmpCollection)
                End If

            Next

        Catch ex As Exception
            'Call MsgBox(ex.Message)
        End Try

    End Sub

    ''' <summary>
    ''' gibt denb Default Varianten-NAmen für die angegebene Rolle zurück 
    ''' </summary>
    ''' <returns></returns>
    Public Function getDefaultVariantNameAccordingUserRole() As String
        Dim tmpResult As String = ""

        If myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager Then
            tmpResult = ptVariantFixNames.pfv.ToString

        End If

        getDefaultVariantNameAccordingUserRole = tmpResult
    End Function


    ''' <summary>
    ''' prüft , ob übergebenes Diagramm ein Rollen Diagramm ist - in R steht ggf als Ergebnis die entsprechende Rollen-Nummer; 0 wenn es kein Rollen Diagramm ist
    ''' </summary>
    ''' <param name="chtobj"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
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
                If (vglValue >= 3 And vglValue <= 11) Or
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

                If CInt(tmpStr(1)) = PTpfdk.Phasen Or
                    CInt(tmpStr(1)) = PTpfdk.PhaseCategories Then

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

                If CInt(tmpStr(1)) = PTpfdk.Meilenstein Or
                    CInt(tmpStr(1)) = PTpfdk.MilestoneCategories Then

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

                ' ''If CInt(tmpStr(1)) = PTpfdk.FitRisiko Or _
                ' ''    CInt(tmpStr(1)) = PTpfdk.FitRisikoVol Or _
                ' ''    CInt(tmpStr(1)) = PTpfdk.ComplexRisiko Or _
                ' ''    CInt(tmpStr(1)) = PTpfdk.Dependencies Or _
                ' ''    CInt(tmpStr(1)) = PTpfdk.ZeitRisiko Then

                found = True

                ''End If

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



    ' tk 3.5.19 de-aktiviert
    'Sub awinRightClickinPortfolioAendern()
    '    Dim myBar As CommandBar
    '    Dim myitem As CommandBarButton
    '    'Dim myitem As CommandBarControl
    '    Dim i As Integer, endofsearch As Integer
    '    Dim found As Boolean
    '    Dim awinevent As clsEventsPfCharts

    '    found = False
    '    i = 1

    '    With appInstance.CommandBars
    '        endofsearch = .Count

    '        While i <= endofsearch And Not found
    '            If .Item(i).Name = "awinRightClickinPortfolio" Then
    '                found = True
    '            Else
    '                i = i + 1
    '            End If
    '        End While
    '    End With

    '    If found Then
    '        Exit Sub
    '    End If

    '    'CommandBars.Item.Name
    '    myBar = appInstance.CommandBars.Add(Name:="awinRightClickinPortfolio", Position:=MsoBarPosition.msoBarPopup, Temporary:=True)


    '    ' Add a menu item
    '    myitem = CType(myBar.Controls.Add(Type:=MsoControlType.msoControlButton), Microsoft.Office.Core.CommandBarButton)
    '    With myitem
    '        .Caption = "Umbenennen"
    '        .Tag = "Umbenennen"
    '        '.OnAction = "awinRenameProject"
    '    End With
    '    'awinevent = New clsAwinEvent
    '    'awinevent.Button3Events = myitem
    '    awinevent = New clsEventsPfCharts
    '    awinevent.PfChartRightClick = myitem
    '    awinButtonEvents.Add(awinevent)


    '    ' Add a menu item
    '    myitem = CType(myBar.Controls.Add(Type:=MsoControlType.msoControlButton), Microsoft.Office.Core.CommandBarButton)
    '    With myitem
    '        .Caption = "Löschen"
    '        .Tag = "Loesche aus Portfolio"
    '        '.OnAction = "awinDeleteChartorProject"
    '    End With
    '    'awinevent = New clsAwinEvent
    '    'awinevent.Button3Events = myitem
    '    awinevent = New clsEventsPfCharts
    '    awinevent.PfChartRightClick = myitem
    '    awinButtonEvents.Add(awinevent)

    '    ' Add a menu item
    '    myitem = CType(myBar.Controls.Add(Type:=MsoControlType.msoControlButton), Microsoft.Office.Core.CommandBarButton)
    '    With myitem
    '        .Caption = "Show / Noshow"
    '        .Tag = "Show / Noshow"
    '        '.OnAction = "awinShowNoShowProject"
    '    End With
    '    'awinevent = New clsAwinEvent
    '    'awinevent.Button3Events = myitem
    '    awinevent = New clsEventsPfCharts
    '    awinevent.PfChartRightClick = myitem
    '    awinButtonEvents.Add(awinevent)

    '    ' Add a menu item
    '    myitem = CType(myBar.Controls.Add(Type:=MsoControlType.msoControlButton), Microsoft.Office.Core.CommandBarButton)
    '    With myitem
    '        .Caption = "Bearbeiten Projekt-Attribute"
    '        .Tag = "Bearbeiten Projekt-Attribute"
    '        '.OnAction = "awinEditDataProject"
    '    End With
    '    'awinevent = New clsAwinEvent
    '    'awinevent.Button3Events = myitem
    '    awinevent = New clsEventsPfCharts
    '    awinevent.PfChartRightClick = myitem
    '    awinButtonEvents.Add(awinevent)

    '    ' Add a menu item
    '    myitem = CType(myBar.Controls.Add(Type:=MsoControlType.msoControlButton), Microsoft.Office.Core.CommandBarButton)
    '    With myitem
    '        .Caption = "Beauftragen"
    '        .Tag = "Beauftragen"
    '        '.OnAction = "awinBeauftrageProject"
    '    End With
    '    'awinevent = New clsAwinEvent
    '    'awinevent.Button3Events = myitem
    '    awinevent = New clsEventsPfCharts
    '    awinevent.PfChartRightClick = myitem
    '    awinButtonEvents.Add(awinevent)

    'End Sub

    '''' <summary>
    '''' aktiviert die Right Clicks in den Charts 
    '''' </summary>
    '''' <remarks></remarks>
    'Sub awinRightClickinPRCCharts()
    '    Dim myBar As CommandBar
    '    Dim myitem As CommandBarButton
    '    Dim i As Integer, endofsearch As Integer
    '    Dim found As Boolean
    '    'Dim awinevent As clsAwinEvents
    '    Dim awinevent As clsEventsPrcCharts

    '    found = False
    '    i = 1

    '    With appInstance.CommandBars
    '        endofsearch = .Count

    '        While i <= endofsearch And Not found
    '            If .Item(i).Name = "awinRightClickinPRCChart" Then
    '                found = True
    '            Else
    '                i = i + 1
    '            End If
    '        End While
    '    End With

    '    If found Then
    '        Exit Sub
    '    End If

    '    'CommandBars.Item.Name
    '    myBar = appInstance.CommandBars.Add(Name:="awinRightClickinPRCChart", Position:=MsoBarPosition.msoBarPopup, Temporary:=True)


    '    ' Add a menu item
    '    myitem = CType(myBar.Controls.Add(Type:=MsoControlType.msoControlButton), Microsoft.Office.Core.CommandBarButton)
    '    With myitem
    '        .Caption = "Löschen"
    '        .Tag = "Löschen"
    '        '.OnAction = "awinDeleteChartorProject"
    '    End With

    '    'awinevent = New clsAwinEvent
    '    'awinevent.Button4Events = myitem
    '    awinevent = New clsEventsPrcCharts
    '    awinevent.PrcChartRightClick = myitem
    '    awinButtonEvents.Add(awinevent)

    '    ' Add a menu item
    '    myitem = CType(myBar.Controls.Add(Type:=MsoControlType.msoControlButton), Microsoft.Office.Core.CommandBarButton)
    '    With myitem
    '        .Caption = "Röntgenblick ein/aus"
    '        .Tag = "Bedarf anzeigen"
    '        '.OnAction = "awinShowNeedsOfProjects"
    '    End With
    '    'awinevent = New clsAwinEvent
    '    'awinevent.Button4Events = myitem
    '    awinevent = New clsEventsPrcCharts
    '    awinevent.PrcChartRightClick = myitem
    '    awinButtonEvents.Add(awinevent)

    '    ' Add a menu item
    '    myitem = CType(myBar.Controls.Add(Type:=MsoControlType.msoControlButton), Microsoft.Office.Core.CommandBarButton)
    '    With myitem
    '        .Caption = "nach Freiheitsgraden optimieren"
    '        .Tag = "Optimieren"
    '        '.OnAction = "awinOptimizeStartOfProjects"
    '    End With
    '    'awinevent = New clsAwinEvent
    '    'awinevent.Button4Events = myitem
    '    awinevent = New clsEventsPrcCharts
    '    awinevent.PrcChartRightClick = myitem
    '    awinButtonEvents.Add(awinevent)

    '    ' ergänzt am 2.11.2014
    '    ' Add a menu item
    '    myitem = CType(myBar.Controls.Add(Type:=MsoControlType.msoControlButton), Microsoft.Office.Core.CommandBarButton)
    '    With myitem
    '        .Caption = "nach Varianten optimieren"
    '        .Tag = "Varianten optimieren"
    '        '.OnAction = "awinOptimizeStartOfProjects"
    '    End With
    '    'awinevent = New clsAwinEvent
    '    'awinevent.Button4Events = myitem
    '    awinevent = New clsEventsPrcCharts
    '    awinevent.PrcChartRightClick = myitem
    '    awinButtonEvents.Add(awinevent)

    'End Sub

    'Sub awinKontextReset()

    '    Try
    '        appInstance.CommandBars("awinRightClickinPortfolio").Delete()
    '    Catch ex As Exception

    '    End Try

    '    Try
    '        appInstance.CommandBars("awinRightClickinPRCChart").Delete()
    '    Catch ex As Exception

    '    End Try


    '    ' die Short Cut Menues aus Excel wieder alle aktivieren ...
    '    'Dim cbar As CommandBar

    '    'For Each cbar In appInstance.CommandBars

    '    '    cbar.Enabled = True
    '    '    'Try
    '    '    '    cbar.Reset()
    '    '    'Catch ex As Exception

    '    '    'End Try

    '    'Next


    'End Sub



    '

    '
    ''' <summary>
    ''' gibt die Überdeckung zurück zwischen den beiden Zeiträumen definiert durch showRangeLeft /showRangeRight und anfang / ende
    ''' </summary>
    ''' <param name="anfang">Anfang Zeitraum 2</param>
    ''' <param name="ende">Ende Zeitraum 2</param>
    ''' <param name="ixZeitraum">gibt an , in welchem Monat des Zeitraums die Überdeckung anfängt: 0 = 1. Monat</param>
    ''' <param name="ix">gibt an, in welchem Monat des durch Anfang / Ende definierten Zeitraums die Überdeckung anfängt</param>
    ''' <param name="anzahl">enthält die Breite der Überdeckung</param>
    ''' <remarks></remarks>
    Sub awinIntersectZeitraum(anfang As Integer, ende As Integer,
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

        With CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.MPT)), Excel.Worksheet)

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

        With appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.MPT))
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
                        If (chtTitle = DiagramList.getDiagramm(i).DiagrammTitel) And
                                        (DiagramList.getDiagramm(i).isCockpitChart = True) And
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



        With appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.MPT))

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
            If (chtTitle = DiagramList.getDiagramm(i).DiagrammTitel) And
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
            worksheetShapes = CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.MPT)), Excel.Worksheet).Shapes
            ReDim shpArray(selectedProjekte.Count - 1)

            For Each kvp In selectedProjekte.Liste

                hproj = kvp.Value
                i = i + 1
                Try
                    shpElement = CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.MPT)), Excel.Worksheet).Shapes.Item(hproj.name)
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
    Sub awinDeSelect(Optional ByVal selectDummyCell As Boolean = False)
        Dim srow As Integer = 1
        'Dim hziel As Integer
        'Dim vziel As Integer


        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False


        ' Selektierte Projekte auf Null setzen 

        If selectedProjekte.Count > 0 Then
            selectedProjekte.Clear(False)
            If awinSettings.showValuesOfSelected Then
                Call awinNeuZeichnenDiagramme(8)
            End If

        End If


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



    Public Function magicBoardIstFrei(ByVal mycollection As Collection, ByVal pname As String, ByVal zeile As Integer,
                                      ByVal startDate As Date, ByVal laenge As Integer, ByVal anzahlZeilen As Integer) As Boolean
        Dim istfrei = True
        Dim ix As Integer = 1
        Dim anzahlP As Integer = ShowProjekte.Count


        For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

            If pname <> kvp.Key And Not mycollection.Contains(kvp.Key) And kvp.Value.shpUID <> "" Then
                With kvp.Value
                    If .tfZeile >= zeile And .tfZeile <= zeile + anzahlZeilen - 1 Then
                        If startDate.Date <= .startDate.Date Then
                            If startDate.AddDays(laenge - 1).Date > .startDate.Date Then
                                istfrei = False
                                Exit For
                            Else
                                istfrei = True
                            End If
                        ElseIf startDate < .endeDate Then
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

            If Not magicBoardIstFrei(mycollection, pname, zeile, hproj.startDate, hproj.dauerInDays, anzahlzeilen) Then
                tryoben = zeile - 1
                tryunten = zeile + 1

                ' jetzt ggf eine neue Position für das Shape suchen - dabei iterierend unten bzw oben suchen
                zeile = tryunten
                lookDown = True

                While Not magicBoardIstFrei(mycollection, pname, zeile, hproj.startDate, hproj.dauerInDays, anzahlzeilen)
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
                ' tk 9.4.19 nicht mehr zeigen 
                'formMilestone.Visible = False
                'formStatus.Visible = False
                'formPhase.Visible = False

                typCollection.Add(CInt(PTshty.milestoneN).ToString, CInt(PTshty.milestoneN).ToString)
                typCollection.Add(CInt(PTshty.phaseN).ToString, CInt(PTshty.phaseN).ToString)

            Case 1
                ' tk 9.4.19 nicht mehr zeigen
                'formMilestone.Visible = False
                typCollection.Add(CInt(PTshty.milestoneN).ToString, CInt(PTshty.milestoneN).ToString)

            Case 2
                ' tk 9.4.19 nicht mehr zeigen
                'formStatus.Visible = False

            Case 3
                ' tk 9.4.19 nicht mehr zeigen
                'formPhase.Visible = False
                typCollection.Add(CInt(PTshty.phaseN).ToString, CInt(PTshty.phaseN).ToString)

            Case Else
                appInstance.EnableEvents = formerEE
                enableOnUpdate = formereO
                Exit Sub
        End Select

        Try
            worksheetShapes = CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.MPT)), Excel.Worksheet).Shapes



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
            worksheetShapes = CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.MPT)), Excel.Worksheet).Shapes

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
        Dim minvorne As Boolean = False, minhinten As Boolean = False,
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
    Public Function calcHryFullname(ByVal elemName As String, ByVal breadcrumb As String,
                                    Optional ByVal pvKennung As String = "") As String

        If pvKennung = "" Then
            If breadcrumb = "" Then
                calcHryFullname = elemName
            Else
                calcHryFullname = breadcrumb & "#" & elemName
            End If

        Else
            If breadcrumb = "" Then
                calcHryFullname = "[" & pvKennung & "]" & elemName
            Else
                calcHryFullname = "[" & pvKennung & "]" & breadcrumb & "#" & elemName
            End If
        End If


    End Function

    ''' <summary>
    ''' berechnet den Namen, der in selectedphases bzw. selectedMilestones reinkommt, wenn awinsettings.considercategory true ist  
    ''' Breadcrumb und elemName; Breadcrumb und die einzelnen Stufen des Breadcrumbs sind getrennt durch #
    ''' </summary>
    ''' <param name="category">bezeichnet die Darstellungsklasse des Meilensteins , der Phase</param>
    ''' <param name="isMilestone">gibt an ob es sich um einen Meilenstein oder eine Phase handelt</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function calcHryCategoryName(ByVal category As String, ByVal isMilestone As Boolean) As String

        If isMilestone Then
            calcHryCategoryName = "[C:" & category & "]M"
        Else
            calcHryCategoryName = "[C:" & category & "]P"
        End If

    End Function

    ''' <summary>
    ''' bestimmt den eindeutigen Namen des Shapes für einen Meilenstein oder eine Phase 
    ''' der Name enthält pName, vname und ElemID : (pname#vname)ElemID
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

        If tmpName.Length > 240 Then
            If awinSettings.englishLanguage Then
                Call MsgBox("Name too long: " & tmpName & vbLf & tmpName.Length & " characters; please reduce to less than 240 characters")
            Else
                Call MsgBox("Name ist zu lang: " & tmpName & vbLf & tmpName.Length & " Zeichen; bitte auf weniger als 240 Zeichen reduzieren")
            End If

            tmpName = tmpName.Substring(0, 240)
        End If

        calcPPTShapeName = tmpName

    End Function

    ''' <summary>
    ''' gibt den Elem-Name und Breadcrumb als einzelne Strings zurück
    ''' es kann unterschieden werden zwischen [P:Projekt-Name], 
    ''' [V:Vorlagen-Name] und [C:Category-Name]
    ''' </summary>
    ''' <param name="fullname"></param>
    ''' <param name="elemName"></param>
    ''' <param name="breadcrumb"></param>
    ''' <remarks></remarks>
    Public Sub splitHryFullnameTo2(ByVal fullname As String,
                                   ByRef elemName As String, ByRef breadcrumb As String,
                                   ByRef type As Integer, ByRef pvName As String)
        Dim tmpstr() As String
        Dim tmpBC As String = ""
        Dim anzahl As Integer

        ' enthält der pvName die Kennung für Vorlage oder Projekt ? 
        If fullname.StartsWith("[P:") Or fullname.StartsWith("[V:") Or
            fullname.StartsWith("[C:") Then
            If fullname.StartsWith("[P:") Then
                type = PTItemType.projekt
            ElseIf fullname.StartsWith("[V:") Then
                type = PTItemType.vorlage
            Else
                type = PTItemType.categoryList
            End If

            Dim startPos As Integer = 3
            Dim endPos As Integer = fullname.IndexOf("]") + 1
            pvName = fullname.Substring(startPos, endPos - startPos - 1)

            fullname = fullname.Substring(endPos)
        Else
            type = -1
            pvName = ""
        End If

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
    ''' zerhackt den übergebenen String in seine Bestandteile [C:KAtegorie], [V:vorlagen-name] bzw. [P:projekt-name] und 
    ''' Breadcrumb-Name
    ''' 
    ''' </summary>
    ''' <param name="fullname"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function splitHryFullnameTo1(ByVal fullname As String) As String
        Dim tmpResult As String = ""
        Dim elemName As String = ""
        Dim breadCrumb As String = ""
        Dim type As Integer = -1
        Dim pvName As String = ""

        Call splitHryFullnameTo2(fullname, elemName, breadCrumb, type, pvName)

        If type = PTItemType.categoryList Then
            ' hier steht im pvName der Name der Kategorie ...
            tmpResult = pvName
        Else
            If breadCrumb = "" Then
                tmpResult = elemName
            Else
                tmpResult = breadCrumb.Replace("#", "-") & "-" & elemName
            End If
        End If


        splitHryFullnameTo1 = tmpResult

    End Function

    ''' <summary>
    ''' zerteilt einen String, der folgendes Format hat: breadcrumb#elemName#lfdnr in seine Bestandteile 
    ''' </summary>
    ''' <param name="fullname"></param>
    ''' <param name="elemName"></param>
    ''' <param name="breadcrumb"></param>
    ''' <param name="lfdNr"></param>
    ''' <remarks></remarks>
    Public Sub splitBreadCrumbFullnameTo3(ByVal fullname As String, ByRef elemName As String, ByRef breadcrumb As String, ByRef lfdNr As Integer,
                                          ByRef type As Integer, ByRef pvName As String)
        Dim tmpstr() As String
        Dim tmpBC As String = ""
        Dim anzahl As Integer

        tmpstr = fullname.Split(New Char() {CChar("#")}, 20)
        anzahl = tmpstr.Length
        If tmpstr.Length = 1 Then
            elemName = tmpstr(0)
            breadcrumb = ""
            lfdNr = 1
            type = -1
            pvName = ""
        ElseIf tmpstr.Length > 1 Then
            lfdNr = CInt(tmpstr(anzahl - 1))
            For i As Integer = 0 To anzahl - 2
                If i = 0 Then
                    tmpBC = tmpstr(i)
                Else
                    tmpBC = tmpBC & "#" & tmpstr(i)
                End If
            Next

            Call splitHryFullnameTo2(tmpBC, elemName, breadcrumb, type, pvName)
        Else
            elemName = "?"
            breadcrumb = ""
            lfdNr = 0
            type = -1
            pvName = ""
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

    ''' <summary>
    ''' gibt zurück, ob ein Projekt-NAme gültig ist oder nicht 
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function isValidProjectName(ByVal pName As String) As Boolean

        If IsNothing(pName) Then
            isValidProjectName = False
        Else
            pName = pName.Trim
            If pName.Length >= 1 Then
                If pName.Contains("#") Or
                    pName.Contains("(") Or pName.Contains(")") Or
                    pName.Contains(vbLf) Or pName.Contains(vbCr) Then
                    isValidProjectName = False
                Else
                    isValidProjectName = True
                End If
            Else
                isValidProjectName = False
            End If
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


    ' '' Änderung tk - das wird nicht mehr benötigt - budgetwerte wird jetzt immer berechnet
    '' ''' <summary>
    '' ''' erzeugt die monatlichen Budget Werte für ein Projekt
    '' ''' berechnet aus dem Wert für Erloes, verteilt nach einem Schlüssel, der sich aus Marge und Kostenbedarf ergibt 
    '' ''' </summary>
    '' ''' <param name="hproj"></param>
    '' ''' <remarks></remarks>

    'Public Sub awinCreateBudgetWerte(ByRef hproj As clsProjekt)


    '    Dim costValues() As Double, budgetValues() As Double
    '    Dim curBudget As Double, avgbudget As Double

    '    ' Ergänzung am 26.5.14: wenn hproj in den Längen der Bedarfe Arrays nicht konsistent ist: 
    '    ' anpassen 
    '    If Not hproj.isConsistent Then
    '        Call hproj.syncXWertePhases()
    '    End If

    '    costValues = hproj.getGesamtKostenBedarf
    '    ReDim budgetValues(costValues.Length - 1)

    '    curBudget = hproj.Erloes
    '    avgbudget = curBudget / costValues.Length

    '    If curBudget > 0 Then
    '        If costValues.Sum > 0 Then
    '            Dim pMarge As Double = hproj.ProjectMarge
    '            For i = 0 To costValues.Length - 1
    '                budgetValues(i) = costValues(i) * (1 + pMarge)
    '            Next
    '        Else
    '            For i = 0 To costValues.Length - 1
    '                budgetValues(i) = avgbudget
    '            Next
    '        End If
    '    End If


    '    hproj.budgetWerte = budgetValues


    'End Sub

    ' nicht mehr notwendig, budgetWerte ist jetzt readonly Eigenschaft ... 
    '' ''' <summary>
    '' ''' aktualisiert die Budget werte , wobei die Charakteristik erhalten bleibt 
    '' ''' Vorbedingung ist, daß das bisherige Budget > 0 Null ist 
    '' ''' </summary>
    '' ''' <param name="hproj"></param>
    '' ''' <param name="newBudget">Gesamt Wert des neuen Budgets</param>
    '' ''' <remarks></remarks>
    ''Public Sub awinUpdateBudgetWerte(ByRef hproj As clsProjekt, ByVal newBudget As Double)



    ''    Dim curValues() As Double, budgetValues() As Double
    ''    Dim oldBudget As Double
    ''    Dim faktor As Double

    ''    curValues = hproj.budgetWerte
    ''    ReDim budgetValues(curValues.Length - 1)
    ''    oldBudget = curValues.Sum

    ''    If oldBudget = 0 Then
    ''        Throw New Exception("altes Budget darf beim Update nicht Null sein")
    ''    Else
    ''        If newBudget <= 0 Then
    ''            ' budgetvalues ist bereits auf Null gesetzt  
    ''        Else
    ''            faktor = newBudget / oldBudget
    ''            For i = 0 To curValues.Length - 1
    ''                budgetValues(i) = curValues(i) * faktor
    ''            Next
    ''        End If

    ''    End If

    ''    hproj.budgetWerte = budgetValues

    ''End Sub

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

    ''' <summary>
    ''' verteilt die Summe der 
    ''' </summary>
    ''' <param name="startdate"></param>
    ''' <param name="endedate"></param>
    ''' <param name="oldXwerte"></param>
    ''' <param name="corrFakt"></param>
    ''' <returns></returns>
    Public Function calcVerteilungAufMonate(ByVal startdate As Date, ByVal endedate As Date,
                                            ByVal oldXwerte() As Double, ByVal corrFakt As Double) As Double()
        Dim k As Integer
        Dim newXwerte() As Double
        Dim gesBedarf As Double
        Dim Rest As Integer
        Dim hDatum As Date
        Dim anzDaysthisMonth As Double
        Dim newLength As Integer = getColumnOfDate(endedate) - getColumnOfDate(startdate) + 1
        Dim gesBedarfReal As Double = 0.0
        Dim dauerinDays As Long = DateDiff(DateInterval.Day, startdate, endedate) + 1

        ReDim newXwerte(newLength - 1)

        ' Vorbereitung für Summen Berechnung nur bei Forecast
        'Dim hasActualData As Boolean = Me.parentProject.actualDataUntil <> Date.MinValue
        'Dim actualDataColumn As Integer = -1

        'If hasActualData Then
        '    actualDataColumn = getColumnOfDate(Me.parentProject.actualDataUntil)
        'End If

        ' nur wenn überhaupt was zu verteilen ist, muss alles folgende gemacht werdne 
        ' andernfalls ist eh schon alles richtig 
        If oldXwerte.Sum > 0 Then

            Try

                gesBedarfReal = oldXwerte.Sum * corrFakt
                gesBedarf = System.Math.Round(gesBedarfReal)


                If newLength = oldXwerte.Length Then

                    'Bedarfe-Verteilung bleibt wie gehabt ... allerdings unter Berücksichtigung corrFakt


                    For i = 0 To newLength - 1
                        newXwerte(i) = oldXwerte(i) * corrFakt
                    Next

                    ' jetzt ggf die Reste verteilen 
                    Rest = CInt(gesBedarf - newXwerte.Sum)

                    k = newXwerte.Length - 1
                    While Rest <> 0

                        If Rest > 0 Then
                            newXwerte(k) = newXwerte(k) + 1
                            Rest = Rest - 1
                        Else

                            If newXwerte(k) - 1 >= 0 Then
                                newXwerte(k) = newXwerte(k) - 1
                                Rest = Rest + 1
                            End If

                        End If
                        k = k - 1
                        If k < 0 Then
                            k = newXwerte.Length - 1
                        End If

                    End While

                    ' letzter Test: wenn jetzt durch die Rundungen immer noch ein abs(Rest) von < 1 ist 
                    k = newXwerte.Length - 1
                    If newXwerte.Sum <> gesBedarfReal Then
                        Dim RestDbl As Double = gesBedarfReal - newXwerte.Sum
                        If Math.Abs(RestDbl) <= 1 And Math.Abs(RestDbl) >= 0 Then
                            ' alles ok 

                            ' positioniere auf ein k, dessen Wert größer ist als abs(restdbl) 
                            Do While newXwerte(k) < Math.Abs(RestDbl) And k > 0
                                k = k - 1
                            Loop
                            ' jetzt ist ein k erreicht 
                            newXwerte(k) = newXwerte(k) + RestDbl
                            If newXwerte(k) < 0 Then
                                newXwerte(k) = 0.0 ' darf eigentlich nie passieren ..
                            End If

                        Else
                            Dim a As Double = RestDbl ' kann / darf eigentlich nicht sein 
                        End If
                    End If


                Else

                    Dim tmpSum As Double = 0
                    For k = 0 To newXwerte.Length - 1

                        If k = 0 Then
                            ' damit ist 00:00 des Startdates gemeint 
                            hDatum = startdate

                            anzDaysthisMonth = DateDiff(DateInterval.Day, hDatum, hDatum.AddDays(-1 * hDatum.Day + 1).AddMonths(1))

                            'anzDaysthisMonth = DateDiff("d", hDatum, DateSerial(hDatum.Year, hDatum.Month + 1, hDatum.Day))
                            'anzDaysthisMonth = anzDaysthisMonth - DateDiff("d", DateSerial(hDatum.Year, hDatum.Month, 1), hDatum) - 1

                        ElseIf k = newXwerte.Length - 1 Then
                            ' damit hDatum das End-Datum um 23.00 Uhr

                            anzDaysthisMonth = endedate.Day
                            'hDatum = endedate.AddHours(23)
                            'anzDaysthisMonth = DateDiff("d", DateSerial(hDatum.Year, hDatum.Month, 1), hDatum)

                        Else
                            hDatum = startdate
                            anzDaysthisMonth = DateDiff(DateInterval.Day, startdate.AddMonths(k), startdate.AddMonths(k + 1))
                            'anzDaysthisMonth = DateDiff("d", DateSerial(hDatum.Year, hDatum.Month + k, hDatum.Day), DateSerial(hDatum.Year, hDatum.Month + k + 1, hDatum.Day))
                        End If

                        newXwerte(k) = System.Math.Round(anzDaysthisMonth / (dauerinDays * corrFakt) * gesBedarf)
                        tmpSum = tmpSum + anzDaysthisMonth
                    Next k

                    ' Kontrolle für Test ... aChck muss immer Null sein !
                    'Dim aChck As Double = Me.dauerInDays - tmpSum


                    ' Rest wird auf alle newXwerte verteilt

                    Rest = CInt(gesBedarf - newXwerte.Sum)

                    k = newXwerte.Length - 1
                    While Rest <> 0
                        If Rest > 0 Then
                            newXwerte(k) = newXwerte(k) + 1
                            Rest = Rest - 1
                        Else
                            If newXwerte(k) - 1 >= 0 Then
                                newXwerte(k) = newXwerte(k) - 1
                                Rest = Rest + 1
                            End If
                        End If
                        k = k - 1
                        If k < 0 Then
                            k = newXwerte.Length - 1
                        End If

                    End While

                    ' letzter Test: wenn jetzt durch die Rundungen immer noch ein abs(Rest) von < 1 ist 
                    k = newXwerte.Length - 1
                    If newXwerte.Sum <> gesBedarfReal Then
                        Dim RestDbl As Double = gesBedarfReal - newXwerte.Sum
                        If Math.Abs(RestDbl) <= 1 And Math.Abs(RestDbl) >= 0 Then
                            ' alles ok 

                            ' positioniere auf ein k, dessen Wert größer ist als abs(restdbl) 
                            Do While newXwerte(k) < Math.Abs(RestDbl) And k > 0
                                k = k - 1
                            Loop
                            ' jetzt ist ein k erreicht 
                            newXwerte(k) = newXwerte(k) + RestDbl
                            If newXwerte(k) < 0 Then
                                newXwerte(k) = 0.0 ' darf eigentlich nie passieren ..
                            End If

                        Else
                            Dim a As Double = RestDbl ' kann / darf eigentlich nicht sein 
                        End If
                    End If

                End If



            Catch ex As Exception

            End Try

        Else
            ' alles auf Null setzen 
            For ix = 0 To newLength - 1
                newXwerte(ix) = 0
            Next
        End If

        calcVerteilungAufMonate = newXwerte

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
    Public Function calcPhaseUeberdeckung(ByVal startDate1 As Date, endDate1 As Date,
                                              ByVal startDate2 As Date, ByVal enddate2 As Date) As Double

        Dim duration1 As Long = DateDiff(DateInterval.Day, startDate1, endDate1) + 1
        Dim duration2 As Long = DateDiff(DateInterval.Day, startDate2, enddate2) + 1
        Dim ueberdeckungsStart As Date, ueberdeckungsEnde As Date
        Dim ueberdeckungsduration As Long
        Dim ergebnis As Double


        If DateDiff(DateInterval.Day, endDate1, startDate2) > 0 Or
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

        If Len(strName) < 2 Then
            ' ProjektName soll mehr als 1 Zeichen haben
            found = True
        ElseIf AlleProjekte.Containskey(key) Then
            found = True

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

            If completeText.Contains(appearanceDefinitions.ElementAt(index).Key.Trim) And
                    isMilestone = appearanceDefinitions.ElementAt(index).Value.isMilestone Then
                found = True
                ergebnis = appearanceDefinitions.ElementAt(index).Key
            Else
                index = index + 1
            End If

        Loop

        mapToAppearance = ergebnis

    End Function


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
    Public Sub retrieveSelections(ByVal fName As String, ByVal menuOption As Integer,
                                       ByRef selectedBUs As Collection, ByRef selectedTyps As Collection,
                                       ByRef selectedPhases As Collection, ByRef selectedMilestones As Collection,
                                       ByRef selectedRoles As Collection, ByRef selectedCosts As Collection)

        Dim lastFilter As clsFilter

        If menuOption = PTmenue.filterdefinieren Or
            menuOption = PTmenue.sessionFilterDefinieren Or
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
    ''' Gibt zurück Selectiontyp der aktuell selektierten Elemente
    ''' 0 = Projekt-Struktur (Vorlage)
    ''' 1 = Projekt-Struktur(Projekt)
    ''' 2 = Namensliste
    ''' 3 = Category Liste
    ''' </summary>
    ''' <param name="selPh"></param>
    ''' <param name="selMs"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function selectionTyp(ByVal selPh As Collection, ByVal selMs As Collection) As Integer

        Dim element As String = ""
        Dim tmpresult As Integer = PTItemType.nameList
        Dim i As Integer = 1

        ' Check ob selectedPhases P:, V: oder C: enthält
        Do While tmpresult <> PTItemType.projekt And i <= selPh.Count

            element = selPh.Item(i).ToString

            Dim elemName As String = ""
            Dim bc As String = ""
            Dim tmpType As Integer = -1
            Dim pvcName As String = ""
            Call splitHryFullnameTo2(element, elemName, bc, tmpType, pvcName)

            If tmpType = PTItemType.vorlage Or tmpType = PTItemType.projekt Then
                tmpresult = tmpType
            ElseIf tmpresult = PTItemType.nameList And tmpType = PTItemType.categoryList Then
                tmpresult = tmpType
            End If

            i = i + 1
        Loop


        ' Schleife ist nur solange notwendig, solange tmpResult nicht gleich Projekt-Typ ist 
        i = 1
        Do While tmpresult <> PTItemType.projekt And i <= selMs.Count

            element = selMs.Item(i).ToString

            Dim elemName As String = ""
            Dim bc As String = ""
            Dim tmpType As Integer = -1
            Dim pvcName As String = ""
            Call splitHryFullnameTo2(element, elemName, bc, tmpType, pvcName)

            If tmpType = PTItemType.vorlage Or tmpType = PTItemType.projekt Then
                tmpresult = tmpType
            ElseIf tmpresult = PTItemType.nameList And tmpType = PTItemType.categoryList Then
                tmpresult = tmpType
            End If

            i = i + 1

        Loop


        ' gleiches noch todo für Milensteine, rollen, kosten, ...

        selectionTyp = tmpresult

    End Function


    ''' <summary>
    ''' kennzeichnet ein Powerpoint Slide als ein Slide, das Smart Elements enthält 
    ''' fügt die Kennzeichnung "SMART" mit type an, StartofCalendar, CreationDate und Datenbank Infos 
    ''' </summary>
    ''' <param name="pptSlide"></param>
    ''' <remarks></remarks>
    Public Sub addSmartPPTSlideCalInfo(ByRef pptSlide As PowerPoint.Slide,
                                       ByVal calendarLeft As Date,
                                       ByVal calendarRight As Date,
                                       Optional ByVal projectTimeStamp As Date = Nothing)

        If Not IsNothing(pptSlide) Then
            With pptSlide


                If .Tags.Item("SMART").Length > 0 Then
                    ' es muss nichts mehr gemacht werden, es ist bereoits gekennzeichnet 
                    '.Tags.Delete("SMART")
                Else
                    .Tags.Add("SMART", "visbo")
                End If

                If .Tags.Item("SOC").Length > 0 Then
                    .Tags.Delete("SOC")
                End If
                .Tags.Add("SOC", StartofCalendar.ToShortDateString)

                If .Tags.Item("CALL").Length > 0 Then
                    .Tags.Delete("CALL")
                End If
                .Tags.Add("CALL", calendarLeft.ToShortDateString)

                If .Tags.Item("CALR").Length > 0 Then
                    .Tags.Delete("CALR")
                End If
                .Tags.Add("CALR", calendarRight.ToShortDateString)


            End With
        End If


    End Sub

    ''' <summary>
    ''' kennzeichnet die Seite als Smart VISBO Seite 
    ''' </summary>
    ''' <param name="pptSlide"></param>
    ''' <param name="projectTimeStamp"></param>
    ''' <remarks></remarks>
    Public Sub addSmartPPTSlideBaseInfo(ByRef pptSlide As PowerPoint.Slide,
                                        ByVal projectTimeStamp As Date,
                                        ByVal type As Integer)

        If Not IsNothing(pptSlide) Then
            With pptSlide

                If .Tags.Item("PRPF").Length > 0 Then
                    .Tags.Delete("PRPF")
                End If
                .Tags.Add("PRPF", type.ToString)

                If .Tags.Item("SMART").Length > 0 Then
                    .Tags.Delete("SMART")
                End If
                .Tags.Add("SMART", "visbo")

                If IsNothing(projectTimeStamp) Then

                    projectTimeStamp = Date.Now
                ElseIf projectTimeStamp = Date.MinValue Then
                    projectTimeStamp = Date.Now
                End If

                If .Tags.Item("SOC").Length > 0 Then
                    .Tags.Delete("SOC")
                End If
                .Tags.Add("SOC", StartofCalendar.ToShortDateString)


                If .Tags.Item("CRD").Length > 0 Then
                    .Tags.Delete("CRD")
                End If
                .Tags.Add("CRD", projectTimeStamp.ToString)


                If awinSettings.databaseURL.Length > 0 And awinSettings.databaseName.Length > 0 Then
                    If awinSettings.databaseURL.Length > 0 Then
                        If .Tags.Item("DBURL").Length > 0 Then
                            .Tags.Delete("DBURL")
                        End If
                        .Tags.Add("DBURL", awinSettings.databaseURL)
                    End If

                    If awinSettings.databaseName.Length > 0 Then
                        If .Tags.Item("DBNAME").Length > 0 Then
                            .Tags.Delete("DBNAME")
                        End If
                        .Tags.Add("DBNAME", awinSettings.databaseName)
                    End If

                    If Not IsNothing(awinSettings.proxyURL) Then
                        If awinSettings.proxyURL.Length > 0 Then
                            If .Tags.Item("PRXYC").Length > 0 Then
                                .Tags.Delete("PRXYC")
                            End If

                            .Tags.Add("PRXYC", awinSettings.proxyURL)

                            If .Tags.Item("PRXYL").Length > 0 Then
                                .Tags.Delete("PRXYL")
                            End If
                            .Tags.Add("PRXYL", awinSettings.proxyURL)
                        End If
                    End If


                    If .Tags.Item("DBSSL").Length > 0 Then
                        .Tags.Delete("DBSSL")
                    End If
                    .Tags.Add("DBSSL", awinSettings.DBWithSSL.ToString)

                    Dim enryptedUserRole As String = myCustomUserRole.encrypt
                    If .Tags.Item("CURS").Length > 0 Then
                        .Tags.Delete("CURS")
                    End If
                    .Tags.Add("CURS", enryptedUserRole)

                End If

                If .Tags.Item("REST").Length > 0 Then
                    .Tags.Delete("REST")
                End If
                .Tags.Add("REST", awinSettings.visboServer.ToString)


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
    Public Sub addSmartPPTMsPhInfo(ByRef pptShape As PowerPoint.Shape,
                                   ByVal hproj As clsProjekt,
                                          ByVal fullBreadCrumb As String, ByVal classifiedName As String, ByVal shortName As String, ByVal originalName As String,
                                          ByVal bestShortName As String, ByVal bestLongName As String,
                                          ByVal startDate As Date, ByVal endDate As Date,
                                          ByVal ampelColor As Integer, ByVal ampelErlaeuterung As String,
                                          ByVal lieferumfaenge As String,
                                          ByVal verantwortlich As String,
                                          ByVal percentDone As Double,
                                          ByVal docUrl As String)

        Dim nullDate As Date = Nothing

        If Not IsNothing(pptShape) Then
            With pptShape
                If Not IsNothing(hproj.vpID) Then
                    If .Tags.Item("VPID").Length > 0 Then
                        .Tags.Delete("VPID")
                    End If
                    .Tags.Add("VPID", hproj.vpID)
                End If

                ' die Tag Werte müssen immer !! gelöscht werden; andernfalls behalten die Shapes diese Werte und der Update über die Zeit zeigt falsche Ergebnise !!

                If .Tags.Item("BC").Length > 0 Then
                    .Tags.Delete("BC")
                End If
                If Not IsNothing(fullBreadCrumb) Then
                    .Tags.Add("BC", fullBreadCrumb)
                End If

                If .Tags.Item("CN").Length > 0 Then
                    .Tags.Delete("CN")
                End If
                If Not IsNothing(classifiedName) Then
                    .Tags.Add("CN", classifiedName)
                End If

                If .Tags.Item("SN").Length > 0 Then
                    .Tags.Delete("SN")
                End If
                If Not IsNothing(shortName) Then
                    If shortName <> classifiedName And shortName <> "" Then
                        .Tags.Add("SN", shortName)
                    End If
                End If

                If .Tags.Item("ON").Length > 0 Then
                    .Tags.Delete("ON")
                End If
                If Not IsNothing(originalName) Then
                    If originalName <> classifiedName And originalName <> "" Then
                        .Tags.Add("ON", originalName)
                    End If
                End If

                If .Tags.Item("BSN").Length > 0 Then
                    .Tags.Delete("BSN")
                End If
                If Not IsNothing(bestShortName) Then
                    If bestShortName <> shortName Then
                        .Tags.Add("BSN", bestShortName)
                    End If
                End If

                If .Tags.Item("BLN").Length > 0 Then
                    .Tags.Delete("BLN")
                End If
                If Not IsNothing(bestLongName) Then
                    If bestLongName <> classifiedName Then
                        .Tags.Add("BLN", bestLongName)
                    End If
                End If

                If .Tags.Item("SD").Length > 0 Then
                    .Tags.Delete("SD")
                End If
                If Not IsNothing(startDate) Then
                    If Not startDate = nullDate Then
                        .Tags.Add("SD", startDate.ToShortDateString)
                    End If
                End If

                If .Tags.Item("ED").Length > 0 Then
                    .Tags.Delete("ED")
                End If
                If Not IsNothing(endDate) Then
                    If Not endDate = nullDate Then
                        .Tags.Add("ED", endDate.ToShortDateString)
                    End If

                End If

                If .Tags.Item("AC").Length > 0 Then
                    .Tags.Delete("AC")
                End If
                If .Tags.Item("AE").Length > 0 Then
                    .Tags.Delete("AE")
                End If

                If Not IsNothing(ampelColor) Then
                    If ampelColor >= 0 And ampelColor <= 3 Then
                        .Tags.Add("AC", ampelColor.ToString)
                    Else
                        .Tags.Add("AC", "0")
                    End If
                End If

                If Not IsNothing(ampelErlaeuterung) Then
                    If ampelErlaeuterung.Length > 0 Then
                        .Tags.Add("AE", ampelErlaeuterung)
                    End If
                End If

                If .Tags.Item("LU").Length > 0 Then
                    .Tags.Delete("LU")
                End If
                If Not IsNothing(lieferumfaenge) Then
                    If lieferumfaenge.Length > 0 Then
                        .Tags.Add("LU", lieferumfaenge)
                    End If

                End If

                If .Tags.Item("VE").Length > 0 Then
                    .Tags.Delete("VE")
                End If
                If Not IsNothing(verantwortlich) Then
                    If verantwortlich.Trim.Length > 0 Then
                        .Tags.Add("VE", verantwortlich.Trim)
                    End If

                End If

                If .Tags.Item("PD").Length > 0 Then
                    .Tags.Delete("PD")
                End If
                If Not IsNothing(percentDone) Then
                    If percentDone > 0 Then
                        Dim tmpValue As Double = 100 * percentDone
                        .Tags.Add("PD", tmpValue.ToString("0#."))
                    End If

                End If

                ' central document link ..
                If .Tags.Item("DUC").Length > 0 Then
                    .Tags.Delete("DUC")
                End If
                If Not IsNothing(docUrl) Then
                    If docUrl.Length > 0 Then
                        .Tags.Add("DUC", docUrl)
                    End If

                End If

            End With
        End If



    End Sub

    ''' <summary>
    ''' das Shape wurde als Projekt-Karte identifiziert - jetzt werden an das Shape die Projekt-Karten Infos angeheftet ...
    ''' </summary>
    ''' <param name="pptShape"></param>
    ''' <param name="hproj"></param>
    Public Sub addSmartPPTprCardShapeInfo(ByRef pptShape As PowerPoint.Shape,
                                          ByVal hproj As clsProjekt,
                                          ByVal relevantPhase As String,
                                          ByVal isTopProject As Boolean)

        Dim nullDate As Date = Nothing
        Dim bigtype As Integer = ptReportBigTypes.planelements
        Dim detailID As Integer = ptReportComponents.prCard
        Dim tmpStr As String = ""
        Dim kennung As String = ""
        Dim cphase As clsPhase = Nothing

        If relevantPhase <> "" Then
            Dim elemName As String = ""
            Dim breadCrumb As String = ""
            Dim type As Integer = -1
            Dim pvname As String = ""
            Call splitHryFullnameTo2(relevantPhase, elemName, breadCrumb, type, pvname)
            cphase = hproj.getPhase(elemName, breadCrumb)
        End If


        If Not IsNothing(pptShape) Then
            With pptShape

                ' hier kommt der Projekt-Name rein, ggf muss der Phasen-Name berücksichtigt werden 
                If Not IsNothing(cphase) Then
                    tmpStr = hproj.getShapeText & vbLf & " - " & cphase.name
                Else
                    tmpStr = hproj.getShapeText
                End If

                kennung = "CN"
                If .Tags.Item(kennung).Length > 0 Then
                    .Tags.Delete(kennung)
                End If
                .Tags.Add(kennung, tmpStr)

                ' hier kommt nochmal der Projekt-Name rein, das wird in smartInfo ausgewertet ..
                tmpStr = hproj.name
                kennung = "PNM"
                If .Tags.Item(kennung).Length > 0 Then
                    .Tags.Delete(kennung)
                End If
                .Tags.Add(kennung, tmpStr)

                ' hier kommt nochmal der  Projekt-Name rein, das wird in smartInfo ausgewertet ..
                tmpStr = hproj.variantName
                kennung = "VNM"
                If .Tags.Item(kennung).Length > 0 Then
                    .Tags.Delete(kennung)
                End If
                .Tags.Add(kennung, tmpStr)

                ' hier kommt nochmal der vpid rein, das wird in smartInfo ausgewertet ..
                tmpStr = hproj.vpID
                kennung = "VPID"
                If .Tags.Item(kennung).Length > 0 Then
                    .Tags.Delete(kennung)
                End If
                .Tags.Add(kennung, tmpStr)


                ' jetzt das Startdatum des Projektes bzw. der Phase
                If Not IsNothing(cphase) Then
                    tmpStr = cphase.getStartDate.ToShortDateString
                Else
                    tmpStr = hproj.startDate.ToShortDateString
                End If
                kennung = "SD"
                If .Tags.Item(kennung).Length > 0 Then
                    .Tags.Delete(kennung)
                End If
                .Tags.Add(kennung, tmpStr)

                ' jetzt das Ende-Datum des Projekts
                If Not IsNothing(cphase) Then
                    tmpStr = cphase.getEndDate.ToShortDateString
                Else
                    tmpStr = hproj.endeDate.ToShortDateString
                End If
                kennung = "ED"
                If .Tags.Item(kennung).Length > 0 Then
                    .Tags.Delete(kennung)
                End If
                .Tags.Add(kennung, tmpStr)

                ' jetzt die Ampel des Projektes
                tmpStr = hproj.ampelStatus.ToString
                kennung = "AC"
                If .Tags.Item(kennung).Length > 0 Then
                    .Tags.Delete(kennung)
                End If
                .Tags.Add(kennung, tmpStr)

                ' jetzt die Ampel-Erläuterung des Projekts 
                tmpStr = hproj.ampelErlaeuterung
                kennung = "AE"
                If .Tags.Item(kennung).Length > 0 Then
                    .Tags.Delete(kennung)
                End If
                .Tags.Add(kennung, tmpStr)

                ' jetzt die Ziele des Projektes bzw. der Phase
                If Not IsNothing(cphase) Then
                    Try
                        If cphase.countMilestones > 0 Then
                            Dim milestone As clsMeilenstein = cphase.getMilestone(cphase.countMilestones)
                            If Not IsNothing(milestone) Then
                                tmpStr = milestone.getAllDeliverables
                            End If
                        End If
                    Catch ex As Exception
                        tmpStr = "-"
                    End Try

                Else
                    tmpStr = hproj.fullDescription
                End If

                kennung = "LU"
                If .Tags.Item(kennung).Length > 0 Then
                    .Tags.Delete(kennung)
                End If
                .Tags.Add(kennung, tmpStr)

                'jetzt die Risiken des Projektes 
                tmpStr = ""
                kennung = "RSK"
                If .Tags.Item(kennung).Length > 0 Then
                    .Tags.Delete(kennung)
                End If
                .Tags.Add(kennung, tmpStr)

                ' jetzt Verantwortlicher des Projektes 
                tmpStr = hproj.leadPerson
                kennung = "VE"
                If .Tags.Item(kennung).Length > 0 Then
                    .Tags.Delete(kennung)
                End If
                .Tags.Add(kennung, tmpStr)

                ' jetzt PercentDone des Projektes 
                ' aktuell noch leer lassen 

                ' jetzt die BigType ID 
                tmpStr = CStr(ptReportBigTypes.planelements)
                kennung = "BID"
                If .Tags.Item(kennung).Length > 0 Then
                    .Tags.Delete(kennung)
                End If
                .Tags.Add(kennung, tmpStr)

                ' jetzt die Detail ID 
                If isTopProject Then
                    tmpStr = CStr(ptReportComponents.prCard)
                Else
                    tmpStr = CStr(ptReportComponents.prCardinvisible)
                End If

                kennung = "DID"
                If .Tags.Item(kennung).Length > 0 Then
                    .Tags.Delete(kennung)
                End If
                .Tags.Add(kennung, tmpStr)

            End With
        End If



    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="pptShape">repräsentiert das Powerpoint Shape, wird per Ref übergeben </param>
    ''' <param name="ampelfarbID">gibt die Kennung für die Ampel an: 
    ''' 0 - none
    ''' 1 - green
    ''' 2 - yellow 
    ''' 3 - red</param>
    ''' <remarks></remarks>
    Public Sub switchOnTrafficLightColor(ByRef pptShape As PowerPoint.Shape, ByVal ampelfarbID As Integer)

        Dim redAmpel As PowerPoint.Shape = Nothing
        Dim yellowAmpel As PowerPoint.Shape = Nothing
        Dim greenAmpel As PowerPoint.Shape = Nothing
        Dim ampelGeruest As PowerPoint.Shape = Nothing

        Try
            If pptShape.GroupItems.Count > 1 Then

                Dim trafficLightShape As PowerPoint.ShapeRange = pptShape.Ungroup

                For Each tmpShape As PowerPoint.Shape In trafficLightShape
                    ' Licht auf Off setzen 

                    If tmpShape.Title = "VisboAmpelRed" Then
                        redAmpel = tmpShape
                        redAmpel.Fill.Transparency = 0.95
                    ElseIf tmpShape.Title = "VisboAmpelYellow" Then
                        yellowAmpel = tmpShape
                        yellowAmpel.Fill.Transparency = 0.95
                    ElseIf tmpShape.Title = "VisboAmpelGreen" Then
                        greenAmpel = tmpShape
                        greenAmpel.Fill.Transparency = 0.95
                    Else
                        ampelGeruest = tmpShape
                    End If
                Next


                ' jetzt die richtige Ampel setzen
                If ampelfarbID = 1 Then

                    If Not IsNothing(greenAmpel) Then
                        greenAmpel.Fill.Transparency = 0.0
                    End If

                ElseIf ampelfarbID = 2 Then

                    If Not IsNothing(yellowAmpel) Then
                        yellowAmpel.Fill.Transparency = 0.0
                    End If

                ElseIf ampelfarbID = 3 Then

                    If Not IsNothing(redAmpel) Then
                        redAmpel.Fill.Transparency = 0.0
                    End If

                End If

                Try
                    If Not IsNothing(ampelGeruest) Then
                        ' nach vorne holen 
                        ampelGeruest.ZOrder(MsoZOrderCmd.msoBringToFront)
                    End If
                Catch ex As Exception

                End Try


                ' jetzt wieder gruppieren 
                pptShape = trafficLightShape.Group
                pptShape.Title = "SymTrafficLight"

            End If
        Catch ex As Exception
            ' wenn es gar kein zusammengesetztes Shape ist ... 
        End Try


    End Sub

    ''' <summary>
    ''' ergänzt Chart Tags für ein Projekt wie Portfolio Chart 
    ''' </summary>
    ''' <param name="pptShape"></param>
    ''' <param name="scInfo"></param>
    Public Sub addSmartPPTChartInfo(ByRef pptShape As PowerPoint.Shape, ByVal scinfo As clsSmartPPTChartInfo)

        If scinfo.prPF = ptPRPFType.portfolio Then
            ' alles klar ... 
        Else
            If IsNothing(scinfo.hproj) Then
                Exit Sub
            End If
        End If

        ' tk 23.4.19 hier wird 
        'Dim pName As String = scinfo.hproj.name
        'Dim vName As String = scinfo.hproj.variantName
        Dim pName As String = scinfo.pName
        Dim vName As String = scinfo.vName
        Dim vpid As String = scinfo.vpid

        Dim chtObjName As String = ""

        'Dim encryptedUR As String = encryptmyCustomUserRole
        Try

            If Not IsNothing(pptShape) Then

                If pptShape.HasChart = MsoTriState.msoTrue Then
                    Dim pptChart As PowerPoint.Chart = pptShape.Chart
                    chtObjName = pptChart.Name

                    With pptShape

                        If .Tags.Item("CHON").Length > 0 Then
                            .Tags.Delete("CHON")
                        End If
                        If Not IsNothing(chtObjName) Then
                            .Tags.Add("CHON", chtObjName)
                        End If

                        If .Tags.Item("CHT").Length > 0 Then
                            .Tags.Delete("CHT")
                        End If
                        If Not IsNothing(scinfo.chartTyp) Then
                            .Tags.Add("CHT", CStr(CInt(scinfo.chartTyp)))
                        End If

                        If .Tags.Item("ASW").Length > 0 Then
                            .Tags.Delete("ASW")
                        End If
                        If Not IsNothing(scinfo.einheit) Then
                            .Tags.Add("ASW", CStr(CInt(scinfo.einheit)))
                        End If

                        If .Tags.Item("VGLA").Length > 0 Then
                            .Tags.Delete("VGLA")
                        End If
                        .Tags.Add("VGLA", CStr(CInt(scinfo.vergleichsArt)))

                        If .Tags.Item("VGLT").Length > 0 Then
                            .Tags.Delete("VGLT")
                        End If
                        .Tags.Add("VGLT", CStr(CInt(scinfo.vergleichsTyp)))

                        If .Tags.Item("VGLD").Length > 0 Then
                            .Tags.Delete("VGLD")
                        End If
                        .Tags.Add("VGLD", scinfo.vergleichsDatum.ToString)


                        If .Tags.Item("PRPF").Length > 0 Then
                            .Tags.Delete("PRPF")
                        End If
                        '.Tags.Add("PRPF", CStr(scinfo.hproj.projectType))
                        .Tags.Add("PRPF", CStr(scinfo.prPF))

                        If .Tags.Item("PNM").Length > 0 Then
                            .Tags.Delete("PNM")
                        End If
                        If Not IsNothing(pName) Then
                            .Tags.Add("PNM", pName)
                        End If

                        If .Tags.Item("VNM").Length > 0 Then
                            .Tags.Delete("VNM")
                        End If
                        If Not IsNothing(vName) Then
                            .Tags.Add("VNM", vName)
                        End If

                        If .Tags.Item("VPID").Length > 0 Then
                            .Tags.Delete("VPID")
                        End If
                        If Not IsNothing(vpid) Then
                            .Tags.Add("VPID", vpid)
                        End If


                        If .Tags.Item("Q1").Length > 0 Then
                            .Tags.Delete("Q1")
                        End If
                        .Tags.Add("Q1", CStr(CInt(scinfo.elementTyp)))


                        If .Tags.Item("Q2").Length > 0 Then
                            .Tags.Delete("Q2")
                        End If
                        .Tags.Add("Q2", scinfo.q2)


                        If .Tags.Item("SRLD").Length > 0 Then
                            .Tags.Delete("SRLD")
                        End If


                        If .Tags.Item("SRRD").Length > 0 Then
                            .Tags.Delete("SRRD")
                        End If

                        If scinfo.zeitRaumLeft > Date.MinValue Then
                            .Tags.Add("SRLD", CStr(scinfo.zeitRaumLeft))
                            .Tags.Add("SRRD", CStr(scinfo.zeitRaumRight))
                        End If

                        If .Tags.Item("BID").Length > 0 Then
                            .Tags.Delete("BID")
                        End If
                        .Tags.Add("BID", CStr(CInt(scinfo.bigType)))

                        If .Tags.Item("DID").Length > 0 Then
                            .Tags.Delete("DID")
                        End If
                        .Tags.Add("DID", CStr(CInt(scinfo.detailID)))


                    End With

                End If

            End If

        Catch ex As Exception
            Dim a As Integer = 1
        End Try


    End Sub


    ''' <summary>
    ''' fügt für  Reporting Komponenten die entsprechenden Smart-Infos hinzu, so dass 
    ''' der Powerpoint Add-In die Komponente selbstständig aktualisieren kann 
    ''' </summary>
    ''' <param name="pptShape"></param>
    ''' <param name="prpf"></param>
    ''' <param name="qualifier"></param>
    ''' <remarks></remarks>
    Public Sub addSmartPPTCompInfo(ByRef pptShape As PowerPoint.Shape, ByVal hproj As clsProjekt, ByVal hportfolio As clsConstellation,
                                     ByVal prpf As Integer, ByVal qualifier As String, ByVal qualifier2 As String,
                                     ByVal bigType As Integer, ByVal detailID As Integer)
        Try
            Dim pName As String = ""
            Dim vName As String = ""
            Dim vpid As String = ""

            If prpf = ptPRPFType.portfolio And Not IsNothing(hportfolio) Then
                ' hier handelt es sich um ein Portfolio
                pName = hportfolio.constellationName
                vName = ""
                vpid = hportfolio.vpID

            Else
                ' hier handelt es sich um ein Projekt
                If prpf = ptPRPFType.project And Not IsNothing(hproj) Then

                    pName = hproj.name
                    vName = hproj.variantName
                    vpid = hproj.vpID

                    ' jetzt kommen noch die Ergänzungen, die je nach Typ notwendig sind ...
                    If bigType = ptReportBigTypes.tables Then
                        ' sonst keine weiteren Dinge ... das wird in der eigenen Methode addSmartPPTTableInfo gemacht 

                    ElseIf bigType = ptReportBigTypes.components Then
                        ' bei Symbolen muss noch was ergänzt werden 

                        If detailID = ptReportComponents.prSymTrafficLight Or
                            detailID = ptReportComponents.prSymRisks Or
                            detailID = ptReportComponents.prSymDescription Or
                            detailID = ptReportComponents.prSymFinance Or
                            detailID = ptReportComponents.prSymProject Or
                            detailID = ptReportComponents.prSymSchedules Or
                            detailID = ptReportComponents.prSymTeam Then

                            Call updateSmartPPTSymTxt(pptShape, hproj, detailID)

                        End If


                        ' sonst keine weiteren Dinge 

                    Else
                        ' noch nicht implementiert 
                    End If

                Else

                    Exit Sub

                End If

            End If


            If Not IsNothing(pptShape) Then

                ' das bekommen alle ...
                With pptShape

                    If .Tags.Item("PRPF").Length > 0 Then
                        .Tags.Delete("PRPF")
                    End If
                    If Not IsNothing(prpf) Then
                        .Tags.Add("PRPF", prpf.ToString)
                    End If

                    If .Tags.Item("PNM").Length > 0 Then
                        .Tags.Delete("PNM")
                    End If
                    If Not IsNothing(pName) Then
                        .Tags.Add("PNM", pName)
                    End If

                    If .Tags.Item("VNM").Length > 0 Then
                        .Tags.Delete("VNM")
                    End If
                    If Not IsNothing(vName) Then
                        .Tags.Add("VNM", vName)
                    End If

                    If .Tags.Item("VPID").Length > 0 Then
                        .Tags.Delete("VPID")
                    End If
                    If Not IsNothing(vpid) Then
                        .Tags.Add("VPID", vpid)
                    End If

                    If .Tags.Item("Q1").Length > 0 Then
                        .Tags.Delete("Q1")
                    End If
                    If Not IsNothing(qualifier) Then
                        .Tags.Add("Q1", qualifier)
                    End If

                    If .Tags.Item("Q2").Length > 0 Then
                        .Tags.Delete("Q2")
                    End If
                    If Not IsNothing(qualifier2) Then
                        .Tags.Add("Q2", qualifier2)
                    End If

                    If .Tags.Item("SRLD").Length > 0 Then
                        .Tags.Delete("SRLD")
                    End If

                    If .Tags.Item("SRRD").Length > 0 Then
                        .Tags.Delete("SRRD")
                    End If

                    If showRangeLeft >= 0 Then
                        .Tags.Add("SRLD", CStr(getDateofColumn(showRangeLeft, False)))
                        .Tags.Add("SRRD", CStr(getDateofColumn(showRangeRight, False)))
                    End If

                    If .Tags.Item("BID").Length > 0 Then
                        .Tags.Delete("BID")
                    End If
                    If Not IsNothing(bigType) Then
                        .Tags.Add("BID", bigType.ToString)
                    End If

                    If .Tags.Item("DID").Length > 0 Then
                        .Tags.Delete("DID")
                    End If
                    If Not IsNothing(detailID) Then
                        .Tags.Add("DID", detailID.ToString)
                    End If


                End With

            End If


        Catch ex As Exception
            Dim a As Integer = 1
        End Try

    End Sub

    ''' <summary>
    ''' aktualisiert bei Symbolen den Tag TXT entsprechend der übergebenen detailID 
    ''' </summary>
    ''' <param name="pptShape"></param>
    ''' <param name="hproj"></param>
    ''' <param name="detailID"></param>
    ''' <remarks></remarks>
    Public Sub updateSmartPPTSymTxt(ByRef pptShape As PowerPoint.Shape, ByVal hproj As clsProjekt, ByVal detailID As Integer)
        Dim tmpText As String = ""

        If detailID = ptReportComponents.prSymTrafficLight Then
            tmpText = "Version: " & hproj.timeStamp.ToShortDateString & vbLf & vbLf
            tmpText = tmpText & hproj.ampelErlaeuterung

        ElseIf detailID = ptReportComponents.prSymRisks Then
            ' aktuell gibt es im Datenmodell noch keine Risiken
            tmpText = "Version: " & hproj.timeStamp.ToShortDateString & vbLf & vbLf

            If Not IsNothing(hproj.getCustomSField("Risiko")) Then
                tmpText = tmpText & hproj.getCustomSField("Risiko")
            Else
                tmpText = tmpText & "--"
            End If

        ElseIf detailID = ptReportComponents.prSymDescription Then

            tmpText = "Version: " & hproj.timeStamp.ToShortDateString & vbLf & vbLf
            tmpText = tmpText & hproj.fullDescription

        ElseIf detailID = ptReportComponents.prSymFinance Then
            ' es werden Budget, Personalkosten, sosnt. Kosten und Forecast Ergebnis angezeigt
            Dim budget As Double, pk As Double, sk As Double, rk As Double, forecast As Double
            Call hproj.calculateRoundedKPI(budget, pk, sk, rk, forecast)

            tmpText = "Version: " & hproj.timeStamp.ToShortDateString & vbLf & vbLf

            tmpText = tmpText & "Budget: " & budget.ToString & " T€" & vbLf
            tmpText = tmpText & "Personnel-Cost: " & pk.ToString & " T€" & vbLf
            tmpText = tmpText & "Other Costs: " & sk.ToString & " T€" & vbLf
            If forecast < 0 Then
                tmpText = tmpText & "Loss: " & forecast.ToString & " T€"
            Else
                tmpText = tmpText & "Profit: " & forecast.ToString & " T€"
            End If


        ElseIf detailID = ptReportComponents.prSymSchedules Then

            tmpText = ""
            Try

                ' es werden die Gesamt-Anzahl Überfälliger Meilensteine / Vorgänge angezeigt 
                ' es werden die Gesamt-Anzahl roter, gelber, grüner, Nicht-bewerteter Meilensteine / Vorgänge angezeigt 
                Dim sortedListOfMilestones As SortedList(Of Date, String) = hproj.getMilestones
                Dim sortedListOfTasks As SortedList(Of Date, String) = hproj.getPhases

                tmpText = "Version: " & hproj.timeStamp.ToShortDateString & vbLf & vbLf

                Dim vglDatum As Date = hproj.timeStamp
                Dim anz As Integer = 0
                Dim overDue As Integer = 0

                ' jetzt die Meilensteine überprüfen 
                If sortedListOfMilestones.Count > 0 Then
                    Dim ix As Integer = 0
                    Dim curDate As Date = sortedListOfMilestones.ElementAt(ix).Key

                    Do While DateDiff(DateInterval.Day, curDate, vglDatum) > 0 And ix <= sortedListOfMilestones.Count - 1

                        If hproj.getMilestoneByID(sortedListOfMilestones.ElementAt(ix).Value).percentDone < 1 Then
                            overDue = overDue + 1
                        End If

                        anz = anz + 1
                        ix = ix + 1

                        curDate = sortedListOfMilestones.ElementAt(ix).Key
                    Loop


                End If

                ' jetzt die Phasen überprüfen 
                If sortedListOfTasks.Count > 0 Then
                    Dim ix As Integer = 0
                    Dim curDate As Date = sortedListOfTasks.ElementAt(ix).Key

                    Do While DateDiff(DateInterval.Day, curDate, vglDatum) > 0 And ix <= sortedListOfTasks.Count - 1

                        If hproj.getPhaseByID(sortedListOfTasks.ElementAt(ix).Value).percentDone < 1 Then
                            overDue = overDue + 1
                        End If

                        anz = anz + 1
                        ix = ix + 1

                        curDate = sortedListOfMilestones.ElementAt(ix).Key
                    Loop

                End If

                tmpText = tmpText & "Number overdue / total number until version-date: " & overDue.ToString & " / " & anz.ToString & vbLf & vbLf

                ' jetzt werden die Anzahl roten, gelben, grünen, grauen Bewertungen gezählt ..
                Dim anzRed As Integer = 0, anzYellow As Integer = 0, anzGreen As Integer = 0, anzNoColor As Integer = 0

                anz = sortedListOfMilestones.Count + sortedListOfTasks.Count

                If sortedListOfMilestones.Count > 0 Then
                    For Each kvp As KeyValuePair(Of Date, String) In sortedListOfMilestones
                        Dim tmpColor As Integer = hproj.getMilestoneByID(kvp.Value).ampelStatus
                        If tmpColor = 0 Then
                            anzNoColor = anzNoColor + 1
                        ElseIf tmpColor = 1 Then
                            anzGreen = anzGreen + 1
                        ElseIf tmpColor = 2 Then
                            anzYellow = anzYellow + 1
                        ElseIf tmpColor = 3 Then
                            anzRed = anzRed + 1
                        Else
                            anzNoColor = anzNoColor + 1
                        End If
                    Next

                End If

                If sortedListOfTasks.Count > 0 Then
                    For Each kvp As KeyValuePair(Of Date, String) In sortedListOfTasks
                        Dim tmpColor As Integer = hproj.getPhaseByID(kvp.Value).ampelStatus
                        If tmpColor = 0 Then
                            anzNoColor = anzNoColor + 1
                        ElseIf tmpColor = 1 Then
                            anzGreen = anzGreen + 1
                        ElseIf tmpColor = 2 Then
                            anzYellow = anzYellow + 1
                        ElseIf tmpColor = 3 Then
                            anzRed = anzRed + 1
                        Else
                            anzNoColor = anzNoColor + 1
                        End If
                    Next

                End If

                tmpText = tmpText & "Total number of Milestones/Tasks: " & anz.ToString & vbLf
                tmpText = tmpText & "No rating: " & anzNoColor.ToString & vbLf
                tmpText = tmpText & "Green: " & anzGreen.ToString & vbLf
                tmpText = tmpText & "Yellow: " & anzYellow.ToString & vbLf
                tmpText = tmpText & "Red: " & anzRed.ToString & vbLf



            Catch ex As Exception

            End Try


        ElseIf detailID = ptReportComponents.prSymProject Then
            ' es werden Informationen zum Projekt angezeigt 
            ' eigentlich wäre es hier am besten ein 

            tmpText = "Version: " & hproj.timeStamp.ToShortDateString & vbLf & vbLf
            tmpText = ""

        ElseIf detailID = ptReportComponents.prSymTeam Then
            ' es wird das Team angezeigt ...

            tmpText = "Version: " & hproj.timeStamp.ToShortDateString & vbLf & vbLf

            Dim allNames As Collection = hproj.getRoleNames

            Dim responsible As String = hproj.leadPerson
            tmpText = tmpText & "Project-Lead: " & responsible & vbLf & vbLf

            Dim allResponsible As Collection = hproj.getResponsibleNames
            If allResponsible.Count > 0 Then
                tmpText = tmpText & "Team:" & vbLf
                For Each tmpName As String In allResponsible
                    tmpText = tmpText & tmpName & vbLf
                Next
            End If

            If allResponsible.Count > 0 Then
                tmpText = tmpText & vbLf
            End If

            For Each tmpName As String In allNames
                tmpText = tmpText & tmpName & vbLf
            Next

        End If

        ' jetzt wird das unter dem Tag TXT eingetragen
        With pptShape
            If .Tags.Item("TXT").Length > 0 Then
                .Tags.Delete("TXT")
            End If
            .Tags.Add("TXT", tmpText)
        End With

    End Sub

    ''' <summary>
    ''' fügt der Tabelle die Smart Table Info hinzu 
    ''' </summary>
    ''' <param name="pptShape"></param>
    ''' <param name="prpf"></param>
    ''' <param name="pnm"></param>
    ''' <param name="vnm"></param>
    ''' <param name="q1"></param>
    ''' <param name="q2"></param>
    ''' <param name="bigtype"></param>
    ''' <param name="detailID"></param>
    ''' <param name="nameIDS"></param>
    ''' <remarks></remarks>
    Public Sub addSmartPPTTableInfo(ByRef pptShape As PowerPoint.Shape,
                                        ByVal prpf As Integer, ByVal pnm As String, ByVal vnm As String, ByVal vpid As String,
                                        ByVal q1 As String, ByVal q2 As String,
                                        ByVal bigtype As Integer, ByVal detailID As Integer,
                                        ByVal nameIDS As Collection)

        If nameIDS.Count = 0 And bigtype = ptReportBigTypes.charts Then
            Exit Sub
        End If

        Dim nameIDString As String = ""
        nameIDString = convertCollToNids(nameIDS)

        Try

            If Not IsNothing(pptShape) Then

                ' das bekommen alle ...
                With pptShape
                    If .Tags.Item("PRPF").Length > 0 Then
                        .Tags.Delete("PRPF")
                    End If
                    If Not IsNothing(prpf) Then
                        .Tags.Add("PRPF", prpf.ToString)
                    End If

                    If .Tags.Item("PNM").Length > 0 Then
                        .Tags.Delete("PNM")
                    End If
                    If Not IsNothing(pnm) Then
                        .Tags.Add("PNM", pnm)
                    End If

                    If .Tags.Item("VNM").Length > 0 Then
                        .Tags.Delete("VNM")
                    End If
                    If Not IsNothing(vnm) Then
                        .Tags.Add("VNM", vnm)
                    End If

                    If .Tags.Item("VPID").Length > 0 Then
                        .Tags.Delete("VPID")
                    End If
                    If Not IsNothing(vpid) Then
                        .Tags.Add("VPID", vpid)
                    End If

                    If .Tags.Item("Q1").Length > 0 Then
                        .Tags.Delete("Q1")
                    End If
                    If Not IsNothing(q1) Then
                        .Tags.Add("Q1", q1)
                    End If


                    If .Tags.Item("Q2").Length > 0 Then
                        .Tags.Delete("Q2")
                    End If
                    If Not IsNothing(q2) Then
                        .Tags.Add("Q2", q2)
                    End If

                    If .Tags.Item("BID").Length > 0 Then
                        .Tags.Delete("BID")
                    End If
                    If Not IsNothing(bigtype) Then
                        .Tags.Add("BID", bigtype.ToString)
                    End If

                    If .Tags.Item("DID").Length > 0 Then
                        .Tags.Delete("DID")
                    End If
                    If Not IsNothing(detailID) Then
                        .Tags.Add("DID", detailID.ToString)
                    End If

                    If .Tags.Item("NIDS").Length > 0 Then
                        .Tags.Delete("NIDS")
                    End If
                    If Not IsNothing(nameIDString) Then
                        .Tags.Add("NIDS", nameIDString)
                    End If

                End With

            End If

        Catch ex As Exception
            Dim a As Integer = 1
        End Try

    End Sub

    ''' <summary>
    ''' zeichnet bzw. aktualisiert die Powerpoint Table Milestone-Übersicht 
    ''' wenn todoCollection leer, dann wird die Perfromance Metrik gezeichnet 
    ''' </summary>
    ''' <param name="pptShape"></param>
    ''' <param name="hproj"></param>
    ''' <param name="bproj"></param>
    ''' <param name="lproj"></param>
    ''' <param name="toDoCollection">enthält die NAmen, ggf incl P:, V: oder C: Qualifier und ggf inkl Breadcrumb Anteilen </param>
    ''' <param name="q1">0, später die Anzahl Phasen </param>
    ''' <param name="q2">Anzahl Milestones; aktuell redundant, da identisch mit Anzahl in der Collection</param>
    Public Sub zeichneTableMilestoneAPVCV(ByRef pptShape As pptNS.Shape, ByVal hproj As clsProjekt, ByVal bproj As clsProjekt, ByVal lproj As clsProjekt,
                                     ByVal toDoCollection As Collection, ByVal q1 As String, ByVal q2 As String)

        Dim repmsg() As String
        ' Performance Ratio 1 ist das Verhältnis zwischen der Anzahl aktuell erreichter Meilensteine im betrachteten Monat versus der Anzahl erreichter Meilensteine im betrachteten Monat im stand zur Beauftragung 
        ' Performance Ratio 2 ist das Verhältnis zwischen der Anzahl aktuell erreichter Meilensteine im betrachteten Monat versus der Anzahl erreichter Meilensteine im betrachteten Monat im Stand der letzten Planung
        repmsg = {"total sum of milestones", "finished until last month", "due this month", "total sum of finished milestones up-to-date", "due next month", "Overdue", "Performance Ratio (first)", "Performance Ratio (last)"}

        Dim txtPKI(7) As String

        txtPKI(0) = repmsg(0) ' total sum of milestones
        txtPKI(1) = repmsg(1) ' finished until last month
        txtPKI(2) = repmsg(2) ' due this month
        txtPKI(3) = repmsg(3) ' total sum of finished milestones up-to-date
        txtPKI(4) = repmsg(4) ' due next month
        txtPKI(5) = repmsg(5) ' Overdue
        txtPKI(6) = repmsg(6) ' Performance Ratio (first)
        txtPKI(7) = repmsg(7) ' Performance Ratio (last)


        ' steuert Einrückung ja, nein im Not Overview Modus 
        Dim einrueckung As Integer = 0

        Dim tabelle As pptNS.Table
        Dim anzSpalten As Integer


        Dim bigType As Integer = ptReportBigTypes.tables
        Dim compID As Integer = PTpptTableTypes.prMilestoneAPVCV

        ' wenn die gleich sind, sollen keine zwei Spalten mit identischen Werten ausgewiesen werden 
        If Not IsNothing(lproj) And Not IsNothing(bproj) Then
            If lproj.timeStamp = bproj.timeStamp Then
                lproj = Nothing
            End If
        End If

        Dim considerFapr As Boolean = Not IsNothing(bproj)
        Dim considerLapr As Boolean = Not IsNothing(lproj)

        Dim anzMilestones As Integer = 0
        Dim anzPhases As Integer = 0

        ' in q1, q2 sind dei Anzahl Rollen bzw Kosten drin, sofern in toDoCollection was angegeben ist 
        Try
            anzPhases = CInt(q1)
            anzMilestones = CInt(q2)
        Catch ex As Exception

        End Try

        Dim showOverviewOnly As Boolean = (toDoCollection.Count = 0)

        ' jetzt wird SmartTableInfo gesetzt 
        ' jetzt wird die SmartTableInfo gesetzt 
        Call addSmartPPTTableInfo(pptShape,
                                  hproj.projectType, hproj.name, hproj.variantName, hproj.vpID,
                                  q1, q2, bigType, compID,
                                  toDoCollection)

        ' jetzt werden die einzelnen Zeilen geschrieben 

        Try
            tabelle = pptShape.Table
            anzSpalten = tabelle.Columns.Count
            If anzSpalten = 6 Then
                ' dann ist alles in Ordnung .. 


                ' jetzt überprüfen, ob die Tabelle aktuell nur aus 2 Zeilen besteht ...
                If tabelle.Rows.Count > 2 Then
                    Do While tabelle.Rows.Count > 2
                        tabelle.Rows(2).Delete()
                    Loop
                End If

                ' jetzt die Werte in den 6 Spalten zurücksetzen 
                Try
                    With tabelle
                        For i As Integer = 1 To 6
                            .Cell(2, i).Shape.TextFrame2.TextRange.Text = ""
                        Next
                    End With
                Catch ex As Exception

                End Try


                Dim faprDate As Date = Date.MinValue
                Dim laprDate As Date = Date.MinValue
                Dim curDate As Date = Date.MinValue

                If Not IsNothing(hproj) Then
                    curDate = hproj.timeStamp
                End If
                If Not IsNothing(bproj) Then
                    faprDate = bproj.timeStamp
                End If
                If Not IsNothing(lproj) Then
                    laprDate = lproj.timeStamp
                End If

                ' jetzt die Headerzeile schreiben 
                Call schreibeAPVCVHeaderZeile(tabelle, faprDate, laprDate, curDate, considerFapr, considerLapr)

                Dim tabellenzeile As Integer = 2
                Try

                    If Not showOverviewOnly Then

                        einrueckung = 1

                        ' dient dazu , zu bestimmen, wann die Kostenarten kommen um vorher eine Neue Zeile  einzufügen ...
                        Dim firstMilestone As Boolean = True

                        Dim curValue As Date = Date.MinValue ' not defined
                        Dim faprValue As Date = Date.MinValue  ' first approved version 
                        Dim laprValue As Date = Date.MinValue  ' last approved version

                        If anzPhases > 0 Then
                            ' 
                            'tabelle.Cell(tabellenzeile, 1).Shape.TextFrame2.TextRange.Text = repMessages.getmsg(51)
                            tabelle.Cell(tabellenzeile, 1).Shape.TextFrame2.TextRange.Text = "Phases"
                            tabelle.Rows.Add()
                            tabellenzeile = tabellenzeile + 1
                        End If

                        ' nimmt die eindeutigen IDs auf 
                        Dim listOfIDs As New Collection

                        For m As Integer = 1 To toDoCollection.Count

                            Dim tmpCollection As New Collection From {
                                CStr(toDoCollection.Item(m))
                            }

                            'Dim hprojBreadcrumbs() As String = hproj.getBreadCrumbArray(Nothing, hproj.getElemIdsOf(tmpCollection, True))
                            'Dim bprojBreadCrumbs() As String = Nothing
                            'Dim lprojBreadCrumbs() As String = Nothing

                            Dim hprojLIDs As Collection = hproj.getElemIdsOf(tmpCollection, True)
                            Dim bProjLIDs As Collection = Nothing
                            Dim lprojLIDs As Collection = Nothing

                            If considerFapr Then
                                bProjLIDs = bproj.getElemIdsOf(tmpCollection, True)
                            End If

                            If considerLapr Then
                                lprojLIDs = lproj.getElemIdsOf(tmpCollection, True)
                            End If

                            ' hproj steuert jetzt die Schleife 
                            For hix As Integer = 1 To hprojLIDs.Count

                                Dim hprojMsID As String = CStr(hprojLIDs.Item(hix))
                                Dim curItem As String = elemNameOfElemID(hprojMsID)

                                curValue = Date.MinValue
                                faprValue = Date.MinValue
                                laprValue = Date.MinValue

                                Dim hMilestone As clsMeilenstein = hproj.getMilestoneByID(hprojMsID)

                                If Not IsNothing(hMilestone) Then
                                    curValue = hMilestone.getDate
                                End If

                                Dim bMileStone As clsMeilenstein = Nothing
                                If considerFapr Then
                                    If hix <= bProjLIDs.Count Then
                                        bMileStone = bproj.getMilestoneByID(CStr(bProjLIDs.Item(hix)))
                                        If Not IsNothing(bMileStone) Then
                                            faprValue = bMileStone.getDate
                                        End If
                                    End If
                                End If

                                Dim lMilestone As clsMeilenstein = Nothing
                                If considerLapr Then
                                    If hix <= lprojLIDs.Count Then
                                        lMilestone = lproj.getMilestoneByID(CStr(lprojLIDs.Item(hix)))
                                        If Not IsNothing(lMilestone) Then
                                            laprValue = lMilestone.getDate
                                        End If
                                    End If
                                End If

                                Call schreibeMilestoneAPVCVZeile(tabelle, tabellenzeile, curItem, faprValue, laprValue, curValue,
                                                          considerFapr, considerLapr)

                                tabelle.Rows.Add()
                                tabellenzeile = tabellenzeile + 1


                            Next

                        Next


                    Else

                        Call MsgBox("noch nicht implementiert ...")

                    End If

                    ' jetzt letzte Zeile löschen  ...
                    tabelle.Rows(tabellenzeile).Delete()
                    tabellenzeile = tabellenzeile - 1

                Catch ex1 As Exception

                End Try
            Else
                Throw New Exception("Tabelle should have 6 columns ... exit ...")
            End If
        Catch ex As Exception

        End Try


    End Sub


    ''' <summary>
    ''' zeichnet bzw. aktualisiert die Powerpoint Table Kosten-Übersicht 
    ''' wenn q1=0 und q2=0 , dann wird die Gesamt-Übersicht Budget, Personal-Kosten, sonstige Kosten, Ergebnis gezeichnet
    ''' wenn q1=-1, q2=-1 , dann wird %used" gezeichnet 
    ''' </summary>
    ''' <param name="pptShape"></param>
    ''' <param name="hproj"></param>
    ''' <param name="bproj"></param>
    ''' <param name="lproj"></param>
    ''' <param name="q1">gibt an, wieviele Rollen, die ersten q1 in der todoCollection sind Rollen</param>
    ''' <param name="q2">gibt an wieviele Kosten</param>
    Public Sub zeichneTableBudgetCostAPVCV(ByRef pptShape As pptNS.Shape, ByVal hproj As clsProjekt, ByVal bproj As clsProjekt, ByVal lproj As clsProjekt,
                                    ByVal q1 As String, ByVal q2 As String)


        Dim repmsg() As String
        Dim toDoCollectionR As New Collection
        Dim toDoCollectionC As New Collection
        Dim showEuro As Boolean = True

        If q1 = "PT" Then
            showEuro = False
        End If

        repmsg = {"Budget", "Personalkosten", "Sonstige Kosten", "Ergebnis-Prognose"}
        'repmsg(1) = {"Budget", "Personnel Costs", "Other Costs", "Profit/Loss"}


        Dim txtPKI(3) As String

        txtPKI(0) = repmsg(0) ' Budget
        txtPKI(1) = repmsg(1) ' Personalkosten
        txtPKI(2) = repmsg(2) ' Sonstige Kosten
        txtPKI(3) = repmsg(3) ' Ergebnis-Prognose

        Dim curValue As Double = -1.0 ' not defined
        Dim faprValue As Double = -1.0 ' first approved version 
        Dim laprValue As Double = -1.0 ' last approved version



        ' steuert Einrückung ja, nein im Not Overview Modus 
        Dim einrueckung As Integer = 0

        Dim tabelle As pptNS.Table
        Dim anzSpalten As Integer


        Dim bigType As Integer = ptReportBigTypes.tables
        Dim compID As Integer = PTpptTableTypes.prBudgetCostAPVCV

        ' wenn die gleich sind, sollen keine zwei Spalten mit identischen Werten ausgewiesen werden 
        If Not IsNothing(lproj) And Not IsNothing(bproj) Then
            If lproj.timeStamp = bproj.timeStamp Then
                lproj = Nothing
            End If
        End If

        Dim considerFapr As Boolean = Not IsNothing(bproj)
        Dim considerLapr As Boolean = Not IsNothing(lproj)




        ' jetzt wird SmartTableInfo gesetzt 
        ' jetzt wird die SmartTableInfo gesetzt 
        Dim emptyCollection As New Collection
        Call addSmartPPTTableInfo(pptShape,
                                  hproj.projectType, hproj.name, hproj.variantName, hproj.vpID,
                                  q1, q2, bigType, compID,
                                  emptyCollection)

        ' jetzt werden die einzelnen Zeilen geschrieben 

        ' in q1, q2 sind dei Anzahl Rollen bzw Kosten drin, sofern in toDoCollection was angegeben ist 
        ' die folgende Variable wird nur gebraucht, um im Falle  Auftreten Person-ID; und Person-ID; TeamID jedes Auftreten separat auszuwerten und 
        ' nicht bei Person-'ID die Summe aus allem aufzuschlüsseln 

        Dim takeITAsIs As Boolean = False
        Try

            If q2 = "-1" Or q2 = "%used%" Then
                takeITAsIs = True
                ' das ist das signal, dass erst die gemeinsame Liste bestimmt werden soll 
                toDoCollectionR = getCommonListOfRoleNameIDs(hproj, lproj, bproj)
                toDoCollectionC = getCommonListOfCostNames(hproj, lproj, bproj)
            Else
                ' es sind im q2 eine durch vblf bzw vbcr getrennte Rollen und Kosten angegeben

            End If
        Catch ex As Exception

        End Try

        Dim showOverviewOnly As Boolean = (toDoCollectionR.Count = 0)


        Try
            tabelle = pptShape.Table
            anzSpalten = tabelle.Columns.Count
            If anzSpalten = 6 Then
                ' dann ist alles in Ordnung .. 


                ' jetzt überprüfen, ob die Tabelle aktuell nur aus 2 Zeilen besteht ...
                If tabelle.Rows.Count > 2 Then
                    Do While tabelle.Rows.Count > 2
                        tabelle.Rows(2).Delete()
                    Loop
                End If

                ' jetzt die Werte in den 6 Spalten zurücksetzen 
                Try
                    With tabelle
                        For i As Integer = 1 To 6
                            .Cell(2, i).Shape.TextFrame2.TextRange.Text = ""
                        Next
                    End With
                Catch ex As Exception

                End Try


                Dim faprDate As Date = Date.MinValue
                Dim laprDate As Date = Date.MinValue
                Dim curDate As Date = Date.MinValue

                If Not IsNothing(hproj) Then
                    curDate = hproj.timeStamp
                End If
                If Not IsNothing(bproj) Then
                    faprDate = bproj.timeStamp
                End If
                If Not IsNothing(lproj) Then
                    laprDate = lproj.timeStamp
                End If

                ' jetzt die Headerzeile schreiben 
                Call schreibeAPVCVHeaderZeile(tabelle, faprDate, laprDate, curDate, considerFapr, considerLapr)

                Dim tabellenzeile As Integer = 2
                Try
                    ' erstmal in Abhängigkeit von der Rolle den Überblick zeichnen  
                    If myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager Or
                            myCustomUserRole.customUserRole = ptCustomUserRoles.ProjektLeitung Then

                        ' Überblick zeichnen ... 
                        Dim curPKI() As Double = {-1, -1, -1, -1}
                        Dim faprPKI() As Double = {-1, -1, -1, -1}
                        Dim laprPKI() As Double = {-1, -1, -1, -1}


                        Dim tmpValue As Double
                        Call hproj.calculateRoundedKPI(curPKI(0), curPKI(1), curPKI(2), tmpValue, curPKI(3), False)

                        If considerFapr Then
                            Call bproj.calculateRoundedKPI(faprPKI(0), faprPKI(1), faprPKI(2), tmpValue, faprPKI(3), False)
                        End If

                        If considerLapr Then
                            Call lproj.calculateRoundedKPI(laprPKI(0), laprPKI(1), laprPKI(2), tmpValue, laprPKI(3), False)
                        End If


                        ' jetzt das Gesamt Budget, Personalkosten, Sonstige Kosten und Ergebnis schreiben 

                        For i = 0 To 3
                            Call schreibeBudgetCostAPVCVZeile(tabelle, tabellenzeile, txtPKI(i), faprPKI(i), laprPKI(i), curPKI(i),
                                                          considerFapr, considerLapr)
                            tabelle.Rows.Add()
                            tabellenzeile = tabellenzeile + 1
                        Next

                        tabelle.Rows.Add()
                        tabellenzeile = tabellenzeile + 1

                    ElseIf myCustomUserRole.customUserRole = ptCustomUserRoles.RessourceManager Or myCustomUserRole.customUserRole = ptCustomUserRoles.TeamManager Then

                        Dim curItem As String = myCustomUserRole.specifics
                        Dim isRole As Boolean = RoleDefinitions.containsNameOrID(curItem)

                        If isRole Then

                            curValue = hproj.getRessourcenBedarf(curItem, inclSubRoles:=True, outPutInEuro:=showEuro).Sum

                            If considerLapr Then
                                laprValue = lproj.getRessourcenBedarf(curItem, inclSubRoles:=True, outPutInEuro:=showEuro).Sum
                            Else
                                laprValue = 0.0
                            End If

                            If considerFapr Then
                                faprValue = bproj.getRessourcenBedarf(curItem, inclSubRoles:=True, outPutInEuro:=showEuro).Sum
                            Else
                                faprValue = 0.0
                            End If

                            Dim zeilenItem As String = curItem


                            Call schreibeBudgetCostAPVCVZeile(tabelle, tabellenzeile, zeilenItem, faprValue, laprValue, curValue,
                                                      considerFapr, considerLapr)
                            tabelle.Rows.Add()
                            tabelle.Rows.Add()
                            tabellenzeile = tabellenzeile + 2

                        End If

                    End If

                    ' ---------------------------------------------------------------------
                    ' wenn dann noch Details gezeigt werden sollen ... 
                    ' 
                    If Not showOverviewOnly Then

                        If myCustomUserRole.customUserRole = ptCustomUserRoles.RessourceManager Or myCustomUserRole.customUserRole = ptCustomUserRoles.TeamManager Then
                            ' den im Überblick gezeigten specifics nicht noch mal zeigen, falls der aufgeführt ist ... 
                            Dim tmpNameID As String = myCustomUserRole.specifics
                            If toDoCollectionR.Contains(tmpNameID) Then
                                toDoCollectionR.Remove(tmpNameID)
                            End If
                        End If

                        einrueckung = 1

                        ' dient dazu , zu bestimmen, wann die Kostenarten kommen um vorher eine Neue Zeile  einzufügen ...
                        Dim firstCost As Boolean = True


                        ' keine zusätzliche Zeile schreiben ... macht das ganze nur unübersichtlicher  
                        'If anzRoles > 0 And anzCosts > 0 Then
                        '    ' 
                        '    'tabelle.Cell(tabellenzeile, 1).Shape.TextFrame2.TextRange.Text = repMessages.getmsg(51)
                        '    tabelle.Cell(tabellenzeile, 1).Shape.TextFrame2.TextRange.Text = repmsg(1)
                        '    tabelle.Rows.Add()
                        '    tabellenzeile = tabellenzeile + 1
                        'End If

                        For m As Integer = 1 To toDoCollectionR.Count

                            ' wegen Einrückung in Details ...
                            Dim curItem As String = CStr(toDoCollectionR.Item(m))
                            Dim isRole As Boolean = RoleDefinitions.containsNameOrID(curItem)

                            If isRole Then

                                curValue = hproj.getRessourcenBedarf(curItem, inclSubRoles:=True,
                                                                     outPutInEuro:=showEuro, takeITAsIs:=takeITAsIs).Sum

                                If considerLapr Then
                                    laprValue = lproj.getRessourcenBedarf(curItem, inclSubRoles:=True,
                                                                          outPutInEuro:=showEuro, takeITAsIs:=takeITAsIs).Sum
                                Else
                                    laprValue = 0.0
                                End If

                                If considerFapr Then
                                    faprValue = bproj.getRessourcenBedarf(curItem, inclSubRoles:=True,
                                                                          outPutInEuro:=showEuro, takeITAsIs:=takeITAsIs).Sum
                                Else
                                    faprValue = 0.0
                                End If

                                Dim zeilenItem As String = curItem


                                Call schreibeBudgetCostAPVCVZeile(tabelle, tabellenzeile, zeilenItem, faprValue, laprValue, curValue,
                                                          considerFapr, considerLapr)
                                tabelle.Rows.Add()
                                tabellenzeile = tabellenzeile + 1

                            End If

                        Next

                        For m As Integer = 1 To toDoCollectionC.Count

                            ' wegen Einrückung in Details ...
                            Dim curItem As String = CStr(toDoCollectionC.Item(m))
                            Dim isCost As Boolean = CostDefinitions.containsName(curItem)


                            If isCost Then

                                curValue = hproj.getKostenBedarfNew(curItem).Sum

                                If considerLapr Then
                                    laprValue = lproj.getKostenBedarfNew(curItem).Sum
                                Else
                                    laprValue = 0.0
                                End If

                                If considerFapr Then
                                    faprValue = bproj.getKostenBedarfNew(curItem).Sum
                                Else
                                    faprValue = 0.0
                                End If

                                Dim zeilenItem As String = curItem

                                Call schreibeBudgetCostAPVCVZeile(tabelle, tabellenzeile, zeilenItem, faprValue, laprValue, curValue,
                                                          considerFapr, considerLapr)
                                tabelle.Rows.Add()
                                tabellenzeile = tabellenzeile + 1

                            End If


                        Next


                    End If

                    ' jetzt letzte Zeile löschen  ...
                    tabelle.Rows(tabellenzeile).Delete()
                    tabellenzeile = tabellenzeile - 1

                Catch ex1 As Exception

                End Try
            Else
                Throw New Exception("Tabelle should have 6 columns ... exit ...")
            End If
        Catch ex As Exception

        End Try

    End Sub

    ''' <summary>
    ''' ergänzt den Text in der Tabelle BudgetCOst Approved versions versus current Version
    ''' 
    ''' </summary>
    ''' <param name="table"></param>
    ''' <param name="faprDate"></param>
    ''' <param name="laprDate"></param>
    ''' <param name="curDate"></param>
    Private Sub schreibeAPVCVHeaderZeile(ByRef table As pptNS.Table,
                                                   ByVal faprDate As Date, ByVal laprDate As Date, ByVal curDate As Date,
                                                   ByVal considerFapr As Boolean, ByVal considerLapr As Boolean)

        With table

            Dim faprText As String
            Dim laprText As String
            Dim curText As String

            If Not considerFapr Then
                faprDate = Date.MinValue
            End If

            If Not considerLapr Then
                laprDate = Date.MinValue
            End If

            curText = addDateToText(table.Cell(1, 2).Shape.TextFrame2.TextRange.Text, curDate)
            laprText = addDateToText(table.Cell(1, 3).Shape.TextFrame2.TextRange.Text, laprDate)
            faprText = addDateToText(table.Cell(1, 5).Shape.TextFrame2.TextRange.Text, faprDate)

            table.Cell(1, 2).Shape.TextFrame2.TextRange.Text = curText
            table.Cell(1, 3).Shape.TextFrame2.TextRange.Text = laprText
            table.Cell(1, 5).Shape.TextFrame2.TextRange.Text = faprText


        End With

    End Sub

    ''' <summary>
    ''' ergänzt den String header um vbVerticalTab (myDate.toString)
    ''' </summary>
    ''' <param name="header"></param>
    ''' <param name="myDate"></param>
    ''' <returns></returns>
    Private Function addDateToText(ByVal header As String, ByVal myDate As Date) As String

        Dim tmpResult As String = ""
        Dim dateString As String = ""
        Dim tmpStr() As String = header.Split(New Char() {CChar("("), CChar(")")})

        If DateDiff(DateInterval.Day, Date.MinValue, myDate) > 0 Then
            dateString = "(" & myDate.ToShortDateString & ")"
        End If

        If tmpStr(0).EndsWith(vbVerticalTab) Then
            tmpResult = tmpStr(0) & dateString
        Else
            tmpResult = tmpStr(0) & vbVerticalTab & dateString
        End If


        addDateToText = tmpResult
    End Function

    ''' <summary>
    ''' schreibt eine Zeile in die Tabelle BudgetCost Approved Versions versus curVersion
    ''' </summary>
    ''' <param name="table"></param>
    ''' <param name="zeile"></param>
    ''' <param name="itemNameID"></param>
    ''' <param name="faprValue"></param>
    ''' <param name="laprValue"></param>
    ''' <param name="curValue"></param>
    Private Sub schreibeBudgetCostAPVCVZeile(ByRef table As pptNS.Table, ByVal zeile As Integer,
                                             ByVal itemNameID As String, ByVal faprValue As Double, ByVal laprValue As Double, ByVal curValue As Double,
                                             ByVal considerFapr As Boolean, ByVal considerLapr As Boolean)

        Dim deltaFMC As String = "-" ' niummt das Delta auf zwischen Fapr und Current: First minsu Current 
        Dim deltaLMC As String = "-" ' nimmt das Delta auf zwischen Lapr und Current : last minus Current
        Dim dblFormat As String = "#,##0.00"
        Dim cellText As String = "-"
        Dim nada As String = "-"
        Dim isPositiv As Boolean = False

        ' notwendig, solange keine repMessages in der Datenbank sind 
        Dim repmsg() As String
        repmsg = {"Budget", "Personalkosten", "Sonstige Kosten", "Ergebnis-Prognose"}

        Dim roleBezeichner As String = ""

        If repmsg.Contains(itemNameID) Then
            roleBezeichner = itemNameID
        ElseIf RoleDefinitions.containsNameOrID(itemNameID) Then
            roleBezeichner = RoleDefinitions.getBezeichner(itemNameID)
        ElseIf CostDefinitions.containsName(itemNameID) Then
            roleBezeichner = itemNameID
        Else
            If awinSettings.englishLanguage Then
                Call MsgBox("Role/Cost ID " & itemNameID & " isn't defined yet")
            Else
                Call MsgBox("Rolle oder Kostenart " & itemNameID & " ist nicht in der Organisation enthalten")
            End If
            Exit Sub
        End If


        If considerFapr Then
            deltaFMC = (curValue - faprValue).ToString(dblFormat)
        Else
            deltaFMC = nada
        End If

        If considerLapr Then
            deltaLMC = (curValue - laprValue).ToString(dblFormat)
        Else
            deltaLMC = nada
        End If


        ' jetzt wird das geschrieben 
        With table
            Dim tmpValue As String = "-"

            ' Label schreiben
            CType(.Cell(zeile, 1), pptNS.Cell).Shape.TextFrame2.TextRange.Text = roleBezeichner


            ' wird benötigt, um die Schriftfarbe im Delta-Feld wieder auf Normal setzen zu können 
            Dim normalColor As Integer = CType(.Cell(zeile, 1), pptNS.Cell).Shape.TextFrame.TextRange.Font.Color.RGB

            ' Current Value schreiben 
            cellText = curValue.ToString(dblFormat)
            CType(.Cell(zeile, 2), pptNS.Cell).Shape.TextFrame2.TextRange.Text = cellText

            ' last Approved Value schreiben  
            If considerLapr Then
                cellText = laprValue.ToString(dblFormat)
            Else
                cellText = nada
            End If
            CType(.Cell(zeile, 3), pptNS.Cell).Shape.TextFrame2.TextRange.Text = cellText

            ' Delta schreiben 
            CType(.Cell(zeile, 4), pptNS.Cell).Shape.TextFrame2.TextRange.Text = deltaLMC

            ' ggf einfärben 
            If System.Math.Abs(curValue - laprValue) <= 0.5 Then
                ' nichts tun, ausser Farbe auf Nortmal setzen 
                ' das ist notwednig, weil durch den .add Row in der übergeordneten Sub evtl die dort verwendete Farbe Grün oder Rot zur Geltung kommt 
                CType(.Cell(zeile, 4), pptNS.Cell).Shape.TextFrame.TextRange.Font.Color.RGB = normalColor

            ElseIf considerLapr Then

                If roleBezeichner = repmsg(0) Or roleBezeichner = repmsg(3) Then
                    isPositiv = (curValue > laprValue + 0.5)
                Else
                    isPositiv = (laprValue > curValue + 0.5)
                End If


                ' Delta entsprechend einfärben 
                If isPositiv Then
                    CType(.Cell(zeile, 4), pptNS.Cell).Shape.TextFrame.TextRange.Font.Color.RGB = visboFarbeGreen
                Else
                    CType(.Cell(zeile, 4), pptNS.Cell).Shape.TextFrame.TextRange.Font.Color.RGB = visboFarbeRed
                End If


            End If

            ' first Approved Value schreiben  
            If considerFapr Then
                cellText = faprValue.ToString(dblFormat)
            Else
                cellText = nada
            End If
            CType(.Cell(zeile, 5), pptNS.Cell).Shape.TextFrame2.TextRange.Text = cellText

            ' Delta schreiben 
            CType(.Cell(zeile, 6), pptNS.Cell).Shape.TextFrame2.TextRange.Text = deltaFMC

            ' ggf einfärben 
            If System.Math.Abs(curValue - faprValue) <= 0.5 Then
                ' nichts tun, ausser Farbe auf Nortmal setzen 
                ' das ist notwednig, weil durch den .add Row in der übergeordneten Sub evtl die dort verwendete Farbe Grün oder Rot zur Geltung kommt 
                CType(.Cell(zeile, 6), pptNS.Cell).Shape.TextFrame.TextRange.Font.Color.RGB = normalColor

            ElseIf considerFapr Then
                ' If itemName = repMessages.getmsg(49) Or itemName = repMessages.getmsg(53) Then
                If roleBezeichner = repmsg(0) Or roleBezeichner = repmsg(3) Then
                    isPositiv = (curValue > faprValue)
                Else
                    isPositiv = (faprValue > curValue)
                End If

                ' Delta entsprechend einfärben 
                If isPositiv Then
                    CType(.Cell(zeile, 6), pptNS.Cell).Shape.TextFrame.TextRange.Font.Color.RGB = visboFarbeGreen
                Else
                    CType(.Cell(zeile, 6), pptNS.Cell).Shape.TextFrame.TextRange.Font.Color.RGB = visboFarbeRed
                End If

            End If

        End With


    End Sub

    ''' <summary>
    ''' schreibt eine Zeile in die Tabelle Milestones Approved Versions versus curVersion
    ''' </summary>
    ''' <param name="table"></param>
    ''' <param name="zeile"></param>
    ''' <param name="itemName"></param>
    ''' <param name="faprValue"></param>
    ''' <param name="laprValue"></param>
    ''' <param name="curValue"></param>
    ''' <param name="considerFapr"></param>
    ''' <param name="considerLapr"></param>
    Private Sub schreibeMilestoneAPVCVZeile(ByRef table As pptNS.Table, ByVal zeile As Integer,
                                             ByVal itemName As String, ByVal faprValue As Date, ByVal laprValue As Date, ByVal curValue As Date,
                                             ByVal considerFapr As Boolean, ByVal considerLapr As Boolean)

        Dim deltaFMC As Long = 0 ' niummt das Delta auf zwischen Fapr und Current: First minus Current 
        Dim deltaLMC As Long = 0 ' nimmt das Delta auf zwischen Lapr und Current : last minus Current

        Dim cellText As String = "-"
        Dim nada As String = "-"
        Dim isPositiv As Boolean = False


        If considerFapr And faprValue > Date.MinValue Then
            deltaFMC = DateDiff(DateInterval.Day, curValue.Date, faprValue.Date)
        Else
            deltaFMC = 0
        End If

        If considerLapr And laprValue > Date.MinValue Then
            deltaLMC = DateDiff(DateInterval.Day, curValue.Date, laprValue.Date)
        Else
            deltaLMC = 0
        End If


        ' jetzt wird das geschrieben 
        With table
            Dim tmpValue As String = "-"

            ' Label schreiben
            CType(.Cell(zeile, 1), pptNS.Cell).Shape.TextFrame2.TextRange.Text = itemName


            ' wird benötigt, um die Schriftfarbe im Delta-Feld wieder auf Normal setzen zu können 
            Dim normalColor As Integer = CType(.Cell(zeile, 1), pptNS.Cell).Shape.TextFrame.TextRange.Font.Color.RGB

            ' Current Value schreiben 
            cellText = curValue.ToShortDateString
            CType(.Cell(zeile, 2), pptNS.Cell).Shape.TextFrame2.TextRange.Text = cellText

            ' last Approved Value schreiben  
            If considerLapr And laprValue > Date.MinValue Then
                cellText = laprValue.ToShortDateString
            Else
                cellText = nada
            End If

            CType(.Cell(zeile, 3), pptNS.Cell).Shape.TextFrame2.TextRange.Text = cellText

            ' Delta schreiben 
            If cellText = nada Then
                CType(.Cell(zeile, 4), pptNS.Cell).Shape.TextFrame2.TextRange.Text = ""
            Else
                CType(.Cell(zeile, 4), pptNS.Cell).Shape.TextFrame2.TextRange.Text = deltaLMC.ToString
            End If


            ' ggf einfärben 
            If deltaLMC = 0 Then
                ' nichts tun, ausser Farbe auf Normal setzen 
                ' das ist notwednig, weil durch den .add Row in der übergeordneten Sub evtl die dort verwendete Farbe Grün oder Rot zur Geltung kommt 
                CType(.Cell(zeile, 4), pptNS.Cell).Shape.TextFrame.TextRange.Font.Color.RGB = normalColor

            ElseIf considerLapr And laprValue > Date.MinValue Then

                isPositiv = (deltaLMC > 0)

                ' Delta entsprechend einfärben 
                If isPositiv Then
                    CType(.Cell(zeile, 4), pptNS.Cell).Shape.TextFrame.TextRange.Font.Color.RGB = visboFarbeGreen
                Else
                    CType(.Cell(zeile, 4), pptNS.Cell).Shape.TextFrame.TextRange.Font.Color.RGB = visboFarbeRed
                End If


            End If

            ' first Approved Value schreiben  
            If considerFapr And faprValue > Date.MinValue Then
                cellText = faprValue.ToShortDateString
            Else
                cellText = nada
            End If

            CType(.Cell(zeile, 5), pptNS.Cell).Shape.TextFrame2.TextRange.Text = cellText

            ' Delta schreiben 
            If cellText = nada Then
                CType(.Cell(zeile, 6), pptNS.Cell).Shape.TextFrame2.TextRange.Text = ""
            Else
                CType(.Cell(zeile, 6), pptNS.Cell).Shape.TextFrame2.TextRange.Text = deltaFMC.ToString
            End If


            ' ggf einfärben 
            If deltaFMC = 0 Then
                ' nichts tun, ausser Farbe auf Normal setzen 
                ' das ist notwendig, weil durch den .add Row in der übergeordneten Sub evtl die dort verwendete Farbe Grün oder Rot zur Geltung kommt 
                CType(.Cell(zeile, 6), pptNS.Cell).Shape.TextFrame.TextRange.Font.Color.RGB = normalColor

            ElseIf considerFapr And faprValue > Date.MinValue Then
                ' If itemName = repMessages.getmsg(49) Or itemName = repMessages.getmsg(53) Then
                isPositiv = (deltaFMC > 0)

                ' Delta entsprechend einfärben 
                If isPositiv Then
                    CType(.Cell(zeile, 6), pptNS.Cell).Shape.TextFrame.TextRange.Font.Color.RGB = visboFarbeGreen
                Else
                    CType(.Cell(zeile, 6), pptNS.Cell).Shape.TextFrame.TextRange.Font.Color.RGB = visboFarbeRed
                End If

            End If

        End With


    End Sub

    ''' <summary>
    ''' wird benötigt, um Collections in Shape.Tags unterzubringen 
    ''' </summary>
    ''' <param name="nameCollection"></param>
    ''' <returns></returns>
    Public Function convertCollToNids(ByVal nameCollection As Collection) As String
        Dim nids As String = ""
        For Each tmpName As String In nameCollection

            If tmpName.Contains("#") Then
                tmpName = tmpName.Replace("#", "^")
            End If

            If nids = "" Then
                nids = tmpName.Trim
            Else
                nids = nids & "#" & tmpName.Trim
            End If
        Next
        convertCollToNids = nids
    End Function

    ''' <summary>
    ''' wird benötigt, um Collection-Infos aus Shape.Tags wieder in eine Collection zu bringen  
    ''' </summary>
    ''' <param name="nids"></param>
    ''' <returns></returns>
    Public Function convertNidsToColl(ByVal nids As String) As Collection
        Dim nameCollection As New Collection
        Dim tmpStr() As String = nids.Split(New Char() {CChar("#")})

        For Each tmpName In tmpStr

            If tmpName.Contains("^") Then
                tmpName = tmpName.Replace("^", "#")
            End If

            If tmpName.Trim.Length > 0 Then
                nameCollection.Add(tmpName)
            End If
        Next
        convertNidsToColl = nameCollection
    End Function


    ''' <summary>
    ''' bestimmt aus dem Namen eines Charts die Informationen, die benötigt werden, um ein PPTShape aus der Smart-PPT heraus zu aktualisieren ...
    ''' </summary>
    ''' <param name="chtObjName"></param>
    ''' <param name="prpfTyp"></param>
    ''' <param name="pName"></param>
    ''' <param name="vName"></param>
    ''' <param name="chartTyp"></param>
    ''' <param name="auswahl"></param>
    ''' <remarks></remarks>
    Public Sub bestimmeChartInfosFromName(ByVal chtObjName As String,
                                              ByRef prpfTyp As Integer,
                                              ByRef prcTyp As Integer,
                                              ByRef pName As String,
                                              ByRef vName As String,
                                              ByRef chartTyp As Integer,
                                              ByRef auswahl As Integer)


        Dim tmpStr() As String = chtObjName.Split(New Char() {CChar("#")})

        ' bestimme, ob es sich um ein pf oder pr Diagramm handelt ...  

        If tmpStr(0) = "pr" Then
            ' bestimme den Charttyp ...
            prpfTyp = ptPRPFType.project

        ElseIf tmpStr(0) = "pf" Then
            prpfTyp = ptPRPFType.portfolio
        Else

        End If

        chartTyp = CInt(tmpStr(1))

        If chartTyp = PTprdk.KostenBalken Or
            chartTyp = PTprdk.KostenPie Then
            prcTyp = ptElementTypen.costs
        ElseIf chartTyp = PTprdk.PersonalBalken Or
            chartTyp = PTprdk.PersonalPie Then
            prcTyp = ptElementTypen.roles
        Else
            prcTyp = ptElementTypen.ergebnis
        End If


        ' bestimme pName und vName 
        Dim fullName As String = tmpStr(2)

        If fullName.Contains("[") And fullName.Contains("]") Then
            Dim tmpstr1() As String = fullName.Split(New Char() {CChar("["), CChar("]")})
            pName = tmpstr1(0)
            vName = tmpstr1(1)
        Else
            pName = fullName
            vName = ""
        End If

        ' bestimme, um welche Auswahl es sich handelt ... 
        auswahl = CInt(tmpStr(3))


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

    ''' <summary>
    ''' setzt maxScreenWidth, maxScreenHeight und die dazugehörigen Default Windows- und Chartbreiten 
    ''' </summary>
    Public Sub setWindowParameters()
        With appInstance.ActiveWindow

            If .WindowState = Excel.XlWindowState.xlMaximized Then
                'maxScreenHeight = .UsableHeight
                maxScreenHeight = .Height
                'maxScreenWidth = .UsableWidth
                maxScreenWidth = .Width
            Else
                'Dim formerState As Excel.XlWindowState = .WindowState
                .WindowState = Excel.XlWindowState.xlMaximized
                'maxScreenHeight = .UsableHeight
                maxScreenHeight = .Height
                'maxScreenWidth = .UsableWidth
                maxScreenWidth = .Width
                '.WindowState = formerState
            End If


        End With

        ' jetzt das ProjectboardWindows (0) setzen 
        projectboardWindows(PTwindows.mpt) = appInstance.ActiveWindow

        chartHeight = maxScreenHeight / 6
        chartWidth = maxScreenWidth / 5

        If chartHeight < 120 Then
            chartHeight = 120
        End If

        If chartWidth < 140 Then
            chartWidth = 140
        End If

        Dim oGrenze As Integer = UBound(frmCoord, 1)
        ' hier werden die Top- & Left- Default Positionen der Formulare gesetzt 
        For i = 0 To oGrenze
            frmCoord(i, PTpinfo.top) = maxScreenHeight * 0.3
            frmCoord(i, PTpinfo.left) = maxScreenWidth * 0.4
        Next

        ' jetzt setzen der Werte für Status-Information und Milestone-Information
        frmCoord(PTfrm.projInfo, PTpinfo.top) = 125
        frmCoord(PTfrm.projInfo, PTpinfo.left) = My.Computer.Screen.WorkingArea.Width - 500

        frmCoord(PTfrm.msInfo, PTpinfo.top) = 125 + 280
        frmCoord(PTfrm.msInfo, PTpinfo.left) = My.Computer.Screen.WorkingArea.Width - 500
    End Sub


    ''' <summary>
    ''' initialisert das Logfile
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub logfileInit()

        Try

            With CType(xlsLogfile.Worksheets(1), Excel.Worksheet)
                .Name = "logBuch"
                CType(.Cells(1, 1), Excel.Range).Value = "logfile erzeugt " & Date.Now.ToString
                CType(.Columns(1), Excel.Range).ColumnWidth = 100
                CType(.Columns(2), Excel.Range).ColumnWidth = 50
                CType(.Columns(3), Excel.Range).ColumnWidth = 20
            End With
        Catch ex As Exception

        End Try


    End Sub
    ''' <summary>
    ''' schreibt in das logfile 
    ''' </summary>
    ''' <param name="text"></param>
    ''' <param name="addOn"></param>
    ''' <remarks></remarks>
    Public Sub logfileSchreiben(ByVal text As String, ByVal addOn As String, ByRef anzFehler As Long)

        Dim obj As Object

        Try
            obj = CType(CType(xlsLogfile.Worksheets("logBuch"), Excel.Worksheet).Rows(1), Excel.Range).Insert(Excel.XlInsertShiftDirection.xlShiftDown)

            With CType(xlsLogfile.Worksheets("logBuch"), Excel.Worksheet)
                CType(.Cells(1, 1), Excel.Range).Value = text
                CType(.Cells(1, 2), Excel.Range).Value = addOn
                CType(.Cells(1, 3), Excel.Range).Value = Date.Now
                CType(.Cells(1, 3), Excel.Range).NumberFormat = "m/d/yyyy h:mm"
            End With
            anzFehler = anzFehler + 1


        Catch ex As Exception

        End Try

    End Sub

    ''' <summary>
    ''' schreibt die Inhalte der Collection als String in das Logfile
    ''' </summary>
    ''' <param name="meldungen"></param>
    Public Sub logfileSchreiben(ByVal meldungen As Collection)
        Dim obj As Object
        Dim anzZeilen As Integer = meldungen.Count

        Try

            For i As Integer = 1 To anzZeilen

                ' neue Zeile einfügen 
                obj = CType(CType(xlsLogfile.Worksheets("logBuch"), Excel.Worksheet).Rows(1), Excel.Range).Insert(Excel.XlInsertShiftDirection.xlShiftDown)

                Dim text As String = CStr(meldungen.Item(i))
                With CType(xlsLogfile.Worksheets("logBuch"), Excel.Worksheet)
                    CType(.Cells(1, 1), Excel.Range).Value = text
                End With
            Next

        Catch ex As Exception

        End Try
    End Sub

    ''' <summary>
    ''' ganz aanlog zu dem anderen logfile Schrieben, nur dass jetzt ein Array von String Werten übergeben wird, der in die einzelnen Spalten kommt 
    ''' </summary>
    ''' <param name="text"></param>
    Public Sub logfileSchreiben(ByVal text() As String)

        Dim obj As Object
        Try
            Dim anzSpalten As Integer = text.Length
            obj = CType(CType(xlsLogfile.Worksheets("logBuch"), Excel.Worksheet).Rows(1), Excel.Range).Insert(Excel.XlInsertShiftDirection.xlShiftDown)

            With CType(xlsLogfile.Worksheets("logBuch"), Excel.Worksheet)
                For ix As Integer = 1 To anzSpalten
                    CType(.Cells(1, ix), Excel.Range).NumberFormat = "@"
                    CType(.Cells(1, ix), Excel.Range).Value = text(ix - 1)
                Next
                CType(.Cells(1, anzSpalten + 1), Excel.Range).Value = Date.Now
                CType(.Cells(1, anzSpalten + 1), Excel.Range).NumberFormat = "m/d/yyyy h:mm"
            End With
        Catch ex As Exception

        End Try

    End Sub

    Public Sub logfileSchreiben(ByVal text() As String, ByVal values() As Double)

        Dim obj As Object
        Try
            Dim anzSpaltenText As Integer = text.Length
            Dim anzSpaltenValues As Integer = values.Length
            obj = CType(CType(xlsLogfile.Worksheets("logBuch"), Excel.Worksheet).Rows(1), Excel.Range).Insert(Excel.XlInsertShiftDirection.xlShiftDown)

            With CType(xlsLogfile.Worksheets("logBuch"), Excel.Worksheet)
                For ix As Integer = 1 To anzSpaltenText
                    CType(.Cells(1, ix), Excel.Range).NumberFormat = "@"
                    CType(.Cells(1, ix), Excel.Range).Value = text(ix - 1)
                Next

                For ix As Integer = 1 To anzSpaltenValues
                    CType(.Cells(1, ix + anzSpaltenText), Excel.Range).Value = values(ix - 1)
                    CType(.Cells(1, ix + anzSpaltenText), Excel.Range).NumberFormat = "#,##0.##"
                Next
                CType(.Cells(1, anzSpaltenText + anzSpaltenValues + 1), Excel.Range).Value = Date.Now
                CType(.Cells(1, anzSpaltenText + anzSpaltenValues + 1), Excel.Range).NumberFormat = "m/d/yyyy h:mm"
            End With


        Catch ex As Exception

        End Try

    End Sub

    ''' <summary>
    ''' öffnet das LogFile
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub logfileOpen()

        Dim formerSU As Boolean = appInstance.ScreenUpdating
        If appInstance.ScreenUpdating Then
            appInstance.ScreenUpdating = False
        End If


        ' aktives Workbook merken im Variable actualWB
        Dim actualWB As String = appInstance.ActiveWorkbook.Name

        Dim logfileOrdner As String = "logfiles"
        Dim logfilePath As String = My.Computer.FileSystem.CombinePath(awinPath, logfileOrdner)
        Dim logfileName As String = "logfile" & "_" & Date.Now.Year.ToString & Date.Now.Month.ToString("0#") & Date.Now.Day.ToString("0#") & "_" & Date.Now.TimeOfDay.ToString.Replace(":", "-") & ".xlsx"
        Dim logfileNamePath As String = My.Computer.FileSystem.CombinePath(logfilePath, logfileName)

        If Not My.Computer.FileSystem.DirectoryExists(logfilePath) Then
            My.Computer.FileSystem.CreateDirectory(logfilePath)
        End If

        ' Prüfen, ob es bereits ein offenes Logfile gibt ... 
        Try
            If myLogfile <> "" Then
                Call logfileSchliessen()
            End If
        Catch ex As Exception

        End Try

        Try
            ' Logfile neu anlegen 
            xlsLogfile = appInstance.Workbooks.Add
            ' schreibt Sheet Namen in logfile ...  
            Call logfileInit()

            xlsLogfile.SaveAs(logfileNamePath)
            myLogfile = xlsLogfile.Name

        Catch ex As Exception
            logmessage = "Erzeugen von " & logfileNamePath & " fehlgeschlagen" & vbLf &
                                            "bitte schliessen Sie die Anwendung und kontaktieren Sie ggf. ihren System-Administrator"
            appInstance.ScreenUpdating = True
            Throw New ArgumentException(logmessage)
        End Try



        ' Workbook, das vor dem öffnen des Logfiles aktiv war, wieder aktivieren
        appInstance.Workbooks(actualWB).Activate()

        If appInstance.ScreenUpdating <> formerSU Then
            appInstance.ScreenUpdating = formerSU
        End If


    End Sub



    ''' <summary>
    ''' schliesst  das logfile 
    ''' </summary>  
    ''' <remarks></remarks>
    Public Sub logfileSchliessen()

        appInstance.EnableEvents = False

        Try

            If myLogfile <> "" Then
                appInstance.Workbooks(myLogfile).Close(SaveChanges:=True)
                myLogfile = ""
            End If


        Catch ex As Exception
            Call MsgBox("Fehler beim Schließen des Logfiles")
        End Try

        appInstance.EnableEvents = True
    End Sub

    ''' <summary>
    ''' zeigt die in der OutputCollection gesammelten Rückmeldungen in einem Fenster mit Scrollbar 
    ''' </summary>
    ''' <param name="outPutCollection"></param>
    ''' <param name="header"></param>
    ''' <param name="explanation"></param>
    ''' <remarks></remarks>
    Public Sub showOutPut(ByVal outPutCollection As Collection, ByVal header As String, ByVal explanation As String)
        If outPutCollection.Count > 0 Then

            Dim outputFormular As New frmOutputWindow
            With outputFormular
                .Text = header
                .lblOutput.Text = explanation
                .textCollection = outPutCollection
                .ShowDialog()
            End With

        End If
    End Sub


    ''' <summary>
    ''' liefert zurück, ob in dem angegebenen Sheet überhaupt Charts vorhanden sind ... 
    ''' </summary>
    ''' <param name="chType"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function thereAreAnyCharts(ByVal chType As Integer) As Boolean

        Dim anzCharts As Integer = 0


        Try
            If chType = PTwindows.mptpf Then
                anzCharts = CType(CType(CType(appInstance.Workbooks.Item(myProjektTafel), Excel.Workbook) _
                    .Worksheets.Item(arrWsNames(ptTables.mptPfCharts)), Excel.Worksheet).ChartObjects, Excel.ChartObjects).Count

            ElseIf chType = PTwindows.mptpr Then
                anzCharts = CType(CType(CType(appInstance.Workbooks.Item(myProjektTafel), Excel.Workbook) _
                    .Worksheets.Item(arrWsNames(ptTables.mptPrCharts)), Excel.Worksheet).ChartObjects, Excel.ChartObjects).Count

            ElseIf chType = PTwindows.meChart Then
                anzCharts = CType(CType(CType(appInstance.Workbooks.Item(myProjektTafel), Excel.Workbook) _
                    .Worksheets.Item(arrWsNames(ptTables.meCharts)), Excel.Worksheet).ChartObjects, Excel.ChartObjects).Count

            End If

        Catch ex As Exception

        End Try


        thereAreAnyCharts = (anzCharts > 0)

    End Function

    ''' <summary>
    ''' gibt zurück ob das angegebene Window existiert
    ''' wird unter anderem benötigt, um die Caption zu aktualisieren 
    ''' </summary>
    ''' <param name="windowTyp"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function visboWindowExists(ByVal windowTyp As Integer) As Boolean

        Dim tmpResult As Boolean = False
        Dim sheetNameLookedFor As String = ""

        Try
            Select Case windowTyp
                Case PTwindows.mpt
                    sheetNameLookedFor = arrWsNames(ptTables.MPT)
                Case PTwindows.mptpf
                    sheetNameLookedFor = arrWsNames(ptTables.mptPfCharts)
                Case PTwindows.mptpr
                    sheetNameLookedFor = arrWsNames(ptTables.mptPrCharts)
                Case PTwindows.meChart
                    sheetNameLookedFor = arrWsNames(ptTables.meCharts)
                Case PTwindows.massEdit
                    sheetNameLookedFor = arrWsNames(ptTables.meRC)
                Case Else
                    sheetNameLookedFor = "XX?"
            End Select

            For i As Integer = 1 To CType(appInstance.Workbooks.Item(myProjektTafel), Excel.Workbook).Windows.Count
                If CType(CType(CType(appInstance.Workbooks.Item(myProjektTafel), Excel.Workbook).Windows.Item(i), Excel.Window).ActiveSheet, Excel.Worksheet) _
                    .Name = sheetNameLookedFor Then
                    tmpResult = Not IsNothing(projectboardWindows(windowTyp))
                End If
            Next
        Catch ex As Exception

        End Try
        visboWindowExists = tmpResult
    End Function

    ''' <summary>
    ''' liefert den Caption Namen des Windows in Abhängigkeit von Portfolio oder Projekt und in Abhängigkeit von der Sprache 
    ''' </summary>
    ''' <param name="visboWindowTyp"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function bestimmeWindowCaption(ByVal visboWindowTyp As Integer,
                                          Optional ByVal tableTyp As Integer = ptTables.meRC,
                                          Optional ByVal addOnMsg As String = "") As String
        Dim tmpResult As String = ""

        Select Case visboWindowTyp

            Case PTwindows.mptpf
                If awinSettings.englishLanguage Then
                    tmpResult = "Charts for Portfolio '" & currentConstellationName & "'"
                Else
                    tmpResult = "Charts für Portfolio " & currentConstellationName & "'"
                End If

            Case PTwindows.mptpr

                tmpResult = "Charts: " & addOnMsg


            Case PTwindows.meChart
                If awinSettings.englishLanguage Then
                    tmpResult = "Project-Chart and Portfolio-Charts '" & currentConstellationName & "': " & ShowProjekte.Count & " projects"
                Else
                    tmpResult = "Projekt-Chart und Portfolio-Charts '" & currentConstellationName & "': " & ShowProjekte.Count & " Projekte"
                End If

            Case PTwindows.mpt
                Dim outputmsg As String = ""
                Dim roleName As String = myCustomUserRole.customUserRole.ToString

                If myCustomUserRole.customUserRole = ptCustomUserRoles.RessourceManager Or myCustomUserRole.customUserRole = ptCustomUserRoles.TeamManager Then
                    Dim teamID As Integer = -1
                    roleName = roleName & " " & RoleDefinitions.getRoleDefByIDKennung(myCustomUserRole.specifics, teamID).name
                ElseIf myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager Then

                End If

                If currentConstellationName = "" Then
                    outputmsg = " : " & ShowProjekte.Count & " "
                Else
                    outputmsg = " '" & currentConstellationName & "' : " & ShowProjekte.Count & " "
                End If

                tmpResult = roleName & outputmsg & "objects"

                If awinSettings.englishLanguage Then
                    tmpResult = "Projectboard ( " & roleName & " ) " & outputmsg & "objects"
                Else
                    tmpResult = "Projectboard ( " & roleName & " ) " & outputmsg & "Objekte"
                End If

            Case PTwindows.massEdit

                Select Case tableTyp
                    Case ptTables.meRC
                        If awinSettings.englishLanguage Then
                            tmpResult = "Modify Resource and Cost Needs"
                        Else
                            tmpResult = "Personal- und Kostenbedarfe ändern"
                        End If
                    Case ptTables.meTE
                        If awinSettings.englishLanguage Then
                            tmpResult = "Modify Tasks and Milestones"
                        Else
                            tmpResult = "Meilensteine und Vorgänge ändern"
                        End If
                    Case ptTables.meAT
                        If awinSettings.englishLanguage Then
                            tmpResult = "Modify Attributes"
                        Else
                            tmpResult = "Attribute ändern"
                        End If
                End Select


        End Select

        bestimmeWindowCaption = tmpResult
    End Function

    ''' <summary>
    ''' schliesst alle Windows ausser MPT Window; macht dann das MPT Window wieder groß 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub closeAllWindowsExceptMPT()


        Dim vglName As String = CType(projectboardWindows(PTwindows.mpt).ActiveSheet, Excel.Worksheet).Name
        If vglName <> arrWsNames(ptTables.MPT) Then
            Call MsgBox("Window 0 zeigt auf das falsche Sheet: " & vglName)
            Exit Sub
        End If

        '' '' alle Windows schliessen, bis auf das MPT Window 
        ' ''For Each tmpWindow In CType(appInstance.Workbooks.Item(myProjektTafel), Excel.Workbook).Windows
        ' ''    tmpWindow.Activate()

        ' ''    If CType(tmpWindow.ActiveSheet, Excel.Worksheet).Name = vglName Then
        ' ''        ' nichts tun ...
        ' ''    Else

        ' ''        tmpWindow.Close(SaveChanges:=False)
        ' ''    End If
        ' ''Next

        If appInstance.ActiveWindow.WindowState = Excel.XlWindowState.xlMaximized Then
            appInstance.ActiveWindow.WindowState = Excel.XlWindowState.xlNormal
        End If

        Try
            ' jetzt werden die Windows gelöscht, falls sie überhaupt existieren  ...
            If Not IsNothing(projectboardWindows(PTwindows.massEdit)) Then
                Try
                    projectboardWindows(PTwindows.massEdit).Close()
                Catch ex As Exception

                End Try

                projectboardWindows(PTwindows.massEdit) = Nothing
            End If

            If Not IsNothing(projectboardWindows(PTwindows.meChart)) Then
                Try
                    projectboardWindows(PTwindows.meChart).Close()
                Catch ex As Exception

                End Try

                projectboardWindows(PTwindows.meChart) = Nothing
            End If
            If Not IsNothing(projectboardWindows(PTwindows.mptpf)) Then

                'projectboardWindows(PTwindows.mptpf).Activate()
                If appInstance.ActiveWindow.WindowState = Excel.XlWindowState.xlMaximized Then
                    appInstance.ActiveWindow.WindowState = Excel.XlWindowState.xlNormal
                End If
                Try
                    projectboardWindows(PTwindows.mptpf).Close()
                Catch ex As Exception

                End Try
                projectboardWindows(PTwindows.mptpf) = Nothing
            End If
            If Not IsNothing(projectboardWindows(PTwindows.mptpr)) Then
                Try
                    projectboardWindows(PTwindows.mptpr).Close()
                Catch ex As Exception

                End Try

                projectboardWindows(PTwindows.mptpr) = Nothing
            End If
        Catch ex As Exception

            ' '' make MPT Window great again ...
            ''With projectboardWindows(PTwindows.mpt)
            ''    .Visible = True
            ''    .WindowState = XlWindowState.xlMaximized
            ''End With

        End Try


        ' tk / ute jetzt eigentlich überflüssig
        ' jetzt die projectboardWindows = Nothing setzen 
        'projectboardWindows(PTwindows.massEdit) = Nothing
        'projectboardWindows(PTwindows.meChart) = Nothing
        'projectboardWindows(PTwindows.mptpf) = Nothing
        'projectboardWindows(PTwindows.mptpr) = Nothing

        ' make MPT Window great again ...
        With projectboardWindows(PTwindows.mpt)
            .Visible = True
            .WindowState = XlWindowState.xlMaximized
        End With


    End Sub

    ''' <summary>
    ''' bestimmt in Abhängigkeit von TableTyp die Größe und Position des Fensters
    ''' je nachdem, wieviele LegendenEinträge das Chart hat, wird die Höhe etws höher bestimmt ... 
    ''' </summary>
    ''' <param name="tableTyp"></param>
    ''' <param name="chtop"></param>
    ''' <param name="chleft"></param>
    ''' <param name="chwidth"></param>
    ''' <param name="chHeight"></param>
    ''' <remarks></remarks>
    Public Sub bestimmeChartPositionAndSize(ByVal tableTyp As Integer,
                                            ByVal anzLegendEintraege As Integer,
                                                ByRef chtop As Double,
                                                ByRef chleft As Double,
                                                ByRef chwidth As Double,
                                                ByRef chHeight As Double)

        Dim currentWorksheet As Excel.Worksheet =
            CType(CType(appInstance.Workbooks.Item(myProjektTafel), Excel.Workbook).Worksheets.Item(arrWsNames(tableTyp)), Excel.Worksheet)

        Dim tmpTop As Double = 2.0
        Dim tmpLeft As Double = 2

        Dim korrfaktorH1 As Double = 1.2
        Dim korrfaktorH2 As Double = 1.0
        Dim korrfaktorB As Double = 1.0

        Try
            If My.Computer.Screen.Bounds.Height < 1080 Then
                korrfaktorH1 = 1.66
            End If

            If anzLegendEintraege > 2 Then
                korrfaktorH2 = 1 + 0.15 * (CInt(anzLegendEintraege / 2) - 1)
            End If
            'korrfaktorH = 1080 / My.Computer.Screen.Bounds.Height
            'korrfaktorB = 1920 / My.Computer.Screen.Bounds.Width
        Catch ex As Exception

        End Try
        'Dim tmpWidth As Double = maxScreenWidth / 5 - 29
        Dim tmpWidth As Double = chartWidth - 10
        'Dim tmpHeight As Double = (maxScreenHeight - 39) / 5 * korrfaktorH1 * korrfaktorH2
        Dim tmpHeight As Double = chartHeight * korrfaktorH1 * korrfaktorH2


        ' wenn schon Charts existieren: ein neues Chart wird immer als letztes  angehängt ..
        With currentWorksheet
            For Each tmpChtObject As Excel.ChartObject In CType(.ChartObjects, Excel.ChartObjects)
                If tmpChtObject.Top + tmpChtObject.Height + 2 > tmpTop Then
                    tmpTop = tmpChtObject.Top + tmpChtObject.Height + 2
                End If
            Next
        End With

        ' jetzt die Werte setzen 
        chtop = tmpTop
        chleft = tmpLeft
        chwidth = tmpWidth
        chHeight = tmpHeight

    End Sub

    ''' <summary>
    ''' definiert die Windows und Views, die benötigt werden 
    ''' es ist die Tabelle1=mpt aktiviert 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub defineVisboWindowViews()

        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim formerSU As Boolean = appInstance.ScreenUpdating
        Dim formereOU As Boolean = enableOnUpdate

        If enableOnUpdate Then
            enableOnUpdate = False
        End If

        If appInstance.EnableEvents Then
            appInstance.EnableEvents = False
        End If

        If appInstance.ScreenUpdating Then
            appInstance.ScreenUpdating = False
        End If

        ' jetzt werden die Windows aufgebaut ...

        ' dann werden alle auf invisible gesetzt , bis auf projectboardWindows(mpt)


        Dim visboWorkbook As Excel.Workbook = appInstance.Workbooks.Item(myProjektTafel)


        'projectboardWindows(PTwindows.mpt) = appInstance.ActiveWindow.NewWindow
        projectboardWindows(PTwindows.mpt) = appInstance.ActiveWindow


        ' Aus dem aktuellen Window ein benanntes Window machen 

        projectboardWindows(PTwindows.mptpr) = appInstance.ActiveWindow.NewWindow

        ' jetzt auf das Worksheet positionieren ...
        CType(visboWorkbook.Worksheets(arrWsNames(ptTables.mptPrCharts)), Excel.Worksheet).Activate()

        With projectboardWindows(PTwindows.mptpr)
            .WindowState = Excel.XlWindowState.xlNormal
            .EnableResize = True
            .DisplayHorizontalScrollBar = True
            .DisplayVerticalScrollBar = True
            .DisplayFormulas = False
            .DisplayHeadings = False
            .DisplayGridlines = False
            .GridlineColor = RGB(255, 255, 255)
            .DisplayWorkbookTabs = False
            .Caption = bestimmeWindowCaption(PTwindows.mptpr)
            .Visible = False
        End With

        ' Aufbau des Windows windowNames(4): Charts
        projectboardWindows(PTwindows.mptpf) = appInstance.ActiveWindow.NewWindow

        ' jetzt das Worksheet aktivieren ...
        visboWorkbook.Worksheets.Item(arrWsNames(ptTables.mptPfCharts)).activate()

        With projectboardWindows(PTwindows.mptpf)
            .WindowState = Excel.XlWindowState.xlNormal
            .EnableResize = True
            .DisplayHorizontalScrollBar = True
            .DisplayVerticalScrollBar = True
            .DisplayGridlines = False
            .DisplayHeadings = False
            .DisplayRuler = False
            .DisplayOutline = False
            .DisplayWorkbookTabs = False
            .Caption = bestimmeWindowCaption(PTwindows.mptpf)
            .Visible = False
        End With


        ' jetzt das Sheet Multiprojekt-Tafel aktivieren
        visboWorkbook.Worksheets.Item(arrWsNames(ptTables.MPT)).activate()

        'jetzt das MPT Sheet wieder holen 
        With projectboardWindows(PTwindows.mpt)
            .WindowState = XlWindowState.xlMaximized
            .Activate()
        End With

        ' wieder auf den Ausgangszustand setzen ... 
        With appInstance
            If .EnableEvents <> formerEE Then
                .EnableEvents = formerEE
            End If

            If .ScreenUpdating <> formerSU Then
                .ScreenUpdating = formerSU
            End If

            If enableOnUpdate <> formereOU Then
                enableOnUpdate = formereOU
            End If
        End With


    End Sub
    ''' <summary>
    ''' zeigt das angegebene VISBO Window, wenn es nicht ohnehin schon angezeigt wird ...
    ''' tmpmsg ist der optional Ergänzungs-String für den Caption Text im mptpr Window 
    ''' </summary>
    ''' <param name="visboWindowType"></param>
    ''' <remarks></remarks>
    Public Sub showVisboWindow(ByVal visboWindowType As Integer, Optional tmpmsg As String = "")


        ' Voraussetzungen schaffen: kein EnableEvents und kein Flackern und kein EnableOnUpdate ..
        '
        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim formerSU As Boolean = appInstance.ScreenUpdating
        Dim formereOU As Boolean = enableOnUpdate

        'Dim stdPfPrWindowBreite As Double = maxScreenWidth / 5 - 10
        Dim stdPfPrWindowBreite As Double = chartWidth + 10

        If enableOnUpdate Then
            enableOnUpdate = False
        End If

        If appInstance.EnableEvents Then
            appInstance.EnableEvents = False
        End If

        If appInstance.ScreenUpdating Then
            appInstance.ScreenUpdating = False
        End If


        ' Ende Voraussetzungen schaffen 
        Dim pfWindowAlreadyExisting As Boolean = False
        Dim prWindowAlreadyExisting As Boolean = False
        Dim foundWindow As Excel.Window = Nothing
        Dim visboWorkbook As Excel.Workbook = appInstance.Workbooks.Item(myProjektTafel)

        ' 
        For Each tmpWindow As Excel.Window In visboWorkbook.Windows

            If CType(tmpWindow.ActiveSheet, Excel.Worksheet).Name = arrWsNames(ptTables.mptPfCharts) Then

                pfWindowAlreadyExisting = True
                foundWindow = tmpWindow

            End If

            If CType(tmpWindow.ActiveSheet, Excel.Worksheet).Name = arrWsNames(ptTables.mptPrCharts) Then

                prWindowAlreadyExisting = True
                foundWindow = tmpWindow

            End If

        Next



        Select Case visboWindowType
            Case PTwindows.mptpf
                Try
                    If Not pfWindowAlreadyExisting Then

                        ' Aufbau des Windows PTwindows.mptpf Charts
                        'projectboardWindows(PTwindows.mptpf) = appInstance.ActiveWindow.NewWindow
                        projectboardWindows(PTwindows.mptpf) = projectboardWindows(PTwindows.mpt).NewWindow

                        ' jetzt das Worksheet aktivieren ... dazu muss aber wahrscheinlich appinstance.EnableEvents = true sein ? 
                        appInstance.EnableEvents = True
                        CType(visboWorkbook.Worksheets.Item(arrWsNames(ptTables.mptPfCharts)), Excel.Worksheet).Activate()
                        appInstance.EnableEvents = False

                        ' nur wenn ein neues erzeugt wurde, ist das Arragieren notwendig , damit die Größe verändert werdne kann 
                        If Not prWindowAlreadyExisting Then
                            appInstance.Windows.Arrange(Excel.XlArrangeStyle.xlArrangeStyleVertical)
                        End If


                        ' jetzt soll die Größe entsprechend eingestellt werden ..

                        With projectboardWindows(PTwindows.mptpf)
                            .Visible = True
                            .WindowState = Excel.XlWindowState.xlNormal
                            .EnableResize = True
                            '.Left = 4 * maxScreenWidth / 5
                            .Left = maxScreenWidth - stdPfPrWindowBreite
                            .Width = stdPfPrWindowBreite
                            ' wenn prWindows schon existert hat ..
                            If prWindowAlreadyExisting Then
                                .Top = projectboardWindows(PTwindows.mpt).Top
                                .Height = projectboardWindows(PTwindows.mpt).Height
                            End If
                        End With

                    Else
                        projectboardWindows(PTwindows.mptpf) = foundWindow
                        With projectboardWindows(PTwindows.mptpf)
                            .Visible = True
                            .WindowState = Excel.XlWindowState.xlNormal
                            .EnableResize = True
                        End With
                    End If

                    ' soll in allen Fällen gemacht werden 
                    With projectboardWindows(PTwindows.mptpf)
                        .DisplayHorizontalScrollBar = True
                        .DisplayVerticalScrollBar = True
                        .DisplayGridlines = False
                        .DisplayHeadings = False
                        .DisplayRuler = False
                        .DisplayOutline = False
                        .DisplayWorkbookTabs = False
                        .Caption = bestimmeWindowCaption(PTwindows.mptpf)
                    End With

                    ' jetzt muss das mpt Window in der Größe verändert werden, aber nur , wenn das nicht schon vorher existiert hat
                    If Not pfWindowAlreadyExisting Then

                        With projectboardWindows(PTwindows.mpt)
                            If .WindowState = Excel.XlWindowState.xlMaximized Then
                                .WindowState = Excel.XlWindowState.xlNormal
                            End If

                            If prWindowAlreadyExisting Then
                                .Left = 1 + stdPfPrWindowBreite + 1
                                .Width = projectboardWindows(PTwindows.mptpf).Left - 1 - .Left
                            Else
                                .Left = 1
                                .Width = projectboardWindows(PTwindows.mptpf).Left - 1
                            End If


                        End With

                        pfWindowAlreadyExisting = True

                    End If

                Catch ex As Exception

                End Try
            Case PTwindows.mptpr
                Try
                    If Not prWindowAlreadyExisting Then

                        ' Aufbau des Windows PTwindows.mptpr Charts
                        projectboardWindows(PTwindows.mptpr) = projectboardWindows(PTwindows.mpt).NewWindow

                        ' jetzt das Worksheet aktivieren ... dazu muss aber wahrscheinlich appinstance.EnableEvents = true sein ? 
                        appInstance.EnableEvents = True
                        CType(visboWorkbook.Worksheets.Item(arrWsNames(ptTables.mptPrCharts)), Excel.Worksheet).Activate()
                        appInstance.EnableEvents = False

                        ' nur wenn ein neues erzeugt wurde, ist das Arragieren notwendig , damit die Größe verändert werden kann 
                        If Not pfWindowAlreadyExisting Then
                            appInstance.Windows.Arrange(Excel.XlArrangeStyle.xlArrangeStyleVertical)
                        End If


                        ' jetzt soll die Größe entsprechend eingestellt werden ..

                        With projectboardWindows(PTwindows.mptpr)
                            .Visible = True
                            .WindowState = Excel.XlWindowState.xlNormal
                            .EnableResize = True
                            .Left = 1
                            .Width = stdPfPrWindowBreite
                            ' wenn pfWindows schon existert hat ..
                            If pfWindowAlreadyExisting Then
                                .Top = projectboardWindows(PTwindows.mpt).Top
                                .Height = projectboardWindows(PTwindows.mpt).Height
                            End If
                        End With

                    Else
                        projectboardWindows(PTwindows.mptpr) = foundWindow
                        With projectboardWindows(PTwindows.mptpr)
                            .Visible = True
                            .WindowState = Excel.XlWindowState.xlNormal
                            .EnableResize = True
                        End With
                    End If

                    ' soll in allen Fällen gemacht werden 
                    With projectboardWindows(PTwindows.mptpr)
                        .DisplayHorizontalScrollBar = True
                        .DisplayVerticalScrollBar = True
                        .DisplayGridlines = False
                        .DisplayHeadings = False
                        .DisplayRuler = False
                        .DisplayOutline = False
                        .DisplayWorkbookTabs = False
                        .Caption = bestimmeWindowCaption(PTwindows.mptpr, addOnMsg:=tmpmsg)
                    End With

                    If Not prWindowAlreadyExisting Then

                        With projectboardWindows(PTwindows.mpt)
                            If .WindowState = Excel.XlWindowState.xlMaximized Then
                                .WindowState = Excel.XlWindowState.xlNormal
                            End If

                            .Left = projectboardWindows(PTwindows.mptpr).Left +
                                    projectboardWindows(PTwindows.mptpr).Width + 1

                            If pfWindowAlreadyExisting Then
                                .Width = projectboardWindows(PTwindows.mptpf).Left - 1 - .Left
                            Else
                                .Width = maxScreenWidth - (projectboardWindows(PTwindows.mptpr).Width + 1)
                            End If



                        End With

                        prWindowAlreadyExisting = True

                    End If


                Catch ex As Exception

                End Try


            Case PTwindows.mpt
                projectboardWindows(PTwindows.mpt) = foundWindow
                With projectboardWindows(PTwindows.mpt)
                    .Visible = True
                    .WindowState = Excel.XlWindowState.xlNormal
                    .EnableResize = True
                End With
            Case Else
                ' nichts tun 
        End Select


        ' alten Zustand bezgl enableEvents etc wieder herstellen ...
        ' wieder auf den Ausgangszustand setzen ... 
        '

        ' jetzt wieder auf das Haupt-Window positionieren 
        projectboardWindows(PTwindows.mpt).Activate()


        With appInstance
            If .EnableEvents <> formerEE Then
                .EnableEvents = formerEE
            End If

            If .ScreenUpdating <> formerSU Then
                .ScreenUpdating = formerSU
            End If

            If enableOnUpdate <> formereOU Then
                enableOnUpdate = formereOU
            End If
        End With

    End Sub




    ''' <summary>
    ''' bestimmt den rollen-ID-String in der Form: roleUid;teamUid
    ''' </summary>
    ''' <param name="currentCell"></param>
    ''' <returns></returns>
    Public Function getRCNameIDfromExcelCell(ByVal currentCell As Excel.Range,
                                             Optional ByVal returnOnlyValidNameID As Boolean = False) As String

        Dim tmpResult As String = ""
        Try
            If Not IsNothing(currentCell.Value) Then
                Dim tmpRCname As String = CStr(currentCell.Value).Trim

                If tmpRCname <> "" Then
                    If RoleDefinitions.containsName(tmpRCname) Then
                        Dim tmpComment As Excel.Comment = currentCell.Comment
                        Dim tmpTeamName As String = ""
                        If Not IsNothing(tmpComment) Then
                            tmpTeamName = tmpComment.Text
                        End If
                        tmpResult = RoleDefinitions.bestimmeRoleNameID(tmpRCname, tmpTeamName)
                    Else
                        ' im Falle von Kosten soll erst mal die alte Herangehensweise gelten
                        If returnOnlyValidNameID Then
                            tmpResult = ""
                        Else
                            tmpResult = tmpRCname
                        End If

                    End If
                End If

            Else
                tmpResult = ""
            End If

        Catch ex As Exception
            tmpResult = ""
        End Try

        getRCNameIDfromExcelCell = tmpResult

    End Function



    ''' <summary>
    ''' ermittelt in einem Excel Zelle die PhaseNameID  
    ''' </summary>
    ''' <param name="currentCell"></param>
    ''' <returns></returns>
    Public Function getPhaseNameIDfromExcelCell(ByVal currentCell As Excel.Range) As String
        Dim tmpResult As String = ""

        Try

            If Not IsNothing(currentCell.Value) Then
                Dim phaseName As String = CStr(currentCell.Value).Trim
                Dim phaseNameID As String = calcHryElemKey(phaseName, False)
                Dim curComment As Excel.Comment = currentCell.Comment

                If Not IsNothing(curComment) Then
                    phaseNameID = curComment.Text.Trim
                End If

                tmpResult = phaseNameID
            End If

        Catch ex As Exception
            tmpResult = ""
        End Try


        getPhaseNameIDfromExcelCell = tmpResult
    End Function





    ''' <summary>
    ''' prüft ob in dem aktiven Massen-Edit Sheet die übergebene Kombination nocheinmal vorkommt ... 
    ''' wenn nein: Rückgabe true
    ''' wenn ja: Rückgabe false
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <param name="phaseNameID"></param>
    ''' <param name="rcNameID"></param>
    ''' <param name="zeile"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function noDuplicatesInSheet(ByVal pName As String, ByVal phaseNameID As String, ByVal rcNameID As String,
                                             ByVal zeile As Integer) As Boolean
        Dim found As Boolean = False
        Dim curZeile As Integer = 2

        Dim chckName As String
        Dim chckPhNameID As String
        Dim chckRCNameID As String

        Dim teamID As Integer = -1
        Dim isRole As Boolean = Not IsNothing(RoleDefinitions.getRoleDefByIDKennung(rcNameID, teamID))

        Dim meWS As Excel.Worksheet = CType(CType(appInstance.Workbooks(myProjektTafel), Excel.Workbook) _
            .Worksheets(arrWsNames(ptTables.meRC)), Excel.Worksheet)

        With meWS
            chckName = CStr(meWS.Cells(curZeile, 2).value)

            'Dim phaseName As String = CStr(meWS.Cells(curZeile, 4).value)
            chckPhNameID = getPhaseNameIDfromExcelCell(CType(meWS.Cells(curZeile, 4), Excel.Range))

            'chckPhNameID = calcHryElemKey(phaseName, False)
            'Dim curComment As Excel.Comment = CType(meWS.Cells(curZeile, 4), Excel.Range).Comment
            'If Not IsNothing(curComment) Then
            '    chckPhNameID = curComment.Text
            'End If
            If Not isRole Then
                chckRCNameID = CStr(meWS.Cells(curZeile, 5).value)
            Else
                chckRCNameID = getRCNameIDfromExcelCell(CType(meWS.Cells(curZeile, 5), Excel.Range))
            End If


        End With
        ' aus der Funktionalität zeile löschen wird rcName auch mit Nothing aufgerufen ... 
        Do While Not found And curZeile <= visboZustaende.meMaxZeile


            If chckName = pName And
                phaseNameID = chckPhNameID And
                zeile <> curZeile Then

                If IsNothing(rcNameID) Then
                    found = True
                ElseIf rcNameID = chckRCNameID Then
                    found = True
                End If

            End If

            If Not found Then

                curZeile = curZeile + 1

                With meWS
                    chckName = CStr(meWS.Cells(curZeile, 2).value)

                    'Dim phaseName As String = CStr(meWS.Cells(curZeile, 4).value)
                    'chckPhNameID = calcHryElemKey(phaseName, False)
                    chckPhNameID = getPhaseNameIDfromExcelCell(CType(meWS.Cells(curZeile, 4), Excel.Range))
                    'Dim curComment As Excel.Comment = CType(meWS.Cells(curZeile, 4), Excel.Range).Comment
                    'If Not IsNothing(curComment) Then
                    '    chckPhNameID = curComment.Text
                    'End If


                    If Not isRole Then
                        chckRCNameID = CStr(meWS.Cells(curZeile, 5).value)
                    Else
                        chckRCNameID = getRCNameIDfromExcelCell(CType(meWS.Cells(curZeile, 5), Excel.Range))
                    End If


                End With

            End If

        Loop

        noDuplicatesInSheet = Not found

    End Function

    ''' <summary>
    ''' faerbt die Projekt-Karte entsprechend der 
    ''' </summary>
    ''' <param name="pptShape"></param>
    ''' <param name="colorindex"></param>
    Public Sub faerbeProjectCard(ByRef pptShape As PowerPoint.Shape, ByVal colorindex As Integer)

        If Not IsNothing(pptShape) Then

            With pptShape
                .Shadow.Type = MsoShadowType.msoShadow25
                .Shadow.Visible = MsoTriState.msoTrue
                .Shadow.Style = MsoShadowStyle.msoShadowStyleOuterShadow
                .Shadow.Blur = 4
                .Shadow.OffsetX = CInt(.Width / 7)
                .Shadow.OffsetY = CInt(.Width / 7)
                .Shadow.Transparency = 0
                .Shadow.Size = 110
                .Shadow.RotateWithShape = MsoTriState.msoFalse

                If colorindex = 0 Then
                    .Shadow.ForeColor.RGB = visboFarbeNone
                    '.Line.ForeColor.RGB = visboFarbeNone
                ElseIf colorindex = 1 Then
                    .Shadow.ForeColor.RGB = visboFarbeGreen
                    '.Line.ForeColor.RGB = visboFarbeGreen
                ElseIf colorindex = 2 Then
                    .Shadow.ForeColor.RGB = visboFarbeYellow
                    '.Line.ForeColor.RGB = visboFarbeYellow
                ElseIf colorindex = 3 Then
                    .Shadow.ForeColor.RGB = visboFarbeRed
                    '.Line.ForeColor.RGB = visboFarbeRed
                End If

            End With
        End If


    End Sub

    ''' <summary>
    ''' prüft, ob es sich bei dem übergebenen Rollen-Namen um einen validen Rollen-Namen handelt
    ''' </summary>
    ''' <param name="roleName"></param>
    ''' <param name="errMsg"></param>
    ''' <returns></returns>
    Public Function isValidRoleName(ByVal roleName As String, ByRef errMsg As String) As Boolean
        Dim isvalid As Boolean = False
        errMsg = ""

        If roleName.Contains(";") Then
            isvalid = False
            errMsg = "Rollen-Namen dürfen keine ';' enthalten : " & roleName
        Else
            isvalid = True
        End If

        isValidRoleName = isvalid
    End Function

    ''' <summary>
    ''' liefert true, wenn diese URL erreichbar ist, false andernfalls
    ''' </summary>
    ''' <param name="URL"></param>
    ''' <returns></returns>
    Public Function isValidURL(ByVal URL As String) As Boolean
        Try
            Dim Response As Net.WebResponse = Nothing
            Dim WebReq As Net.HttpWebRequest = CType(Net.HttpWebRequest.Create(URL), Net.HttpWebRequest)
            Response = WebReq.GetResponse
            Response.Close()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Test-Funktion: überprüft die Team-Definitionen 
    ''' </summary>
    ''' <returns></returns>
    Public Function checkTeamDefinitions(ByVal roleDefinitionsToCheck As clsRollen, ByRef outputCollection As Collection) As Boolean

        Dim allTeams As SortedList(Of Integer, Double) = roleDefinitionsToCheck.getAllTeamIDs
        Dim atleastOneError As Boolean = False

        For Each kvp As KeyValuePair(Of Integer, Double) In allTeams

            Dim ok As Boolean = True
            Dim teamRole As clsRollenDefinition = roleDefinitionsToCheck.getRoleDefByID(kvp.Key)
            Dim childIDs As SortedList(Of Integer, Double) = teamRole.getSubRoleIDs

            For Each child As KeyValuePair(Of Integer, Double) In childIDs
                Dim childRole As clsRollenDefinition = roleDefinitionsToCheck.getRoleDefByID(child.Key)
                ok = ok And childRole.getTeamIDs.ContainsKey(kvp.Key)
                If Not ok Then
                    Dim outmsg As String = "teamRole " & teamRole.name & " conflicts with " & childRole.name
                    outputCollection.Add(outmsg)
                    atleastOneError = True
                    ok = True
                End If
            Next

        Next
        checkTeamDefinitions = atleastOneError

    End Function

    ''' <summary>
    ''' Test-Funktion für Teams und Überlastung 
    ''' </summary>
    ''' <returns></returns>
    Public Function checkTeamMemberOverloads(ByVal roledefinitionsToCheck As clsRollen, ByRef outputCollection As Collection) As Boolean

        Dim allIDs As SortedList(Of Integer, Double) = roledefinitionsToCheck.getAllIDs
        Dim atleastOneOverload As Boolean = False

        For Each kvp As KeyValuePair(Of Integer, Double) In allIDs

            Dim tmpRole As clsRollenDefinition = roledefinitionsToCheck.getRoleDefByID(kvp.Key)
            Dim memberships As SortedList(Of Integer, Double) = tmpRole.getTeamIDs

            If Not IsNothing(memberships) Then
                If memberships.Count > 0 Then
                    Dim wholeKapa As Double = 0.0
                    For Each membership As KeyValuePair(Of Integer, Double) In memberships
                        wholeKapa = wholeKapa + membership.Value
                    Next

                    If wholeKapa > 1.0 Then
                        atleastOneOverload = True
                        Dim errmsg As String = "Overloaded Role: " & tmpRole.name & "Kapa: " & wholeKapa.ToString("#0.#")
                        outputCollection.Add(errmsg)
                    End If
                End If
            End If

        Next

        checkTeamMemberOverloads = atleastOneOverload

    End Function



    Public Function bestimmeChartYValues(ByVal curProj As clsProjekt, ByVal chartTyp As Integer, ByVal pmrcName As String, ByVal von As Integer, ByVal bis As Integer) As Double()
        bestimmeChartYValues = Nothing
    End Function

    ''' <summary>
    ''' gibt den Titel zurück 
    ''' </summary>
    ''' <param name="rollenKennung"></param>
    ''' <returns></returns>
    Public Function bestimmeRollenDiagrammTitel(ByVal rollenKennung As String) As String

        Dim tmpResult As String = ""

        Dim teamID As Integer
        Try
            tmpResult = RoleDefinitions.getRoleDefByIDKennung(rollenKennung, teamID).name
        Catch ex As Exception

        End Try

        bestimmeRollenDiagrammTitel = tmpResult

    End Function

    ''' <summary>
    ''' gibt für eine sortierte String-Collection und eine sortierte Liste of string, Double die matching list als sortedList of String, double raus 
    ''' </summary>
    ''' <param name="existing"></param>
    ''' <param name="lookingFor"></param>
    ''' <returns></returns>
    Public Function intersectNameIDLists(ByVal existing As Collection,
                                   ByVal lookingFor As SortedList(Of String, Double)) As SortedList(Of String, Double)

        Dim ergebnisListe As New SortedList(Of String, Double)
        Dim teamID As Integer
        Dim roleID As Integer


        Try
            If existing.Count <= lookingFor.Count Then

                For Each key As String In existing

                    If lookingFor.ContainsKey(key) Then
                        ergebnisListe.Add(key, lookingFor.Item(key))
                    End If


                    roleID = RoleDefinitions.parseRoleNameID(key, teamID)
                    If teamID = -1 Then
                        ' fertig 
                    Else
                        ' es muss noch die Anfrage nach nur roleID gestellt werden 
                        Dim key2 As String = RoleDefinitions.bestimmeRoleNameID(roleID, -1)
                        If lookingFor.ContainsKey(key2) Then
                            ergebnisListe.Add(key, lookingFor.Item(key2))
                        End If
                    End If
                Next

            Else
                ' Vorbereitung , mit der indexliste kann in existing schnell nach allen RoleIDs gesucht werden, die vorkommen
                Dim indexListe As New SortedList(Of Integer, Integer())

                ' das muss gemacht werden, weil man sonst nicht auf eine sortierte liste auch über den index zugreifen kann 
                Dim existingSortList As New SortedList(Of String, Double)

                For Each nameID As String In existing
                    If Not existingSortList.ContainsKey(nameID) Then
                        existingSortList.Add(nameID, 1.0)
                    End If
                Next

                Dim oldRoleID As Integer = -1

                ' mit indexListe wird eine Hilfs-Struktur aufgebaut, die den Umstand nutzt, dass die NAmeIDs alle sortiert sind und 
                ' deshalb Rollen mit gleicher RoleID beieinander stehen; deswegen muss auch eine Rolle ohne team mit';' enden 
                For ix As Integer = 0 To existingSortList.Count - 1
                    Dim nameID As String = existingSortList.ElementAt(ix).Key
                    roleID = RoleDefinitions.parseRoleNameID(nameID, teamID)

                    If roleID > 0 Then

                        If roleID <> oldRoleID Then
                            If oldRoleID = -1 Then
                                ' ein neuer start ist entdeckt 
                                Dim wertePaar() As Integer
                                ReDim wertePaar(1)
                                wertePaar(0) = ix
                                wertePaar(1) = ix
                                oldRoleID = roleID
                                indexListe.Add(roleID, wertePaar)

                            Else
                                ' ein neues Ende .. und ein neuer Start ist entdeckt 
                                indexListe.Item(oldRoleID)(1) = ix - 1
                                Dim wertePaar() As Integer
                                ReDim wertePaar(1)
                                wertePaar(0) = ix
                                wertePaar(1) = ix
                                oldRoleID = roleID
                                indexListe.Add(roleID, wertePaar)
                            End If

                        End If

                    End If

                Next

                For Each kvp As KeyValuePair(Of String, Double) In lookingFor

                    roleID = RoleDefinitions.parseRoleNameID(kvp.Key, teamID)


                    If teamID = -1 Then
                        ' aus existing alle bekommen, die mit roleID beginnen
                        If indexListe.ContainsKey(roleID) Then
                            For ix As Integer = indexListe.Item(roleID)(0) To indexListe.Item(roleID)(1)
                                Dim nameID As String = existingSortList.ElementAt(ix).Key
                                If Not ergebnisListe.ContainsKey(nameID) Then
                                    ergebnisListe.Add(nameID, 1.0)
                                End If
                            Next
                        End If
                    Else
                        ' es ist ein Team angegeben , also will man das exakte haben 
                        If existingSortList.ContainsKey(kvp.Key) Then
                            ergebnisListe.Add(kvp.Key, kvp.Value)
                        End If
                    End If

                Next

            End If
        Catch ex As Exception
            Call MsgBox("Fehler in intersectNAmeIDs")
        End Try



        intersectNameIDLists = ergebnisListe
    End Function


    ''' <summary>
    ''' ruft das Formular auf, um die Proxy-Authentifizierung zu erfragen
    ''' </summary>
    ''' <remarks></remarks>
    Public Function askProxyAuthentication(ByRef proxyURL As String, ByRef usr As String, ByRef pwd As String, ByRef domain As String) As Boolean
        Dim proxyAuth As New frmProxyAuth
        Dim returnValue As DialogResult = DialogResult.Retry
        Dim i As Integer = 0

        proxyAuth.proxyURL = proxyURL


        While returnValue <> DialogResult.OK And returnValue <> DialogResult.Cancel

            returnValue = proxyAuth.ShowDialog

        End While

        If returnValue = DialogResult.OK Then

            proxyURL = proxyAuth.proxyURL
            domain = proxyAuth.domain
            usr = proxyAuth.user
            pwd = proxyAuth.pwd
        Else

            askProxyAuthentication = True

        End If

        askProxyAuthentication = True

    End Function

    ''' <summary>
    ''' schreibt in die angegebene MassenEdit Excel-Zelle den Phase-Name als String, 
    ''' die Phase-NameID , wenn nötig als unsichtbaren Kommentar  
    ''' </summary>
    ''' <param name="currentCell"></param>
    ''' <param name="phaseNameID"></param>
    Public Sub writeMEcellWithPhaseNameID(ByRef currentCell As Excel.Range,
                                          ByVal indentlevel As Integer,
                                          ByVal phaseName As String,
                                          ByVal phaseNameID As String)
        ' Phasen-Name 
        currentCell.Value = phaseName
        '    Den Indent schreiben 
        currentCell.IndentLevel = indentlevel
        '    Kommentare alle löschen 
        currentCell.ClearComments()

        ' wenn nötig Kommentar schreiben mit phaseNameID , damit später die ID zweifelsfrei ermitelt werden kann 
        If calcHryElemKey(phaseName, False) <> phaseNameID Then
            currentCell.AddComment(Text:=phaseNameID)
            currentCell.Comment.Visible = False
        End If

    End Sub

    ''' <summary>
    ''' schreibt den Projekt-NAmen, evtl inkl MArkierung dass geschützt und dem Hinweis, wer es geschützt hat
    ''' </summary>
    ''' <param name="currentCell"></param>
    ''' <param name="isProtectedbyOthers"></param>
    ''' <param name="protectiontext"></param>
    Public Sub writeMEcellWithProjectName(ByRef currentCell As Excel.Range,
                                          ByVal pName As String,
                                          ByVal isProtectedbyOthers As Boolean,
                                          ByVal protectiontext As String)

        currentCell.Value = pName

        If isProtectedbyOthers Then

            If isProtectedbyOthers Then
                currentCell.Font.Color = awinSettings.protectedByOtherColor
            End If

            ' Kommentare löschen
            currentCell.ClearComments()

            currentCell.AddComment(Text:=protectiontext)
            currentCell.Comment.Visible = False

        End If


    End Sub

    ''' <summary>
    ''' schreibt in die angegebene MassenEdit Excel-Zelle den Rollen-Namen als String und trägt ggf einen Kommentar mit dem Team-NAmen ein.  
    ''' </summary>
    ''' <param name="currentCell"></param>
    ''' <param name="roleNameID"></param>
    Public Sub writeMECellWithRoleNameID(ByRef currentCell As Excel.Range,
                                         ByVal isLocked As Boolean,
                                         ByVal rcName As String,
                                         ByVal roleNameID As String,
                                         ByVal isRole As Boolean)


        Dim teamID As Integer = -1
        Dim teamName As String = ""

        ' erst mal alle Kommentare löschen 
        currentCell.ClearComments()

        If isRole Then
            If rcName = roleNameID Or roleNameID = "" Then
                ' nichts weiter tun ... rcName wird als Value geschrieben

            ElseIf roleNameID.Length > 0 Then

                If Not IsNothing(RoleDefinitions.getRoleDefByIDKennung(roleNameID, teamID)) Then
                    Dim teamRole As clsRollenDefinition = RoleDefinitions.getRoleDefByID(teamID)

                    If Not IsNothing(teamRole) Then
                        teamName = teamRole.name
                    End If
                End If

            End If

        Else
            ' nichts weiter tun ... rcName wird als Kosten-Name geschrieben

        End If

        ' Jetzt wird die Zelle geschrieben 

        With currentCell
            .Value = rcName
            .Locked = isLocked
            Try
                If Not IsNothing(.Validation) Then
                    .Validation.Delete()
                End If
            Catch ex As Exception

            End Try

            If teamName.Length > 0 Then
                Dim newComment As Excel.Comment = .AddComment(Text:=teamName)
            End If

        End With

    End Sub


End Module
