Imports Microsoft.Office.Interop.Excel

Public Class clsawinSettings
    ' Chart Settings 
    Public Property fontsizeTitle As Integer
    Public Property fontsizeLegend As Integer
    Public Property fontsizeItems As Integer
    Public Property CPfontsizeTitle As Integer
    Public Property CPfontsizeItems As Integer
    Public Property ChartHoehe1 As Double
    Public Property ChartHoehe2 As Double
    Public Property SollIstFarbeB As Long
    Public Property SollIstFarbeL As Long
    Public Property SollIstFarbeC As Long
    Public Property SollIstFarbeArea As Long
    Public Property timeSpanColor As Long
    Public Property missingDefinitionColor As Long
    Public Property showTimeSpanInPT As Boolean

    Public Property AmpelGruen As Long
    Public Property AmpelGelb As Long
    Public Property AmpelRot As Long
    Public Property AmpelNichtBewertet As Long

    Public Property glowColor As Long

    ' hier werden die Settings gesetzt  

    ' Settings für die Projekteingabe
    Public Property lastProjektTyp As String
    Public Property lastModulTyp As String
    Public Property isEndDate As Boolean
    Public Property tryBestFit As Boolean
    Public Property selDate As Date
    Public Property bestFit As clsBestFitObject

    ' Settings für Grundeinstellungen 
    Public Property nullDatum As Date
    Public Property kalenderStart As Date
    Public Property zeitEinheit As String
    Public Property kapaEinheit As String
    Public Property databaseName As String
    Public Property databaseURL As String
    Public Property globalPath As String
    Public Property awinPath As String
    Public Property zeilenhoehe1 As Double
    Public Property zeilenhoehe2 As Double
    Public Property spaltenbreite As Double
    Public Property offsetEinheit As String
    Public Property drawphases As Boolean
    Public Property applyFilter As Boolean
    ' bestimmt ob das Project als Balken dargestellt wird oder einfach als Linie 
    Public Property drawProjectLine As Boolean
    ' bestimmt, ob die Beschriftungen von Meilensteinen und Phasen auf der Projekt-Tafel angezeigt werden sollen
    Public Property showElementNames As Boolean

    ' sollen Meilensteine auch ausserhalb des Projekts liegen dürfen ? 
    Public Property milestoneFreeFloat As Boolean
    ' sollen Bedarfe automatisch in der Array Länge angepasst werden, wenn sich das Projekt verschiebt und in Folge die array Länge 
    ' nicht mehr ganz passt 
    Public Property autoCorrectBedarfe As Boolean

    ' sollen Bedarfe proportional zur Streckung oder Stauchung eines Projekt angepasst werden
    Public Property propAnpassRess As Boolean

    ' soll bei der Leistbarkeit der Phasen anteilig gerechnet werden oder drin = 1
    Public Property phasesProzentual As Boolean = False

    ' sollen die Werte der selektierten Projekte in PRC Summencharts angezeigt werden ? 
    Public Property showValuesOfSelected As Boolean

    ' sollen Shapes aus den Update Informations-Forms heraus erzeugt werden, wenn sie noch nicht da sind 
    Public Property createIfNotThere As Boolean

    ' soll der Original Name angeziegt werden 
    Public Property showOrigName As Boolean

    ' soll der Best-Name (Name mit kürzest-möglichem Breadcrumb um eindeutig zu sein 
    Public Property showBestName As Boolean

    ' Settings für die letzte User Selektion in der Tafel 
    Public Property selectedColumn As Integer
    Public Property selectedRow As Integer

    ' Settings für Import / Export
    Public Property EinzelRessExport As Integer
    ' Settings ob die fehlenden Phase- und Meilenstein-Namen in die Customization eingetragen werden sollen
    Public Property addMissingPhaseMilestoneDef As Boolean
    ' Setting, ob die NAmen eines Templates auf alle Fälle in die Phasedefinitions / MilestoneDefinitions aufgenommen werden soll oder nicht 
    Public Property alwaysAcceptTemplateNames As Boolean
    ' Setting, das regelt, ob unbekannte Namen by default in die Projekt-Struktur aufgenommen werden
    ' sie werden aber auf alle Fälle nicht (!) in die PhaseDefinitions aufgenommen; 
    ' Ausnahme: wenn es sich um ein Template handelt, und alwaysAcceptTemplateNames = true

    ' soll beim rxf-, BMW Excel Import das volle Protokoll ausgegeben werden 
    Public Property fullProtocol As Boolean

    ' im BMW Import Kontext wichtiges Settings
    Property importTyp As Integer

    ' steuert, ob Kinder, die  Duplikate von ihren Eltern sind, eliminiert werden sollen  
    ' Duplikat heisst: gleicher Name und gleicher Termin 
    Property eliminateDuplicates As Boolean

    Public Property importUnknownNames As Boolean
    ' wird beim Import verwendet; steuert, ob beim Import aus MS Project, RPLAN, Excel Files eindeutige Namen 
    ' für gleichlautende Geschwisternamen generiert werden  
    Public Property createUniqueSiblingNames As Boolean

    ' Settings für ToleranzKorridor TimeCost
    Public Property timeToleranzRel As Double
    Public Property timeToleranzAbs As Double

    Public Property costToleranzRel As Double
    Public Property costToleranzAbs As Double

    ' Settings für Multiprojekt-Sichten
    Public Property mppShowAllIfOne As Boolean
    Public Property mppShowMsDate As Boolean
    Public Property mppShowMsName As Boolean
    Public Property mppShowPhDate As Boolean
    Public Property mppShowPhName As Boolean
    Public Property mppShowAmpel As Boolean
    Public Property mppShowProjectLine As Boolean
    Public Property mppVertikalesRaster As Boolean
    Public Property mppShowLegend As Boolean
    Public Property mppFullyContained As Boolean
    Public Property mppSortiertDauer As Boolean
    Public Property mppOnePage As Boolean
    Public Property mppExtendedMode As Boolean
    Public Property mppShowHorizontals As Boolean
    Public Property mppUseAbbreviation As Boolean
    Public Property mppUseOriginalNames As Boolean
    Public Property mppKwInMilestone As Boolean


    ' Settings für Einzelprojekt-Reports
    ' tk 7.2.16 ist überflüssig
    'Public Property eppExtendedMode As Boolean

    ' Settings für Überprüfung, ob Formulare offen / aktiv sind 
    Public Property isHryNameFrmActive As Boolean

    ' Settings für Auswahl-Dialog 
    Public Property useHierarchy As Boolean



    Sub New()

        ' Chart Settings
        _fontsizeTitle = 14
        _fontsizeLegend = 10
        _fontsizeItems = 10
        _CPfontsizeTitle = 10
        _CPfontsizeItems = 8
        _ChartHoehe1 = 150.0
        _ChartHoehe2 = 220.0
        _SollIstFarbeB = RGB(80, 80, 80)
        _SollIstFarbeL = RGB(80, 160, 80)
        _SollIstFarbeC = RGB(80, 240, 80)
        _SollIstFarbeArea = RGB(200, 200, 200)
        _timeSpanColor = RGB(242, 242, 242)
        _missingDefinitionColor = XlRgbColor.rgbCoral
        _showTimeSpanInPT = True


        ' Projekteingabe Settings
        _lastProjektTyp = ""
        _lastModulTyp = ""
        _isEndDate = False
        _tryBestFit = False
        _selDate = Date.Now
        _bestFit = New clsBestFitObject

        ' Settings für Grundeinstellungen
        _nullDatum = #6/23/1914#
        _kalenderStart = #1/1/2012#
        _kapaEinheit = "PT"
        _zeitEinheit = "PM"
        _databaseName = ""
        _databaseURL = ""
        _awinPath = ""

        _selectedColumn = 1
        _offsetEinheit = "d"
        _milestoneFreeFloat = True
        _autoCorrectBedarfe = True
        _propAnpassRess = False
        _phasesProzentual = False
        _drawphases = False
        _showValuesOfSelected = False
        _applyFilter = False
        _createIfNotThere = False
        _showOrigName = False
        _showBestName = True
        _drawProjectLine = True
        _showElementNames = False

        ' Settings für Import / Export 
        _EinzelRessExport = 0
        _addMissingPhaseMilestoneDef = False
        _alwaysAcceptTemplateNames = False
        _fullProtocol = False
        _importTyp = 1

        _eliminateDuplicates = True
        _importUnknownNames = True
        _createUniqueSiblingNames = True

        ' Settings für Besser/Schlechter Diagramm 
        _timeToleranzRel = 0.02
        _timeToleranzAbs = 3
        _costToleranzRel = 0.02
        _costToleranzAbs = 2

        ' Settings für Multiprojekt Sichten 
        _mppShowAllIfOne = False
        _mppShowMsDate = True
        _mppShowMsName = True
        _mppShowPhDate = True
        _mppShowPhName = True
        _mppShowAmpel = False
        _mppShowProjectLine = True
        _mppVertikalesRaster = False
        _mppShowLegend = False
        _mppFullyContained = False
        _mppSortiertDauer = False
        _mppOnePage = False
        _mppExtendedMode = False
        _mppShowHorizontals = False
        _mppUseAbbreviation = True
        _mppUseOriginalNames = False
        _mppKwInMilestone = False


        ' Settings für Einzelprojekt-Reports
        '_eppExtendedMode = True


        If _mppSortiertDauer Then
            _mppShowAllIfOne = True
        End If

        _useHierarchy = True
        _isHryNameFrmActive = False



    End Sub
End Class
