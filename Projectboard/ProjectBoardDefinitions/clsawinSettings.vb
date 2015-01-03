Public Class clsawinSettings
    ' Chart Settings 
    Property fontsizeTitle As Integer
    Property fontsizeLegend As Integer
    Property fontsizeItems As Integer
    Property CPfontsizeTitle As Integer
    Property CPfontsizeItems As Integer
    Property ChartHoehe1 As Double
    Property ChartHoehe2 As Double
    Property SollIstFarbeB As Long
    Property SollIstFarbeL As Long
    Property SollIstFarbeC As Long
    Property SollIstFarbeArea As Long
    Property timeSpanColor As Long
    Property showTimeSpanInPT As Boolean

    Property AmpelGruen As Long
    Property AmpelGelb As Long
    Property AmpelRot As Long
    Property AmpelNichtBewertet As Long

    Property glowColor As Long

    ' Settings für die Projekteingabe
    Property lastProjektTyp As String
    Property isEndDate As Boolean
    Property tryBestFit As Boolean
    Property selDate As Date
    Property bestFit As clsBestFitObject

    ' Settings für Grundeinstellungen 
    Property nullDatum As Date
    Property kalenderStart As Date
    Property zeitEinheit As String
    Property kapaEinheit As String
    Property databaseName As String
    Property zeilenhoehe1 As Double
    Property zeilenhoehe2 As Double
    Property spaltenbreite As Double
    Property offsetEinheit As String
    Property drawphases As Boolean
    Property loadProjectsOnChange As Boolean
    ' sollen Meilensteine auch ausserhalb des Projekts liegen dürfen ? 
    Property milestoneFreeFloat As Boolean
    ' sollen Bedarfe automatisch in der Array Länge angepasst werden, wenn sich das Projekt verschiebt und in Folge die array Länge 
    ' nicht mehr ganz passt 
    Property autoCorrectBedarfe As Boolean

    ' sollen Bedarfe proportional zur Streckung oder Stauchung eines Projekt angepasst werden
    Property propAnpassRess As Boolean

    ' soll bei der Leistbarkeit der Phasen anteilig gerechnet werden oder drin = 1
    Property phasesProzentual As Boolean = False

    ' sollen die Werte der selektierten Projekte in PRC Summencharts angezeigt werden ? 
    Property showValuesOfSelected As Boolean

    ' sollen Shapes aus den Update Informations-Forms heraus erzeugt werden, wenn sie noch nicht da sind 
    Property createIfNotThere As Boolean

    ' Settings für die letzte User Selektion in der Tafel 
    Property selectedColumn As Integer
    Property selectedRow As Integer

    ' Settings für Import / Export
    Property EinzelRessExport As Integer

    ' Settings für ToleranzKorridor TimeCost
    Property timeToleranzRel As Double
    Property timeToleranzAbs As Double

    Property costToleranzRel As Double
    Property costToleranzAbs As Double

    ' Settings für Multiprojekt-Sichten
    Property mppStrict As Boolean
    Property mppFullyContained As Boolean
    Property mppShowMsDate As Boolean
    Property mppShowMsName As Boolean
    Property mppShowPhDate As Boolean
    Property mppShowPhName As Boolean
    Property mppShowAmpel As Boolean
    Property mppShowProjectLine As Boolean
    Property mppVertikalesRaster As Boolean
    Property mppShowLegend As Boolean


    Property importTyp As Integer

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
        _showTimeSpanInPT = True


        ' Projekteingabe Settings
        _lastProjektTyp = ""
        _isEndDate = False
        _tryBestFit = False
        _selDate = Date.Now
        _bestFit = New clsBestFitObject

        ' Settings für Grundeinstellungen
        _nullDatum = #6/23/1914#
        _kalenderStart = #1/1/2012#
        _kapaEinheit = "PT"
        _zeitEinheit = "PM"
        _databaseName = "projectboard"
        _selectedColumn = 1
        _offsetEinheit = "d"
        _milestoneFreeFloat = False
        _autoCorrectBedarfe = True
        _propAnpassRess = False
        _phasesProzentual = False
        _drawphases = False
        _showValuesOfSelected = False
        _loadProjectsOnChange = False
        _createIfNotThere = True

        ' Settings für Import / Export 
        _EinzelRessExport = 0



        ' Settings für Besser/Schlechter Diagramm 
        _timeToleranzRel = 0.02
        _timeToleranzAbs = 3
        _costToleranzRel = 0.02
        _costToleranzAbs = 2

        ' Settings für Multiprojekt Sichten 
        _mppStrict = False
        _mppFullyContained = True
        _mppShowMsDate = True
        _mppShowMsName = True
        _mppShowPhDate = True
        _mppShowPhName = True
        _mppShowAmpel = False
        _mppShowProjectLine = True
        _mppVertikalesRaster = False
        _mppShowLegend = False
        _importTyp = 1


    End Sub
End Class
