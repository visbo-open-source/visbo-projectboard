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
    Public Property showTimeSpanInPT As Boolean

    Public Property AmpelGruen As Long
    Public Property AmpelGelb As Long
    Public Property AmpelRot As Long
    Public Property AmpelNichtBewertet As Long

    Public Property glowColor As Long

    ' Settings für die Projekteingabe
    Public Property lastProjektTyp As String
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
    Public Property zeilenhoehe1 As Double
    Public Property zeilenhoehe2 As Double
    Public Property spaltenbreite As Double
    Public Property offsetEinheit As String
    Public Property drawphases As Boolean
    Public Property applyFilter As Boolean
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

    ' Settings für die letzte User Selektion in der Tafel 
    Public Property selectedColumn As Integer
    Public Property selectedRow As Integer

    ' Settings für Import / Export
    Public Property EinzelRessExport As Integer

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
        _milestoneFreeFloat = True
        _autoCorrectBedarfe = True
        _propAnpassRess = False
        _phasesProzentual = False
        _drawphases = False
        _showValuesOfSelected = False
        _applyFilter = False
        _createIfNotThere = True

        ' Settings für Import / Export 
        _EinzelRessExport = 0



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
        _importTyp = 1


    End Sub
End Class
