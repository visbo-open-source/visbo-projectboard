Imports ProjectBoardDefinitions
Public Class clsCustomizationWeb
    Private _businessUnitDefinitions As List(Of clsBusinessUnit)
    Private _phaseDefinitions As List(Of clsPhasenDefinition)
    Private _milestoneDefinitions As List(Of clsMeilensteinDefinition)

    Private _showtimezone_color As Long
    Private _noshowtimezone_color As Long
    Private _calendarFontColor As Long
    Private _nrOfDaysMonth As Double
    Private _farbeInternOP As Long
    Private _farbeExterne As Long
    Private _iProjektFarbe As Long
    Private _iWertFarbe As Long
    Private _vergleichsfarbe0 As Long
    Private _vergleichsfarbe1 As Long

    Private _SollIstFarbeB As Long
    Private _SollIstFarbeL As Long
    Private _SollIstFarbeC As Long
    Private _AmpelGruen As Long
    Private _AmpelGelb As Long
    Private _AmpelRot As Long
    Private _AmpelNichtBewertet As Long
    Private _glowColor As Long
    ' bis hier Properties definiert

    Private _timeSpanColor As Long
    Private _showTimeSpanInPT As Long
    Private _gridLineColor As Long
    Private _missingDefinitionColor As Long

    Private _allianzIstDatenReferate As String
    Private _autoSetActualDataDate As Boolean
    Private _actualDataMonth As Date

    Private _ergebnisfarbe1 As Long
    Private _ergebnisfarbe2 As Long
    Private _weightStrategicFit As Double
    ' jetzt wird KalenderStart, Zeiteinheit und Datenbank Name ausgelesen 
    Private _kalenderStart As Date
    Private _zeitEinheit As String
    Private _kapaEinheit As String

    Private _offsetEinheit As String
    'ur: 6.08.2015: umgestellt auf Settings in app.config ''awinSettings.databaseName = CStr(.Range("Datenbank").Value)
    Private _EinzelRessExport As Integer
    Private _zeilenhoehe1 As Double
    Private _zeilenhoehe2 As Double
    Private _spaltenbreite As Double
    Private _autoCorrectBedarfe As Boolean = True
    Private _propAnpassRess As Boolean = False
    Private _showValuesOfSelected As Boolean = False


    ' gibt es die Einstellung für ProjectWithNoMPmayPass
    Private _mppProjectsWithNoMPmayPass As Boolean

    ' ist Einstellung für volles Protokoll vorhanden ? 
    Private _fullProtocol As Boolean

    ' Einstellung für addMissingDefinitions
    Private _addMissingPhaseMilestoneDef As Boolean

    ' Einstellung für alwaysAcceptTemplate Names
    Private _alwaysAcceptTemplateNames As Boolean

    ' Einstellungen, um Duplikate zu eliminieren ; 
    Private _eliminateDuplicates As Boolean

    ' Einstellungen, um unbekannte Namen zu importieren 
    Private _importUnknownNames As Boolean

    ' Einstellung, um Geschwister-Namen immer eindeutig zu machen
    Private _createUniqueSiblingNames As Boolean

    ' Einstellung, um das Lesen / Schreiben von MissingDefinitions zu steuern
    Private _readWriteMissingDefinitions As Boolean

    ' Einstellung, um für MassEdit zu steuern, ob %-tuale Spalte auch angezeigt werden soll
    Private _meExtendedColumnsView As Boolean

    ' Einstellung, um im MassEdit für AutoReduce zu steuern, ob nachgefragt wird, bevor von Folge- oder VorgängerMonaten die Ressource zu holen
    Private _meDontAskWhenAutoReduce As Boolean

    ' Einstellung, um zu signalisieren, dass Rollen und Kosten ausschließlich von der DB gelesen werden sollen ; 
    Private _readCostRolesFromDB As Boolean

    ' Import Typ regelt, um welche DateiFormate es sich bei dem Import handelt
    ' 1: Standard
    ' 2: BMW Rplan Export in Excel 
    Private _importTyp As Integer

    ' sollen im Massen-Edit bei der Berechnung der auslastungsWerte die externen aus der Kapa-Datei mitberücksichtigt werden ?    
    Private _meAuslastungIsInclExt As Boolean

    ' welche Sprache soll verwendet werden: wenn english, alles andere ist deutsch

    Private _englishLanguage As Boolean
    Private _menuCult As Globalization.CultureInfo

    ' sollen Sam Try
    Private _showPlaceholderAndAssigned As Boolean

    ' sollen die Risiko Kennzahlen bei der Berechnung der Portfolio / Projekt-Ergebnisse mitgerechnet werden ?  
    Private _considerRiskFee As Boolean

    Public Property businessUnitDefinitions As List(Of clsBusinessUnit)
        Get
            businessUnitDefinitions = _businessUnitDefinitions
        End Get
        Set(value As List(Of clsBusinessUnit))
            If Not IsNothing(value) Then
                _businessUnitDefinitions = value
            End If
        End Set
    End Property

    Public Property phaseDefinitions As List(Of clsPhasenDefinition)
        Get
            phaseDefinitions = _phaseDefinitions
        End Get
        Set(value As List(Of clsPhasenDefinition))
            If Not IsNothing(value) Then
                _phaseDefinitions = value
            End If
        End Set
    End Property

    Public Property milestoneDefinitions As List(Of clsMeilensteinDefinition)
        Get
            milestoneDefinitions = _milestoneDefinitions
        End Get
        Set(value As List(Of clsMeilensteinDefinition))
            If Not IsNothing(value) Then
                _milestoneDefinitions = value
            End If
        End Set
    End Property

    Public Property showtimezone_color As Long
        Get
            showtimezone_color = _showtimezone_color
        End Get
        Set(value As Long)
            If Not IsNothing(value) Then
                _showtimezone_color = value
            End If
        End Set
    End Property

    Public Property noshowtimezone_color As Long
        Get
            noshowtimezone_color = _noshowtimezone_color
        End Get
        Set(value As Long)
            If Not IsNothing(value) Then
                _noshowtimezone_color = value
            End If
        End Set
    End Property

    Public Property calendarFontColor As Long
        Get
            calendarFontColor = _calendarFontColor
        End Get
        Set(value As Long)
            If Not IsNothing(value) Then
                _calendarFontColor = value
            End If
        End Set
    End Property


    Public Property nrOfDaysMonth As Double
        Get
            nrOfDaysMonth = _nrOfDaysMonth
        End Get
        Set(value As Double)
            If Not IsNothing(value) Then
                _nrOfDaysMonth = value
            End If
        End Set
    End Property


    Public Property farbeInternOP As Long
        Get
            farbeInternOP = _farbeInternOP
        End Get
        Set(value As Long)
            If Not IsNothing(value) Then
                _farbeInternOP = value
            End If
        End Set
    End Property

    Public Property farbeExterne As Long
        Get
            farbeExterne = _farbeExterne
        End Get
        Set(value As Long)
            If Not IsNothing(value) Then
                _farbeExterne = value
            End If
        End Set
    End Property

    Public Property iProjektFarbe As Long
        Get
            iProjektFarbe = _iProjektFarbe
        End Get
        Set(value As Long)
            If Not IsNothing(value) Then
                _iProjektFarbe = value
            End If
        End Set
    End Property

    Public Property iWertFarbe As Long
        Get
            iWertFarbe = _iWertFarbe
        End Get
        Set(value As Long)
            If Not IsNothing(value) Then
                _iWertFarbe = value
            End If
        End Set
    End Property

    Public Property vergleichsfarbe0 As Long
        Get
            vergleichsfarbe0 = _vergleichsfarbe0
        End Get
        Set(value As Long)
            If Not IsNothing(value) Then
                _vergleichsfarbe0 = value
            End If
        End Set
    End Property

    Public Property vergleichsfarbe1 As Long
        Get
            vergleichsfarbe1 = _vergleichsfarbe1
        End Get
        Set(value As Long)
            If Not IsNothing(value) Then
                _vergleichsfarbe1 = value
            End If
        End Set
    End Property


    Public Property SollIstFarbeB As Long
        Get
            SollIstFarbeB = _SollIstFarbeB
        End Get
        Set(value As Long)
            If Not IsNothing(value) Then
                _SollIstFarbeB = value
            End If
        End Set
    End Property

    Public Property SollIstFarbeL As Long
        Get
            SollIstFarbeL = _SollIstFarbeL
        End Get
        Set(value As Long)
            If Not IsNothing(value) Then
                _SollIstFarbeL = value
            End If
        End Set
    End Property

    Public Property SollIstFarbeC As Long
        Get
            SollIstFarbeC = _SollIstFarbeC
        End Get
        Set(value As Long)
            If Not IsNothing(value) Then
                _SollIstFarbeC = value
            End If
        End Set
    End Property

    Public Property AmpelGruen As Long
        Get
            AmpelGruen = _AmpelGruen
        End Get
        Set(value As Long)
            If Not IsNothing(value) Then
                _AmpelGruen = value
            End If
        End Set
    End Property

    Public Property AmpelGelb As Long
        Get
            AmpelGelb = _AmpelGelb
        End Get
        Set(value As Long)
            If Not IsNothing(value) Then
                _AmpelGelb = value
            End If
        End Set
    End Property

    Public Property AmpelRot As Long
        Get
            AmpelRot = _AmpelRot
        End Get
        Set(value As Long)
            If Not IsNothing(value) Then
                _AmpelRot = value
            End If
        End Set
    End Property

    Public Property AmpelNichtBewertet As Long
        Get
            AmpelNichtBewertet = _AmpelNichtBewertet
        End Get
        Set(value As Long)
            If Not IsNothing(value) Then
                _AmpelNichtBewertet = value
            End If
        End Set
    End Property

    Public Property glowColor As Long
        Get
            glowColor = _glowColor
        End Get
        Set(value As Long)
            If Not IsNothing(value) Then
                _glowColor = value
            End If
        End Set
    End Property
End Class
