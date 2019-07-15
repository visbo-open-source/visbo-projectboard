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

    Public Property timeSpanColor As Long
        Get
            timeSpanColor = _timeSpanColor
        End Get
        Set(value As Long)
            If Not IsNothing(value) Then
                _timeSpanColor = value
            End If
        End Set
    End Property

    Public Property showTimeSpanInPT As Long
        Get
            showTimeSpanInPT = _showTimeSpanInPT
        End Get
        Set(value As Long)
            If Not IsNothing(value) Then
                _showTimeSpanInPT = value
            End If
        End Set
    End Property

    Public Property gridLineColor As Long
        Get
            gridLineColor = _gridLineColor
        End Get
        Set(value As Long)
            If Not IsNothing(value) Then
                _gridLineColor = value
            End If
        End Set
    End Property

    Public Property missingDefinitionColor As Long
        Get
            missingDefinitionColor = _missingDefinitionColor
        End Get
        Set(value As Long)
            If Not IsNothing(value) Then
                _missingDefinitionColor = value
            End If
        End Set
    End Property



    Public Property allianzIstDatenReferate As String
        Get
            allianzIstDatenReferate = _allianzIstDatenReferate
        End Get
        Set(value As String)
            If Not IsNothing(value) Then
                _allianzIstDatenReferate = value
            End If
        End Set
    End Property

    Public Property autoSetActualDataDate As Boolean
        Get
            autoSetActualDataDate = _autoSetActualDataDate
        End Get
        Set(value As Boolean)
            If Not IsNothing(value) Then
                _autoSetActualDataDate = value
            End If
        End Set
    End Property

    Public Property actualDataMonth As Date
        Get
            actualDataMonth = _actualDataMonth
        End Get
        Set(value As Date)
            If Not IsNothing(value) Then
                _actualDataMonth = value
            End If
        End Set
    End Property

    Public Property ergebnisfarbe1 As Long
        Get
            ergebnisfarbe1 = _ergebnisfarbe1
        End Get
        Set(value As Long)
            If Not IsNothing(value) Then
                _ergebnisfarbe1 = value
            End If
        End Set
    End Property

    Public Property ergebnisfarbe2 As Long
        Get
            ergebnisfarbe2 = _ergebnisfarbe2
        End Get
        Set(value As Long)
            If Not IsNothing(value) Then
                _ergebnisfarbe2 = value
            End If
        End Set
    End Property

    Public Property weightStrategicFit As Double
        Get
            weightStrategicFit = _weightStrategicFit
        End Get
        Set(value As Double)
            If Not IsNothing(value) Then
                _weightStrategicFit = value
            End If
        End Set
    End Property

    Public Property kalenderStart As Date
        Get
            kalenderStart = _kalenderStart
        End Get
        Set(value As Date)
            If Not IsNothing(value) Then
                _kalenderStart = value
            End If
        End Set
    End Property

    Public Property zeitEinheit As String
        Get
            zeitEinheit = _zeitEinheit
        End Get
        Set(value As String)
            If Not IsNothing(value) Then
                _zeitEinheit = value
            End If
        End Set
    End Property

    Public Property kapaEinheit As String
        Get
            kapaEinheit = _kapaEinheit
        End Get
        Set(value As String)
            If Not IsNothing(value) Then
                _kapaEinheit = value
            End If
        End Set
    End Property

    Public Property offsetEinheit As String
        Get
            offsetEinheit = _offsetEinheit
        End Get
        Set(value As String)
            If Not IsNothing(value) Then
                _offsetEinheit = value
            End If
        End Set
    End Property

    Public Property EinzelRessExport As Integer
        Get
            EinzelRessExport = _EinzelRessExport
        End Get
        Set(value As Integer)
            If Not IsNothing(value) Then
                _EinzelRessExport = value
            End If
        End Set
    End Property

    Public Property zeilenhoehe1 As Double
        Get
            zeilenhoehe1 = _zeilenhoehe1
        End Get
        Set(value As Double)
            If Not IsNothing(value) Then
                _zeilenhoehe1 = value
            End If
        End Set
    End Property

    Public Property zeilenhoehe2 As Double
        Get
            zeilenhoehe2 = _zeilenhoehe2
        End Get
        Set(value As Double)
            If Not IsNothing(value) Then
                _zeilenhoehe2 = value
            End If
        End Set
    End Property

    Public Property spaltenbreite As Double
        Get
            spaltenbreite = _spaltenbreite
        End Get
        Set(value As Double)
            If Not IsNothing(value) Then
                _spaltenbreite = value
            End If
        End Set
    End Property
    ' Vorbesetzung hier = true
    Public Property autoCorrectBedarfe As Boolean
        Get
            autoCorrectBedarfe = _autoCorrectBedarfe
        End Get
        Set(value As Boolean)
            If Not IsNothing(value) Then
                _autoCorrectBedarfe = value
            End If
        End Set
    End Property
    ' Vorbesetzung hier = false
    Public Property propAnpassRess As Boolean
        Get
            propAnpassRess = _propAnpassRess
        End Get
        Set(value As Boolean)
            If Not IsNothing(value) Then
                _propAnpassRess = value
            End If
        End Set
    End Property
    ' Vorbesetzung hier = false
    Public Property showValuesOfSelected As Boolean
        Get
            showValuesOfSelected = _showValuesOfSelected
        End Get
        Set(value As Boolean)
            If Not IsNothing(value) Then
                _showValuesOfSelected = value
            End If
        End Set
    End Property



    ' gibt es die Einstellung für ProjectWithNoMPmayPass
    Public Property mppProjectsWithNoMPmayPass As Boolean
        Get
            mppProjectsWithNoMPmayPass = _mppProjectsWithNoMPmayPass
        End Get
        Set(value As Boolean)
            If Not IsNothing(value) Then
                _mppProjectsWithNoMPmayPass = value
            End If
        End Set
    End Property

    ' ist Einstellung für volles Protokoll vorhanden ? 
    Public Property fullProtocol As Boolean
        Get
            fullProtocol = _fullProtocol
        End Get
        Set(value As Boolean)
            If Not IsNothing(value) Then
                _fullProtocol = value
            End If
        End Set
    End Property

    ' Einstellung für addMissingDefinitions
    Public Property addMissingPhaseMilestoneDef As Boolean
        Get
            addMissingPhaseMilestoneDef = _addMissingPhaseMilestoneDef
        End Get
        Set(value As Boolean)
            If Not IsNothing(value) Then
                _addMissingPhaseMilestoneDef = value
            End If
        End Set
    End Property

    ' Einstellung für alwaysAcceptTemplate Names
    Public Property alwaysAcceptTemplateNames As Boolean
        Get
            alwaysAcceptTemplateNames = _alwaysAcceptTemplateNames
        End Get
        Set(value As Boolean)
            If Not IsNothing(value) Then
                _alwaysAcceptTemplateNames = value
            End If
        End Set
    End Property
    ' Einstellungen, um Duplikate zu eliminieren ; 
    Public Property eliminateDuplicates As Boolean
        Get
            eliminateDuplicates = _eliminateDuplicates
        End Get
        Set(value As Boolean)
            If Not IsNothing(value) Then
                _eliminateDuplicates = value
            End If
        End Set
    End Property

    ' Einstellungen, um unbekannte Namen zu importieren 
    Public Property importUnknownNames As Boolean
        Get

            importUnknownNames = _importUnknownNames
        End Get
        Set(value As Boolean)
            If Not IsNothing(value) Then
                _importUnknownNames = value
            End If
        End Set
    End Property

    ' Einstellung, um Geschwister-Namen immer eindeutig zu machen
    Public Property createUniqueSiblingNames As Boolean
        Get

            createUniqueSiblingNames = _createUniqueSiblingNames
        End Get
        Set(value As Boolean)
            If Not IsNothing(value) Then
                _createUniqueSiblingNames = value
            End If
        End Set
    End Property

    ' Einstellung, um das Lesen / Schreiben von MissingDefinitions zu steuern
    Public Property readWriteMissingDefinitions As Boolean
        Get
            readWriteMissingDefinitions = _readWriteMissingDefinitions
        End Get
        Set(value As Boolean)
            If Not IsNothing(value) Then
                _readWriteMissingDefinitions = value
            End If
        End Set
    End Property

    ' Einstellung, um für MassEdit zu steuern, ob %-tuale Spalte auch angezeigt werden soll
    Public Property meExtendedColumnsView As Boolean
        Get
            meExtendedColumnsView = _meExtendedColumnsView
        End Get
        Set(value As Boolean)
            If Not IsNothing(value) Then
                _meExtendedColumnsView = value
            End If
        End Set
    End Property
    ' Einstellung, um im MassEdit für AutoReduce zu steuern, ob nachgefragt wird, bevor von Folge- oder VorgängerMonaten die Ressource zu holen
    Public Property meDontAskWhenAutoReduce As Boolean
        Get
            meDontAskWhenAutoReduce = _meDontAskWhenAutoReduce
        End Get
        Set(value As Boolean)
            If Not IsNothing(value) Then
                _meDontAskWhenAutoReduce = value
            End If
        End Set
    End Property

    ' Einstellung, um zu signalisieren, dass Rollen und Kosten ausschließlich von der DB gelesen werden sollen ; 
    Public Property readCostRolesFromDB As Boolean
        Get
            readCostRolesFromDB = _readCostRolesFromDB
        End Get
        Set(value As Boolean)
            If Not IsNothing(value) Then
                _readCostRolesFromDB = value
            End If
        End Set
    End Property

    ' Import Typ regelt, um welche DateiFormate es sich bei dem Import handelt
    ' 1: Standard
    ' 2: BMW Rplan Export in Excel 
    Public Property importTyp As Integer
        Get
            importTyp = _importTyp
        End Get
        Set(value As Integer)
            If Not IsNothing(value) Then
                _importTyp = value
            End If
        End Set
    End Property
    ' sollen im Massen-Edit bei der Berechnung der auslastungsWerte die externen aus der Kapa-Datei mitberücksichtigt werden ?    
    Public Property meAuslastungIsInclExt As Boolean
        Get
            meAuslastungIsInclExt = _meAuslastungIsInclExt
        End Get
        Set(value As Boolean)
            If Not IsNothing(value) Then
                _meAuslastungIsInclExt = value
            End If
        End Set
    End Property

    ' welche Sprache soll verwendet werden: wenn english, alles andere ist deutsch

    Public Property englishLanguage As Boolean
        Get
            englishLanguage = _englishLanguage
        End Get
        Set(value As Boolean)
            If Not IsNothing(value) Then
                _englishLanguage = value
            End If
        End Set
    End Property

    'Private _menuCult As Globalization.CultureInfo
    Public Property menuCult As Globalization.CultureInfo
        Get
            menuCult = _menuCult
        End Get
        Set(value As Globalization.CultureInfo)
            If Not IsNothing(value) Then
                _menuCult = value
            End If
        End Set
    End Property
    ' sollen Sam Try
    Public Property showPlaceholderAndAssigned As Boolean
        Get
            showPlaceholderAndAssigned = _showPlaceholderAndAssigned
        End Get
        Set(value As Boolean)
            If Not IsNothing(value) Then
                _showPlaceholderAndAssigned = value
            End If
        End Set
    End Property

    ' sollen die Risiko Kennzahlen bei der Berechnung der Portfolio / Projekt-Ergebnisse mitgerechnet werden ?  
    Public Property considerRiskFee As Boolean
        Get
            considerRiskFee = _considerRiskFee
        End Get
        Set(value As Boolean)
            If Not IsNothing(value) Then
                _considerRiskFee = value
            End If
        End Set
    End Property
    Public Sub New()
        businessUnitDefinitions = New List(Of clsBusinessUnit)
        phaseDefinitions = New List(Of clsPhasenDefinition)
        milestoneDefinitions = New List(Of clsMeilensteinDefinition)

        showtimezone_color = Nothing
        noshowtimezone_color
        calendarFontColor
        nrOfDaysMonth
        farbeInternOP
        farbeExterne
        iProjektFarbe As Long
      iWertFarbe As Long
        vergleichsfarbe0 As Long
        vergleichsfarbe1 As Long

        SollIstFarbeB As Long
        SollIstFarbeL As Long
        SollIstFarbeC As Long
        AmpelGruen As Long
        AmpelGelb As Long
        AmpelRot As Long
        AmpelNichtBewertet As Long
        glowColor As Long
        ' bis hier Properties definiert

        timeSpanColor As Long
        showTimeSpanInPT As Long
        gridLineColor As Long
        missingDefinitionColor As Long
            
        allianzIstDatenReferate As String
        autoSetActualDataDate As Boolean
        actualDataMonth As Date

        ergebnisfarbe1 As Long
        ergebnisfarbe2 As Long
        weightStrategicFit As Double
        ' jetzt wird KalenderStart, Zeiteinheit und Datenbank Name ausgelesen 
        kalenderStart As Date
        zeitEinheit As String
        kapaEinheit As String

        offsetEinheit As String
        'ur: 6.08.2015: umgestellt auf Settings in app.config ''awinSettings.databaseName = CStr(.Range("Datenbank").Value)
        EinzelRessExport As Integer
        zeilenhoehe1 As Double
        zeilenhoehe2 As Double
        spaltenbreite As Double
        autoCorrectBedarfe = True
        propAnpassRess = False
        showValuesOfSelected = False


        ' gibt es die Einstellung für ProjectWithNoMPmayPass
        mppProjectsWithNoMPmayPass = False

        ' ist Einstellung für volles Protokoll vorhanden ? 
        fullProtocol = False

        ' Einstellung für addMissingDefinitions
        addMissingPhaseMilestoneDef = False

        ' Einstellung für alwaysAcceptTemplate Names
        alwaysAcceptTemplateNames = False

        ' Einstellungen, um Duplikate zu eliminieren ; 
        eliminateDuplicates = True

        ' Einstellungen, um unbekannte Namen zu importieren 
        importUnknownNames = True

        ' Einstellung, um Geschwister-Namen immer eindeutig zu machen
        createUniqueSiblingNames = True

        ' Einstellung, um das Lesen / Schreiben von MissingDefinitions zu steuern
        readWriteMissingDefinitions = False

        ' Einstellung, um für MassEdit zu steuern, ob %-tuale Spalte auch angezeigt werden soll
        meExtendedColumnsView = False

        ' Einstellung, um im MassEdit für AutoReduce zu steuern, ob nachgefragt wird, bevor von Folge- oder VorgängerMonaten die Ressource zu holen
        meDontAskWhenAutoReduce = True

        ' Einstellung, um zu signalisieren, dass Rollen und Kosten ausschließlich von der DB gelesen werden sollen ; 
        readCostRolesFromDB = True

        ' Import Typ regelt, um welche DateiFormate es sich bei dem Import handelt
        ' 1: Standard
        ' 2: BMW Rplan Export in Excel 
        importTyp = 1

        ' sollen im Massen-Edit bei der Berechnung der auslastungsWerte die externen aus der Kapa-Datei mitberücksichtigt werden ?    
        meAuslastungIsInclExt = True

        ' welche Sprache soll verwendet werden: wenn english, alles andere ist deutsch
        englishLanguage = True
        menuCult = ReportLang(PTSprache.englisch)
        repCult = menuCult

        ' sollen Sam Try
        showPlaceholderAndAssigned = False

        ' sollen die Risiko Kennzahlen bei der Berechnung der Portfolio / Projekt-Ergebnisse mitgerechnet werden ?  
        considerRiskFee = False
    End Sub
End Class
