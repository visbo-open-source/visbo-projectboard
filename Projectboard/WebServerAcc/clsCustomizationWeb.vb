Imports ProjectBoardDefinitions
Public Class clsCustomizationWeb


    Public businessUnitDefinitions As List(Of clsBusinessUnit)
    Public phaseDefinitions As List(Of clsPhasenDefinition)
    Public milestoneDefinitions As List(Of clsMeilensteinDefinition)

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

    ' tk 21.6.21 One Person has one skill 
    Private _onePersonOneRole As Boolean

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

    Public Property onePersonOneRole As Boolean
        Get
            onePersonOneRole = _onePersonOneRole
        End Get
        Set(value As Boolean)
            If Not IsNothing(value) Then
                _onePersonOneRole = value
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

    Public Sub copyTo(ByRef customDef As clsCustomization)

        With customDef

            Dim index As Integer
            For index = 0 To Me.businessUnitDefinitions.Count - 1
                customDef.businessUnitDefinitions.Add(index, Me.businessUnitDefinitions.ElementAt(index))
            Next

            For Each phasedef As clsPhasenDefinition In Me.phaseDefinitions
                customDef.phaseDefinitions.Add(phasedef)
            Next
            For Each msdef As clsMeilensteinDefinition In Me.milestoneDefinitions
                customDef.milestoneDefinitions.Add(msdef)
            Next

            .showtimezone_color = Me.showtimezone_color
            .noshowtimezone_color = Me.noshowtimezone_color
            .calendarFontColor = Me.calendarFontColor
            .nrOfDaysMonth = Me.nrOfDaysMonth
            .farbeInternOP = Me.farbeInternOP
            .farbeExterne = Me.farbeExterne
            .iProjektFarbe = Me.iProjektFarbe
            .iWertFarbe = Me.iWertFarbe
            .vergleichsfarbe0 = Me.vergleichsfarbe0
            .vergleichsfarbe1 = Me.vergleichsfarbe1
            'customizations.vergleichsfarbe2 = vergleichsfarbe2

            .SollIstFarbeB = Me.SollIstFarbeB
            .SollIstFarbeL = Me.SollIstFarbeL
            .SollIstFarbeC = Me.SollIstFarbeC
            .AmpelGruen = Me.AmpelGruen
            'tmpcolor = CType(.Range("AmpelGruen").Interior.Color, Microsoft.Office.Interop.Excel.ColorFormat)
            .AmpelGelb = Me.AmpelGelb
            .AmpelRot = Me.AmpelRot
            .AmpelNichtBewertet = Me.AmpelNichtBewertet
            .glowColor = Me.glowColor

            .timeSpanColor = Me.timeSpanColor
            .showTimeSpanInPT = Me.showTimeSpanInPT

            .gridLineColor = Me.gridLineColor

            .missingDefinitionColor = Me.missingDefinitionColor

            .onePersonOneRole = Me.onePersonOneRole
            .allianzIstDatenReferate = Me.allianzIstDatenReferate

            .autoSetActualDataDate = Me.autoSetActualDataDate

            .actualDataMonth = Me.actualDataMonth
            .ergebnisfarbe1 = Me.ergebnisfarbe1
            .ergebnisfarbe2 = Me.ergebnisfarbe2
            .weightStrategicFit = Me.weightStrategicFit
            .kalenderStart = Me.kalenderStart.ToLocalTime
            .zeitEinheit = Me.zeitEinheit
            .kapaEinheit = Me.kapaEinheit
            .offsetEinheit = Me.offsetEinheit
            .EinzelRessExport = Me.EinzelRessExport
            .zeilenhoehe1 = Me.zeilenhoehe1
            .zeilenhoehe2 = Me.zeilenhoehe2
            .spaltenbreite = Me.spaltenbreite
            .autoCorrectBedarfe = Me.autoCorrectBedarfe
            .propAnpassRess = Me.propAnpassRess
            .showValuesOfSelected = Me.showValuesOfSelected

            .mppProjectsWithNoMPmayPass = Me.mppProjectsWithNoMPmayPass
            .fullProtocol = Me.fullProtocol
            .addMissingPhaseMilestoneDef = Me.addMissingPhaseMilestoneDef
            .alwaysAcceptTemplateNames = Me.alwaysAcceptTemplateNames
            .eliminateDuplicates = Me.eliminateDuplicates
            .importUnknownNames = Me.importUnknownNames
            .createUniqueSiblingNames = Me.createUniqueSiblingNames

            .readWriteMissingDefinitions = Me.readWriteMissingDefinitions
            .meExtendedColumnsView = Me.meExtendedColumnsView
            .meDontAskWhenAutoReduce = Me.meDontAskWhenAutoReduce
            .readCostRolesFromDB = Me.readCostRolesFromDB

            .importTyp = Me.importTyp

            .meAuslastungIsInclExt = Me.meAuslastungIsInclExt

            .englishLanguage = Me.englishLanguage

            .showPlaceholderAndAssigned = Me.showPlaceholderAndAssigned
            .considerRiskFee = Me.considerRiskFee


        End With
    End Sub

    Public Sub copyFrom(ByVal customDef As clsCustomization)

        With customDef

            For Each kvp As KeyValuePair(Of Integer, clsBusinessUnit) In customDef.businessUnitDefinitions
                Me.businessUnitDefinitions.Add(kvp.Value)
            Next
            For Each kvp As KeyValuePair(Of String, clsPhasenDefinition) In customDef.phaseDefinitions.liste
                Me.phaseDefinitions.Add(kvp.Value)
            Next
            For Each kvp As KeyValuePair(Of String, clsMeilensteinDefinition) In customDef.milestoneDefinitions.liste
                Me.milestoneDefinitions.Add(kvp.Value)
            Next
            Me.showtimezone_color = .showtimezone_color
            Me.noshowtimezone_color = .noshowtimezone_color
            Me.calendarFontColor = .calendarFontColor
            Me.nrOfDaysMonth = .nrOfDaysMonth
            Me.farbeInternOP = .farbeInternOP
            Me.farbeExterne = .farbeExterne
            Me.iProjektFarbe = .iProjektFarbe
            Me.iWertFarbe = .iWertFarbe
            Me.vergleichsfarbe0 = .vergleichsfarbe0
            Me.vergleichsfarbe1 = .vergleichsfarbe1
            'customizations.vergleichsfarbe2 = vergleichsfarbe2

            Me.SollIstFarbeB = .SollIstFarbeB
            Me.SollIstFarbeL = .SollIstFarbeL
            Me.SollIstFarbeC = .SollIstFarbeC
            Me.AmpelGruen = .AmpelGruen
            'tmpcolor = CType(.Range("AmpelGruen").Interior.Color, Microsoft.Office.Interop.Excel.ColorFormat)
            Me.AmpelGelb = .AmpelGelb
            Me.AmpelRot = .AmpelRot
            Me.AmpelNichtBewertet = .AmpelNichtBewertet
            Me.glowColor = .glowColor

            Me.timeSpanColor = .timeSpanColor
            Me.showTimeSpanInPT = .showTimeSpanInPT

            Me.gridLineColor = .gridLineColor

            Me.missingDefinitionColor = .missingDefinitionColor

            Me.onePersonOneRole = .onePersonOneRole
            Me.allianzIstDatenReferate = .allianzIstDatenReferate

            Me.autoSetActualDataDate = .autoSetActualDataDate

            Me.actualDataMonth = .actualDataMonth
            Me.ergebnisfarbe1 = .ergebnisfarbe1
            Me.ergebnisfarbe2 = .ergebnisfarbe2
            Me.weightStrategicFit = .weightStrategicFit
            Me.kalenderStart = .kalenderStart.ToUniversalTime
            Me.zeitEinheit = .zeitEinheit
            Me.kapaEinheit = .kapaEinheit
            Me.offsetEinheit = .offsetEinheit
            Me.EinzelRessExport = .EinzelRessExport
            Me.zeilenhoehe1 = .zeilenhoehe1
            Me.zeilenhoehe2 = .zeilenhoehe2
            Me.spaltenbreite = .spaltenbreite
            Me.autoCorrectBedarfe = .autoCorrectBedarfe
            Me.propAnpassRess = .propAnpassRess
            Me.showValuesOfSelected = .showValuesOfSelected

            Me.mppProjectsWithNoMPmayPass = .mppProjectsWithNoMPmayPass
            Me.fullProtocol = .fullProtocol
            Me.addMissingPhaseMilestoneDef = .addMissingPhaseMilestoneDef
            Me.alwaysAcceptTemplateNames = .alwaysAcceptTemplateNames
            Me.eliminateDuplicates = .eliminateDuplicates
            Me.importUnknownNames = .importUnknownNames
            Me.createUniqueSiblingNames = .createUniqueSiblingNames

            Me.readWriteMissingDefinitions = .readWriteMissingDefinitions
            Me.meExtendedColumnsView = .meExtendedColumnsView
            Me.meDontAskWhenAutoReduce = .meDontAskWhenAutoReduce
            Me.readCostRolesFromDB = .readCostRolesFromDB

            Me.importTyp = .importTyp

            Me.meAuslastungIsInclExt = .meAuslastungIsInclExt

            Me.englishLanguage = .englishLanguage

            Me.showPlaceholderAndAssigned = .showPlaceholderAndAssigned
            Me.considerRiskFee = .considerRiskFee

        End With
    End Sub

    Public Sub New()

        businessUnitDefinitions = New List(Of clsBusinessUnit)
        phaseDefinitions = New List(Of clsPhasenDefinition)
        milestoneDefinitions = New List(Of clsMeilensteinDefinition)

        showtimezone_color = 0
        noshowtimezone_color = 0
        calendarFontColor = 0
        nrOfDaysMonth = 20.8
        farbeInternOP = 0
        farbeExterne = 0
        iProjektFarbe = 0
        iWertFarbe = 0
        vergleichsfarbe0 = 0
        vergleichsfarbe1 = 0

        SollIstFarbeB = 0
        SollIstFarbeL = 0
        SollIstFarbeC = 0
        AmpelGruen = 0
        AmpelGelb = 0
        AmpelRot = 0
        AmpelNichtBewertet = 0
        glowColor = 0
        ' bis hier Properties definiert

        timeSpanColor = 0
        showTimeSpanInPT = 0
        gridLineColor = 0
        missingDefinitionColor = 0

        _onePersonOneRole = False

        allianzIstDatenReferate = ""
        autoSetActualDataDate = False
        actualDataMonth = Date.MinValue

        ergebnisfarbe1 = 0
        ergebnisfarbe2 = 0
        weightStrategicFit = 0.00
        ' jetzt wird KalenderStart, Zeiteinheit und Datenbank Name ausgelesen 
        kalenderStart = Date.MinValue
        zeitEinheit = ""
        kapaEinheit = ""

        offsetEinheit = ""
        'ur: 6.08.2015: umgestellt auf Settings in app.config ''awinSettings.databaseName = CStr(.Range("Datenbank").Value)
        EinzelRessExport = 0
        zeilenhoehe1 = 50.0
        zeilenhoehe2 = 20.0
        spaltenbreite = 4.5
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
