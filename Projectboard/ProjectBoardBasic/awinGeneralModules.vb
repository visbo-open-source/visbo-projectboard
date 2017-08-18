Imports ProjectBoardDefinitions
Imports MongoDbAccess
Imports ClassLibrary1
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
Imports System.Drawing
Imports System.Globalization

Imports Microsoft.VisualBasic
Imports System.Security.Principal




Public Module awinGeneralModules

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
        Kapazitaet = 10
        Businessunit = 11
        Beschreibung = 12
        KostenExtern = 13
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
    ''' erstellt die CacheProjekte Liste , vorläufig erstmal die DBCache
    ''' </summary>
    ''' <param name="todoListe"></param>
    ''' <remarks></remarks>
    Public Sub buildCacheProjekte(ByVal todoListe As Collection)
        Dim pName As String
        Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)

        For i As Integer = 1 To todoListe.Count
            pName = CStr(todoListe.Item(i))
            Dim hproj As clsProjekt = Nothing
            Try
                If ShowProjekte.contains(pName) Then
                    hproj = ShowProjekte.getProject(pName, True)

                    If Not noDB Then
                        ' wenn es in der DB existiert, dann im Cache aufbauen 
                        If Request.projectNameAlreadyExists(hproj.name, hproj.variantName, Date.Now) Then
                            ' für den Datenbank Cache aufbauen 
                            Dim dbProj As clsProjekt = Request.retrieveOneProjectfromDB(hproj.name, hproj.variantName, Date.Now)
                            dbCacheProjekte.upsert(dbProj)
                        End If
                    End If
                End If
                

            Catch ex As Exception

            End Try
        Next

    End Sub

    ' Änderung tk: ist ersetzt worden durch writePhaseMilestoneDefinitions

    ' ''' <summary>
    ' ''' schreibt evtl neu durch Inventur hinzugekommene Phasen in 
    ' ''' das Customization File 
    ' ''' </summary>
    ' ''' <remarks></remarks>
    'Public Sub awinWritePhaseDefinitions()

    '    Dim phaseDefs As Excel.Range
    '    Dim milestoneDefs As Excel.Range
    '    'Dim foundRow As Integer
    '    Dim phName As String, phColor As Long
    '    Dim lastrow As Excel.Range

    '    'appInstance.ScreenUpdating = False
    '    appInstance.EnableEvents = False



    '    ' hier muss jetzt das File Projekt Tafel Definitions.xlsx aufgemacht werden ...
    '    ' das File 
    '    Try
    '        appInstance.Workbooks.Open(awinPath & customizationFile)

    '    Catch ex As Exception
    '        Call MsgBox("Customization File nicht gefunden - Abbruch")
    '        Throw New ArgumentException("Customization File nicht gefunden - Abbruch")
    '    End Try

    '    appInstance.Workbooks(myCustomizationFile).Activate()
    '    Dim wsName4 As Excel.Worksheet = CType(appInstance.Worksheets(arrWsNames(4)), _
    '                                            Global.Microsoft.Office.Interop.Excel.Worksheet)

    '    phaseDefs = wsName4.Range("awin_Phasen_Definition")

    '    Dim anzZeilen As Integer = phaseDefs.Rows.Count
    '    lastrow = CType(phaseDefs.Rows(anzZeilen), Excel.Range)

    '    Dim vglsListe As New SortedList(Of String, String)
    '    Dim ergStr As String

    '    For Each c As Excel.Range In phaseDefs
    '        Try
    '            ergStr = CStr(c.Value).Trim

    '            If ergStr.Length > 0 And Not vglsListe.ContainsKey(ergStr) Then

    '                vglsListe.Add(ergStr, ergStr)

    '            End If
    '        Catch ex As Exception

    '        End Try

    '    Next


    '    ' jetzt muss getestet werden, ob jede Phase in PhaseDefinitions bereits in der Customization vorkommt 

    '    Dim i As Integer
    '    Dim darstellungsKlasse As String
    '    For i = 1 To PhaseDefinitions.Count

    '        With PhaseDefinitions.getPhaseDef(i)
    '            phName = .name
    '            phColor = CLng(PhaseDefinitions.getPhaseDef(i).farbe)
    '            darstellungsKlasse = .darstellungsKlasse
    '        End With


    '        If vglsListe.ContainsKey(phName) Then
    '            ' nichts zu tun 
    '        Else
    '            ' eintragen 
    '            lastrow = CType(phaseDefs.Rows(phaseDefs.Rows.Count), Excel.Range)
    '            CType(lastrow.EntireRow, Excel.Range).Insert(Excel.XlInsertShiftDirection.xlShiftDown)
    '            CType(lastrow.Cells(1, 1), Excel.Range).Offset(-1, 0).Value = phName.ToString
    '            CType(lastrow.Cells(1, 1), Excel.Range).Offset(-1, 0).Interior.Color = awinSettings.AmpelNichtBewertet
    '            CType(lastrow.Cells(1, 1), Excel.Range).Offset(-1, 6).Value = darstellungsKlasse


    '        End If



    '    Next i


    '    If awinSettings.addMissingPhaseMilestoneDef Then

    '        'jede Phase, die noch nicht in dem CustomizationFile ist, wird noch hinzugefügt 
    '        ' und in die PhaseDefinitions eingetragen

    '        For mPh As Integer = 1 To missingPhaseDefinitions.Count

    '            Dim missPhaseDef As clsPhasenDefinition = missingPhaseDefinitions.getPhaseDef(mPh)

    '            With missPhaseDef
    '                phName = .name
    '                phColor = CLng(missingPhaseDefinitions.getPhaseDef(mPh).farbe)
    '                darstellungsKlasse = .darstellungsKlasse
    '            End With


    '            If vglsListe.ContainsKey(phName) Then
    '                ' nichts zu tun 
    '            Else
    '                ' eintragen 
    '                lastrow = CType(phaseDefs.Rows(phaseDefs.Rows.Count), Excel.Range)
    '                CType(lastrow.EntireRow, Excel.Range).Insert(Excel.XlInsertShiftDirection.xlShiftDown)
    '                CType(lastrow.Cells(1, 1), Excel.Range).Offset(-1, 0).Value = phName.ToString
    '                CType(lastrow.Cells(1, 1), Excel.Range).Offset(-1, 0).Interior.Color = awinSettings.AmpelNichtBewertet
    '                CType(lastrow.Cells(1, 1), Excel.Range).Offset(-1, 6).Value = darstellungsKlasse

    '                Try
    '                    PhaseDefinitions.Add(missPhaseDef)
    '                Catch ex As Exception

    '                End Try



    '            End If


    '        Next mPh

    '        missingPhaseDefinitions.Clear()

    '    End If

    '    ' jetzt noch die Meilensteine schreiben 
    '    ' awin_Meilenstein_Definition

    '    milestoneDefs = wsName4.Range("awin_Meilenstein_Definition")
    '    anzZeilen = milestoneDefs.Rows.Count
    '    lastrow = CType(milestoneDefs.Rows(anzZeilen), Excel.Range)

    '    ' jetzt muss getestet werden, ob jede Meilenstein  in MilestoneDefinitions bereits in der Customization vorkommt 

    '    vglsListe.Clear()

    '    For Each c As Excel.Range In milestoneDefs
    '        Try
    '            ergStr = CStr(c.Value).Trim

    '            If ergStr.Length > 0 And Not vglsListe.ContainsKey(ergStr) Then

    '                vglsListe.Add(ergStr, ergStr)

    '            End If
    '        Catch ex As Exception

    '        End Try

    '    Next


    '    Dim msName As String
    '    Dim shortName As String
    '    Dim belongsTo As String


    '    For i = 1 To MilestoneDefinitions.Count

    '        With MilestoneDefinitions.elementAt(i - 1)
    '            msName = .name
    '            shortName = .shortName
    '            belongsTo = .belongsTo
    '            darstellungsKlasse = .darstellungsKlasse
    '        End With

    '        If vglsListe.ContainsKey(msName) Then
    '            ' nichts zu tun 
    '        Else
    '            ' eintragen 
    '            lastrow = CType(milestoneDefs.Rows(milestoneDefs.Rows.Count), Excel.Range)
    '            CType(lastrow.EntireRow, Excel.Range).Insert(XlInsertShiftDirection.xlShiftDown)
    '            CType(lastrow.Cells(1, 1), Excel.Range).Offset(-1, 0).Value = msName
    '            CType(lastrow.Cells(1, 1), Excel.Range).Offset(-1, 4).Value = belongsTo
    '            CType(lastrow.Cells(1, 1), Excel.Range).Offset(-1, 5).Value = shortName
    '            CType(lastrow.Cells(1, 1), Excel.Range).Offset(-1, 6).Value = darstellungsKlasse
    '            CType(lastrow.Cells(1, 1), Excel.Range).Offset(-1, 0).Interior.Color = awinSettings.AmpelNichtBewertet

    '        End If



    '    Next i


    '    If awinSettings.addMissingPhaseMilestoneDef Then

    '        ' die Meilensteine, die noch nicht in MilestoneDefinitions enthalten sind, werden nun in CustomizationFile eingetragen 
    '        ' und in die MilestoneDefinitions

    '        For mMs As Integer = 1 To missingMilestoneDefinitions.Count

    '            Dim msDef As clsMeilensteinDefinition = missingMilestoneDefinitions.elementAt(mMs - 1)
    '            With msDef
    '                msName = .name
    '                shortName = .shortName
    '                belongsTo = .belongsTo
    '                darstellungsKlasse = .darstellungsKlasse
    '            End With

    '            If vglsListe.ContainsKey(msName) Then
    '                ' nichts zu tun 
    '            Else
    '                ' eintragen 
    '                lastrow = CType(milestoneDefs.Rows(milestoneDefs.Rows.Count), Excel.Range)
    '                CType(lastrow.EntireRow, Excel.Range).Insert(XlInsertShiftDirection.xlShiftDown)
    '                CType(lastrow.Cells(1, 1), Excel.Range).Offset(-1, 0).Value = msName
    '                CType(lastrow.Cells(1, 1), Excel.Range).Offset(-1, 4).Value = belongsTo
    '                CType(lastrow.Cells(1, 1), Excel.Range).Offset(-1, 5).Value = shortName
    '                CType(lastrow.Cells(1, 1), Excel.Range).Offset(-1, 6).Value = darstellungsKlasse
    '                CType(lastrow.Cells(1, 1), Excel.Range).Offset(-1, 0).Interior.Color = awinSettings.AmpelNichtBewertet
    '                If Not MilestoneDefinitions.Contains(msDef.name) Then
    '                    MilestoneDefinitions.Add(msDef)
    '                End If


    '            End If

    '        Next mMs
    '        missingMilestoneDefinitions.Clear()

    '    End If


    '    appInstance.ActiveWorkbook.Close(SaveChanges:=True)
    '    'appInstance.ScreenUpdating = True
    '    appInstance.EnableEvents = True

    'End Sub

    ''' <summary>
    ''' schreibt evtl neu hinzugekommene Phasen und Meilensteine in 
    ''' das Customization File 
    ''' ausserdem werden Auswahl Validation Dropboxes gesetzt 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub awinWritePhaseMilestoneDefinitions(Optional ByVal writeMappings As Boolean = False)

        Dim formerSU As Boolean = appInstance.ScreenUpdating
        appInstance.ScreenUpdating = False
        appInstance.EnableEvents = False



        ' hier muss jetzt das File Projekt Tafel Definitions.xlsx aufgemacht werden ...
        ' das File 
        Try
            appInstance.Workbooks.Open(awinPath & customizationFile)

        Catch ex As Exception
            Call MsgBox("Customization File nicht gefunden - Abbruch")
            appInstance.EnableEvents = True
            appInstance.ScreenUpdating = formerSU
            Throw New ArgumentException("Customization File nicht gefunden - Abbruch")
        End Try

        appInstance.Workbooks(myCustomizationFile).Activate()

        ' schreibe die Phase- und MilestoneDefinitions
        Call WriteDefinitions(False)
        ' schreibe - in Abhängigkeit von dem Parameter . die MissingPhase- und MissingMilestone-Definitions
        If awinSettings.readWriteMissingDefinitions Then
            Call WriteDefinitions(True)
        End If


        ' prüfen , ob die Mappings-Behandlung auch gemacht werden soll ...
        If writeMappings Then

            '
            ' jetzt werden erstmal die Phase Mappings geschrieben  
            '
            Dim wsName8 As Excel.Worksheet = CType(appInstance.Worksheets(arrWsNames(8)), _
                                                Global.Microsoft.Office.Interop.Excel.Worksheet)
            Dim area As Excel.Range
            Dim letzteZeile As Integer
            Dim aktuelleZeile As Integer

            With wsName8

                ' Synonyme schreiben 
                letzteZeile = System.Math.Max(CInt(CType(.Cells(20000, 1), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlUp).Row), _
                                                CInt(CType(.Cells(20000, 6), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlUp).Row))

                If letzteZeile >= 3 Then
                    area = CType(.Range(.Cells(3, 1), .Cells(letzteZeile, 2)), Excel.Range)
                    ' alte Synonym / regEX Area löschen
                    area.Clear()
                End If
                

                ' neue Area definieren
                area = CType(.Range(.Cells(3, 1), .Cells(phaseMappings.countSynonyms + phaseMappings.countRegEx + 4, 2)), Excel.Range)

                aktuelleZeile = 1
                For ix As Integer = 1 To phaseMappings.countSynonyms
                    CType(area.Cells(aktuelleZeile, 1), Excel.Range).Value = phaseMappings.getSynonymMapping(ix - 1).Key
                    CType(area.Cells(aktuelleZeile, 2), Excel.Range).Value = phaseMappings.getSynonymMapping(ix - 1).Value
                    aktuelleZeile = aktuelleZeile + 1
                Next

                ' regular expressions schreiben 
                For ix As Integer = 1 To phaseMappings.countRegEx
                    CType(area.Cells(aktuelleZeile, 1), Excel.Range).Value = phaseMappings.getRegExMapping(ix - 1).Key
                    CType(area.Cells(aktuelleZeile, 2), Excel.Range).Value = phaseMappings.getRegExMapping(ix - 1).Value
                    aktuelleZeile = aktuelleZeile + 1
                Next


                ' ignoreNames schreiben 
                letzteZeile = CType(.Cells(20000, 6), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlUp).Row

                ' alte area löschen
                If letzteZeile >= 3 Then
                    area = CType(.Range(.Cells(3, 6), .Cells(letzteZeile, 6)), Excel.Range)
                    area.Clear()
                End If
                
                area = CType(.Range(.Cells(3, 6), .Cells(phaseMappings.countIgnore + 4, 6)), Excel.Range)
                aktuelleZeile = 1

                For ix As Integer = 1 To phaseMappings.countIgnore
                    CType(area.Cells(aktuelleZeile, 1), Excel.Range).Value = phaseMappings.getIgnoreElement(ix - 1)
                    aktuelleZeile = aktuelleZeile + 1
                Next
            End With

            '
            ' jetzt werden erstmal die Phase Mappings geschrieben  
            '
            Dim wsName10 As Excel.Worksheet = CType(appInstance.Worksheets(arrWsNames(10)), _
                                                Global.Microsoft.Office.Interop.Excel.Worksheet)

            With wsName10

                ' Synonyme schreiben 
                letzteZeile = System.Math.Max(CInt(CType(.Cells(20000, 1), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlUp).Row), _
                                                CInt(CType(.Cells(20000, 6), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlUp).Row))


                If letzteZeile >= 3 Then
                    area = CType(.Range(.Cells(3, 1), .Cells(letzteZeile, 2)), Excel.Range)
                    ' alte Area löschen
                    area.Clear()
                End If
                

                ' neue Area definieren
                area = CType(.Range(.Cells(3, 1), .Cells(milestoneMappings.countSynonyms + milestoneMappings.countRegEx + 4, 2)), Excel.Range)

                aktuelleZeile = 1
                For ix As Integer = 1 To milestoneMappings.countSynonyms
                    CType(area.Cells(aktuelleZeile, 1), Excel.Range).Value = milestoneMappings.getSynonymMapping(ix - 1).Key
                    CType(area.Cells(aktuelleZeile, 2), Excel.Range).Value = milestoneMappings.getSynonymMapping(ix - 1).Value
                    aktuelleZeile = aktuelleZeile + 1
                Next

                ' regular expressions schreiben 
                For ix As Integer = 1 To milestoneMappings.countRegEx
                    CType(area.Cells(aktuelleZeile, 1), Excel.Range).Value = milestoneMappings.getRegExMapping(ix - 1).Key
                    CType(area.Cells(aktuelleZeile, 2), Excel.Range).Value = milestoneMappings.getRegExMapping(ix - 1).Value
                    aktuelleZeile = aktuelleZeile + 1
                Next

                ' ignoreNames schreiben 
                If letzteZeile >= 3 Then
                    area = CType(.Range(.Cells(3, 6), .Cells(letzteZeile, 6)), Excel.Range)
                    ' alte Area löschen
                    area.Clear()
                End If
                
                area = CType(.Range(.Cells(3, 6), .Cells(milestoneMappings.countIgnore + 4, 6)), Excel.Range)
                aktuelleZeile = 1

                For ix As Integer = 1 To milestoneMappings.countIgnore
                    CType(area.Cells(aktuelleZeile, 1), Excel.Range).Value = milestoneMappings.getIgnoreElement(ix - 1)
                    aktuelleZeile = aktuelleZeile + 1
                Next
            End With


        End If


        appInstance.ActiveWorkbook.Close(SaveChanges:=True)
        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = formerSU

    End Sub

    ''' <summary>
    ''' schreibt die Phase-/MilestoneDefinitions bzw. die missingPhase- Milestone-Definitions
    ''' missingDefinitions werden in WorkSheet MissingDefinitions geschrieben 
    ''' </summary>
    ''' <param name="writeMissingDefinitions"></param>
    ''' <remarks></remarks>
    Private Sub WriteDefinitions(Optional ByVal writeMissingDefinitions As Boolean = False)

        Dim phaseDefs As Excel.Range
        Dim milestoneDefs As Excel.Range
        'Dim foundRow As Integer
        Dim phName As String
        Dim lastrow As Excel.Range
        Dim firstrow As Excel.Range
        Dim tmpAnzahl As Integer

        Dim msName As String
        Dim shortName As String

        Dim darstellungsKlasse As String

        Dim wsName4 As Excel.Worksheet
        

        ' beim Starten der Projekt-Tafel wird sichergestellt, dass auch das Worksheet MissingDefinitions = arrwsnames(15) existiert ...
        ' inkl der Namen der Phase- und MilestoneDefinitions
        If writeMissingDefinitions Then
            Try
                wsName4 = CType(CType(appInstance.Workbooks.Item(myCustomizationFile), Excel.Workbook).Worksheets(arrWsNames(15)), _
                                               Global.Microsoft.Office.Interop.Excel.Worksheet)
            Catch ex As Exception
                Exit Sub
            End Try
            
        Else
            wsName4 = CType(CType(appInstance.Workbooks.Item(myCustomizationFile), Excel.Workbook).Worksheets(arrWsNames(4)), _
                                                Global.Microsoft.Office.Interop.Excel.Worksheet)
        End If

        If writeMissingDefinitions Then
            Try
                phaseDefs = wsName4.Range("Missing_Phasen_Definition")
            Catch ex As Exception
                Exit Sub
            End Try

        Else
            phaseDefs = wsName4.Range("awin_Phasen_Definition")
        End If


        ' diese Range sollte auf alle Fälle mindestens eine Zeile haben 
        Dim anzZeilen As Integer = phaseDefs.Rows.Count
        lastrow = CType(phaseDefs.Rows(anzZeilen), Excel.Range)
        firstrow = CType(phaseDefs.Rows(1), Excel.Range)


        ' das folgende muss nur gemacht werden, wenn die PhaseDefinitions geschrieben werden ... 
        ' jetzt wird geprüft, ob die missingPhaseDefinitions in PhaseDefinitions übertragen werden 
        If awinSettings.addMissingPhaseMilestoneDef Then

            For ix As Integer = 1 To missingPhaseDefinitions.Count
                Try
                    PhaseDefinitions.Add(missingPhaseDefinitions.getPhaseDef(ix))
                Catch ex As Exception

                End Try

            Next

            missingPhaseDefinitions.Clear()

            ' jetzt die Meilensteine
            For ix As Integer = 1 To missingMilestoneDefinitions.Count
                Try
                    MilestoneDefinitions.Add(missingMilestoneDefinitions.getMilestoneDef(ix))
                Catch ex As Exception

                End Try

            Next

            missingMilestoneDefinitions.Clear()

        End If

        ' jetzt können erst die PhaseDefinitions, dann die MilestoneDefinitions geschrieben werden 

        ' hier werden die Validation-String aufgebaut 
        ' hier muss jetzt noch die Validierung rein .... damit der Anwender in einem nächsten Schritt sehr bequem die verschiedenen Darstellungsklassen zuweisen kann 
        Dim milestoneAppearanceClasses As String = ""
        Dim phaseAppearanceClasses As String = ""
        Dim msAppearanceRng As Excel.Range = Nothing
        Dim phAppearanceRng As Excel.Range = Nothing
        Try
            Dim wsAppearances As Excel.Worksheet = CType(CType(appInstance.Workbooks.Item(myCustomizationFile), Excel.Workbook).Worksheets(arrWsNames(7)), _
                                                Global.Microsoft.Office.Interop.Excel.Worksheet)

            For Each cl As Excel.Range In wsAppearances.Range("MeilensteinKlassen")
                If milestoneAppearanceClasses = "" Then
                    milestoneAppearanceClasses = cl.Value
                Else
                    milestoneAppearanceClasses = milestoneAppearanceClasses & ";" & cl.Value
                End If
            Next

            For Each cl As Excel.Range In wsAppearances.Range("PhasenKlassen")
                If phaseAppearanceClasses = "" Then
                    phaseAppearanceClasses = cl.Value
                Else
                    phaseAppearanceClasses = phaseAppearanceClasses & ";" & cl.Value
                End If
            Next
        Catch ex As Exception

        End Try

        ' hier muss erst mal geprüft werden, ob Zeilen eingefügt oder gelöscht werden müssen 
        ' anzZeilen muss immer um 2 größer sein als die Anzahl der Definitionen ; 
        ' die erste und letzte Zeile des Bereichs sind leer  

        Dim anzDefinitions As Integer = PhaseDefinitions.Count

        If writeMissingDefinitions Then
            anzDefinitions = missingPhaseDefinitions.Count
        Else
            anzDefinitions = PhaseDefinitions.Count
        End If


        If anzZeilen = anzDefinitions + 2 Then
        ElseIf anzZeilen < anzDefinitions + 2 Then
            ' Zeilen einfügen 

            tmpAnzahl = anzDefinitions + 2 - anzZeilen
            For ix As Integer = 1 To tmpAnzahl
                CType(lastrow.EntireRow, Excel.Range).Insert(Excel.XlInsertShiftDirection.xlShiftDown)
            Next

            ' anzZeilen und phaseDefinitions.count müssen jetzt genau gleich sein 
            anzZeilen = phaseDefs.Rows.Count

        Else
            ' Zeilen löschen
            tmpAnzahl = anzZeilen - (anzDefinitions + 2)

            For ix As Integer = 1 To tmpAnzahl
                CType(phaseDefs.Rows(2).EntireRow, Excel.Range).Delete(Excel.XlDeleteShiftDirection.xlShiftUp)
            Next

            ' jetzt sind mindestens zwei Zeilen übrig , und zwar genau dann wenn phaseDefinitions.count = 0 
            anzZeilen = phaseDefs.Rows.Count

        End If

        ' jetzt können die Phase-Definitions in den Range geschrieben werden 
        ' und zwar so, dass sie mit der 2. Zeile beginnen 


        For ix As Integer = 1 To anzDefinitions

            If writeMissingDefinitions Then
                With missingPhaseDefinitions.getPhaseDef(ix)
                    phName = .name
                    shortName = .shortName
                    darstellungsKlasse = .darstellungsKlasse
                End With
            Else
                With PhaseDefinitions.getPhaseDef(ix)
                    phName = .name
                    shortName = .shortName
                    darstellungsKlasse = .darstellungsKlasse
                End With
            End If
            

            CType(firstrow.Cells(ix, 1), Excel.Range).Offset(1, 0).Value = phName.ToString
            CType(firstrow.Cells(ix, 1), Excel.Range).Offset(1, 5).Value = shortName
            CType(firstrow.Cells(ix, 1), Excel.Range).Offset(1, 6).Value = darstellungsKlasse

            Try
                If phaseAppearanceClasses.Length > 0 Then
                    With CType(firstrow.Cells(ix, 1), Excel.Range).Offset(1, 6)
                        .Validation.Add(Type:=Excel.XlDVType.xlValidateList, AlertStyle:=Excel.XlDVAlertStyle.xlValidAlertStop, _
                                                                               Formula1:=phaseAppearanceClasses)
                    End With
                End If
                
            Catch ex As Exception

            End Try
            
            

        Next ix

        '
        ' jetzt werden die Meilensteine geschrieben 
        '

        ' erste , letzte Zeile des Meilenstein Ranges setzen 
        ' diese Range sollte auf alle Fälle mindestens eine Zeile haben 

        If writeMissingDefinitions Then
            Try
                milestoneDefs = wsName4.Range("Missing_Meilenstein_Definition")
            Catch ex As Exception
                Exit Sub
            End Try

        Else
            milestoneDefs = wsName4.Range("awin_Meilenstein_Definition")
        End If

        anzZeilen = milestoneDefs.Rows.Count
        lastrow = CType(milestoneDefs.Rows(anzZeilen), Excel.Range)
        firstrow = CType(milestoneDefs.Rows(1), Excel.Range)


        ' hier muss erst mal geprüft werden, ob Zeilen eingefügt oder gelöscht werden müssen 
        If writeMissingDefinitions Then
            anzDefinitions = missingMilestoneDefinitions.Count
        Else
            anzDefinitions = MilestoneDefinitions.Count
        End If

        If anzZeilen = anzDefinitions + 2 Then
        ElseIf anzZeilen < anzDefinitions + 2 Then
            ' Zeilen einfügen 

            tmpAnzahl = anzDefinitions + 2 - anzZeilen

            For ix As Integer = 1 To tmpAnzahl
                CType(lastrow.EntireRow, Excel.Range).Insert(Excel.XlInsertShiftDirection.xlShiftDown)
            Next

            ' anzZeilen und phaseDefinitions.count müssen jetzt genau gleich sein 
            anzZeilen = milestoneDefs.Rows.Count


        Else
            ' Zeilen löschen
            tmpAnzahl = anzZeilen - (anzDefinitions + 2)

            For ix As Integer = 1 To tmpAnzahl
                CType(milestoneDefs.Rows(2).EntireRow, Excel.Range).Delete(Excel.XlDeleteShiftDirection.xlShiftUp)
            Next

            ' jetzt sind mindestens zwei Zeilen übrig , und zwar genau dann wenn phaseDefinitions.count = 0 
            anzZeilen = milestoneDefs.Rows.Count

        End If

        ' jetzt können die Meilenstein-Definitions in den Range geschrieben werden 


        For ix As Integer = 1 To anzDefinitions

            If writeMissingDefinitions Then
                With missingMilestoneDefinitions.getMilestoneDef(ix)
                    msName = .name
                    shortName = .shortName
                    darstellungsKlasse = .darstellungsKlasse
                End With
            Else
                With MilestoneDefinitions.getMilestoneDef(ix)
                    msName = .name
                    shortName = .shortName
                    darstellungsKlasse = .darstellungsKlasse
                End With
            End If
            

            CType(firstrow.Cells(ix, 1), Excel.Range).Offset(1, 0).Value = msName.ToString
            CType(firstrow.Cells(ix, 1), Excel.Range).Offset(1, 5).Value = shortName
            CType(firstrow.Cells(ix, 1), Excel.Range).Offset(1, 6).Value = darstellungsKlasse

            Try
                If milestoneAppearanceClasses.Length > 0 Then
                    With CType(firstrow.Cells(ix, 1), Excel.Range).Offset(1, 6)
                        .Validation.Add(Type:=Excel.XlDVType.xlValidateList, AlertStyle:=Excel.XlDVAlertStyle.xlValidAlertStop, _
                                                                               Formula1:=milestoneAppearanceClasses)
                    End With
                End If
                
            Catch ex As Exception

            End Try

        Next ix



        '
        ' Ende der Behandlung der Phasen-/Meilenstein Behandlung 
    End Sub

    ''' <summary>
    ''' liest das Customization File aus und initialisiert die globalen Variablen entsprechend
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub awinsetTypen(ByVal special As String)
        Try
            ' neu 9.11.2016
            Dim formerSU As Boolean = True
            Dim needToBeSaved As Boolean = False
            '  um dahinter temporär die Darstellungsklassen kopieren zu können , nur für ProjectBoard nötig 
            Dim projectBoardSheet As Excel.Worksheet = Nothing

            Dim i As Integer
            Dim xlsCustomization As Excel.Workbook = Nothing

            ReDim importOrdnerNames(8)
            ReDim exportOrdnerNames(8)


            ' Auslesen des Window Namens 
            Dim accountToken As IntPtr = WindowsIdentity.GetCurrent().Token
            Dim myUser As New WindowsIdentity(accountToken)
            myWindowsName = myUser.Name



            globalPath = awinSettings.globalPath


            ' Debug-Mode?
            If awinSettings.visboDebug Then
                If Not IsNothing(globalPath) Then
                    If globalPath.Length > 0 Then
                        Call MsgBox("GlobalPath:" & globalPath & vbLf & _
                                    "existiert: " & My.Computer.FileSystem.DirectoryExists(globalPath).ToString)
                    Else
                        Call MsgBox("GlobalPath: leerer String")
                    End If
                Else
                    Call MsgBox("GlobalPath: Nothing")
                End If


            End If


            ' awinPath kann relativ oder absolut angegeben sein, beides möglich

            Dim curUserDir As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments

            awinPath = My.Computer.FileSystem.CombinePath(curUserDir, awinSettings.awinPath)


            If Not awinPath.EndsWith("\") Then
                awinPath = awinPath & "\"
            End If


            ' Debug-Mode?
            If awinSettings.visboDebug Then
                Call MsgBox("awinPath:" & vbLf & awinPath)
                Call MsgBox("globalPath:" & vbLf & globalPath)

                Call MsgBox("Betriebssystem: " & appInstance.OperatingSystem & Chr(10) & _
                            "Excel-Version: " & appInstance.Version, vbInformation, "Info")
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



            If Not globalPath.EndsWith("\") Then
                globalPath = globalPath & "\"
            End If

            ' Synchronization von Globalen und Lokalen Pfad

            If awinPath <> globalPath And My.Computer.FileSystem.DirectoryExists(globalPath) Then

                If awinSettings.visboDebug Then
                    Call MsgBox("jetzt wird synchronisiert ...")
                End If

                Call synchronizeGlobalToLocalFolder()

            Else

                If awinSettings.visboDebug Then
                    If awinPath = globalPath Then
                        Call MsgBox("awinPath = globalPath: keine Synchronisierung ...")
                    Else
                        Call MsgBox("globalPath existiert nicht: " & vbLf & globalPath)
                    End If

                End If

                If My.Computer.FileSystem.DirectoryExists(awinPath) And (Dir(awinPath, vbDirectory) = "") Then
                    Throw New ArgumentException("Requirementsordner " & awinSettings.awinPath & " existiert nicht")
                End If

            End If


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
            importOrdnerNames(PTImpExp.massenEdit) = awinPath & "Import\massEdit"
            importOrdnerNames(PTImpExp.scenariodefs) = awinPath & "Import\Scenario Definitions"

            exportOrdnerNames(PTImpExp.visbo) = awinPath & "Export\VISBO Steckbriefe"
            exportOrdnerNames(PTImpExp.rplan) = awinPath & "Export\RPLAN-Excel"
            exportOrdnerNames(PTImpExp.msproject) = awinPath & "Export\MSProject"
            exportOrdnerNames(PTImpExp.simpleScen) = awinPath & "Export\einfache Szenarien"
            exportOrdnerNames(PTImpExp.modulScen) = awinPath & "Export\modulare Szenarien"
            exportOrdnerNames(PTImpExp.massenEdit) = awinPath & "Export\massEdit"
            exportOrdnerNames(PTImpExp.scenariodefs) = awinPath & "Export\Scenario Definitions"

            If special = "ProjectBoard" Then

                ' jetzt werden die Directories alle angelegt, sofern Sie nicht schon existieren ... 
                For di As Integer = 0 To importOrdnerNames.Length - 1
                    Try
                        My.Computer.FileSystem.CreateDirectory(importOrdnerNames(di))
                    Catch ex As Exception

                    End Try
                Next

                For di As Integer = 0 To exportOrdnerNames.Length - 1
                    Try
                        My.Computer.FileSystem.CreateDirectory(exportOrdnerNames(di))
                    Catch ex As Exception

                    End Try
                Next

            End If ' if special

            StartofCalendar = StartofCalendar.Date

            LizenzKomponenten(PTSWKomp.ProjectAdmin) = "ProjectAdmin"
            LizenzKomponenten(PTSWKomp.Swimlanes2) = "Swimlanes2"
            LizenzKomponenten(PTSWKomp.Premium) = "Premium"
            LizenzKomponenten(PTSWKomp.SWkomp2) = "SWkomp2"
            LizenzKomponenten(PTSWKomp.SWkomp3) = "SWkomp3"
            LizenzKomponenten(PTSWKomp.SWkomp4) = "SWkomp4"

            ' 14.11.16 tk nicht mehr notwenig , wird in Module initial gesetzt 
            ''ProjektStatus(0) = "geplant"
            ''ProjektStatus(1) = "beauftragt"
            ''ProjektStatus(2) = "beauftragt, Änderung noch nicht freigegeben"
            ''ProjektStatus(3) = "beendet" ' ein Projekt wurde in seinem Verlauf beendet, ohne es plangemäß abzuschliessen
            ''ProjektStatus(4) = "abgeschlossen"


            DiagrammTypen(0) = "Phase"
            DiagrammTypen(1) = "Rolle"
            DiagrammTypen(2) = "Kostenart"
            DiagrammTypen(3) = "Portfolio"
            DiagrammTypen(4) = "Ergebnis"
            DiagrammTypen(5) = "Meilenstein"
            DiagrammTypen(6) = "Meilenstein Trendanalyse"


            Try
                repMessages = XMLImportReportMsg(repMsgFileName, awinSettings.ReportLanguage)
                Call setLanguageMessages()
            Catch ex As Exception

            End Try

            autoSzenarioNamen(0) = "vor Optimierung"
            autoSzenarioNamen(1) = "1. Optimum"
            autoSzenarioNamen(2) = "2. Optimum"
            autoSzenarioNamen(3) = "3. Optimum"

            '
            ' die Namen der Worksheets Ressourcen und Portfolio verfügbar machen
            ' die Zahlen müssen korrespondieren mit der globalen Enumeration ptTables 
            arrWsNames(1) = "repCharts" ' Tabellenblatt zur Aufnahme der Charts für Reports 
            arrWsNames(2) = "Vorlage" ' depr
            ' arrWsNames(3) = 
            arrWsNames(ptTables.MPT) = "MPT"                          ' Multiprojekt-Tafel 
            arrWsNames(4) = "Einstellungen"                ' in Customization File 
            ' arrWsNames(5) = 
            arrWsNames(ptTables.meRC) = "meRC"                          ' Edit Ressourcen
            arrWsNames(6) = "meTE"                          ' Edit Termine
            arrWsNames(7) = "Darstellungsklassen"           ' wird in awinsettypen hinter MPT kopiert; nimmt für die Laufzeit die Darstellungsklassen auf 
            arrWsNames(8) = "Phasen-Mappings"               ' in Customization
            arrWsNames(9) = "meAT"                          ' Edit Attribute 
            arrWsNames(10) = "Meilenstein-Mappings"         ' in Customization
            ' arrWsNames(11) = 
            arrWsNames(ptTables.meCharts) = "meCharts"                     ' Massen-Edit Charts 
            arrWsNames(ptTables.mptPfCharts) = "mptPfCharts"                     ' vorbereitet: Portfolio Charts 
            arrWsNames(ptTables.mptPrCharts) = "mptPrCharts"                     ' vorbereitet: Projekt Charts 
            arrWsNames(14) = "Objekte" ' depr
            arrWsNames(15) = "missing Definitions"          ' in Customization File 


            awinSettings.applyFilter = False

            showRangeLeft = 0
            showRangeRight = 0

            'selectedRoleNeeds = 0
            'selectedCostNeeds = 0


            If special = "ProjectBoard" Then


                '' Versuch, awinsetTypen allgemeingültiger zu machen

                '  bestimmen der maximalen Breite und Höhe 
                formerSU = appInstance.ScreenUpdating
                appInstance.ScreenUpdating = False

                ' 9.11.2016: wird nun ganz am Anfang von awinsetTypen definiert
                '
                '' ''  um dahinter temporär die Darstellungsklassen kopieren zu können  
                ' ''Dim projectBoardSheet As Excel.Worksheet = CType(appInstance.ActiveSheet, _
                ' ''                                        Global.Microsoft.Office.Interop.Excel.Worksheet)
                projectBoardSheet = CType(appInstance.ActiveSheet, _
                                                    Global.Microsoft.Office.Interop.Excel.Worksheet)

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

                miniHeight = maxScreenHeight / 6
                miniWidth = maxScreenWidth / 10



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

                '  With listOfWorkSheets(arrWsNames(4))


                ' Logfile (als ein ExcelSheet) öffnen und ggf. initialisieren

                Call logfileOpen()

                Call logfileSchreiben("Windows-User: ", myWindowsName, anzFehler)


                '' '--------------------------------------------------------------------------------
                '   Testen, ob der User die passende Lizenz besitzt
                '' '--------------------------------------------------------------------------------
                Dim user As String = myWindowsName
                Dim komponente As String = LizenzKomponenten(PTSWKomp.Premium)     ' Lizenz für Projectboard notwendig

                ' Lesen des Lizenzen-Files

                Dim lizenzen As clsLicences = XMLImportLicences(licFileName)

                ' Prüfen der Lizenzen
                If Not lizenzen.validLicence(user, komponente) Then

                    Call logfileSchreiben("Aktueller User " & myWindowsName & " hat keine passende Lizenz", myWindowsName, anzFehler)

                    ''Call MsgBox("Aktueller User " & myWindowsName & " hat keine passende Lizenz!" _
                    ''            & vbLf & " Bitte kontaktieren Sie ihren Systemadministrator")
                    Throw New ArgumentException("Aktueller User " & myWindowsName & " hat keine passende Lizenz!" _
                                & vbLf & " Bitte kontaktieren Sie ihren Systemadministrator")

                End If

                ' Lizenz ist ok


            End If ' if special = "ProjectBoard"

            If special = "BHTC" Or special = "ReportGen" Then

                appInstance = New Excel.Application

                ' hier muss jetzt das Customization File aufgemacht werden ...
                Try
                    xlsCustomization = appInstance.Workbooks.Open(Filename:=awinPath & customizationFile, [ReadOnly]:=True, Editable:=False)
                    myCustomizationFile = appInstance.ActiveWorkbook.Name

                    Call logfileOpen()

                    Call logfileSchreiben("Windows-User: ", myWindowsName, anzFehler)

                Catch ex As Exception
                    Throw New ArgumentException("Customization File nicht gefunden - Abbruch")
                End Try

            ElseIf special = "ProjectBoard" Then

                ' hier muss jetzt das Customization File aufgemacht werden ...
                Try
                    xlsCustomization = appInstance.Workbooks.Open(awinPath & customizationFile)
                    myCustomizationFile = appInstance.ActiveWorkbook.Name
                Catch ex As Exception
                    appInstance.ScreenUpdating = formerSU

                End Try
            Else
                Throw New ArgumentException("Fehler: awinsettypen wurde mit Parameter '" & special & "' aufgerufen!")

            End If

            'Dim wsName4 As Excel.Worksheet = CType(appInstance.Worksheets(arrWsNames(4)), _
            '                                        Global.Microsoft.Office.Interop.Excel.Worksheet)

            Dim wsName4 As Excel.Worksheet = CType(xlsCustomization.Worksheets(arrWsNames(4)), _
                                                    Global.Microsoft.Office.Interop.Excel.Worksheet
                                                    )
            If special = "ProjectBoard" Then

                If awinSettings.databaseURL <> "" And awinSettings.databaseName <> "" Then

                    noDB = False

                    ' ur: 23.01.2015: Abfragen der Login-Informationen
                    loginErfolgreich = loginProzedur()


                    If Not loginErfolgreich Then
                        ' Customization-File wird geschlossen
                        xlsCustomization.Close(SaveChanges:=False)
                        Call logfileSchreiben("LOGIN cancelled ...", "", -1)
                        Call logfileSchliessen()
                        Throw New ArgumentException("LOGIN cancelled ...")

                    End If

                End If
            End If 'if special="ProjectBoard"


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


                ' Auslesen der Custom Field Definitions
                Try
                    Call readCustomFieldDefinitions(wsName4)
                Catch ex As Exception

                End Try

                ' auslesen der anderen Informationen 
                Call readOtherDefinitions(wsName4)


                If special = "ProjectBoard" Then

                    Try
                        ' die Info, welche Sprache gelten soll, ist in ReadOtherDefinitions ...

                        repMessages = XMLImportReportMsg(repMsgFileName, repCult.Name)
                        Call setLanguageMessages()

                    Catch ex As Exception

                    End Try

                    ' sollen die missingDefinitions gelesen / geschrieben werden 

                    If awinSettings.readWriteMissingDefinitions Then
                        Try
                            Dim wsName15 As Excel.Worksheet
                            Try

                                wsName15 = CType(appInstance.Worksheets(arrWsNames(15)), _
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet)

                                ' Auslesen der MissingPhase Definitionen 
                                Call readPhaseDefinitions(wsName15, True)

                                ' Auslesen der Meilenstein Definitionen 
                                Call readMilestoneDefinitions(wsName15, True)
                            Catch ex1 As Exception

                                ' wenn das Sheet nicht existiert, muss es angelegt werden 
                                needToBeSaved = True
                                wsName15 = appInstance.Worksheets.Add(Count:=appInstance.Worksheets.Count + 1)
                                wsName15.Name = arrWsNames(15)
                                With wsName15

                                    Dim tmpRange As Excel.Range = .Range(.Cells(1, 2), .Cells(2, 2))
                                    tmpRange.Offset(0, -1).Value = "unbekannte Phasen-/Vorgangs-Namen"
                                    .Names.Add(Name:="Missing_Phasen_Definition", RefersToR1C1:=tmpRange)

                                    tmpRange = .Range(.Cells(4, 2), .Cells(5, 2))
                                    tmpRange.Offset(0, -1).Value = "unbekannte Meilenstein-Namen"
                                    .Names.Add(Name:="Missing_Meilenstein_Definition", RefersToR1C1:=tmpRange)
                                End With

                            End Try



                        Catch ex As Exception

                        End Try
                    End If

                End If ' if special="ProjectBoard"



                ' hier muss jetzt das Worksheet Phasen-Mappings aufgemacht werden, das ist in arrwsnames(8) abgelegt 
                wsName7810 = CType(appInstance.Worksheets(arrWsNames(8)), _
                                                        Global.Microsoft.Office.Interop.Excel.Worksheet)

                Call readNameMappings(wsName7810, phaseMappings)


                ' hier muss jetzt das Worksheet Milestone-Mappings aufgemacht werden, das ist in arrwsnames(10) abgelegt 
                wsName7810 = CType(appInstance.Worksheets(arrWsNames(10)), _
                                                        Global.Microsoft.Office.Interop.Excel.Worksheet)

                Call readNameMappings(wsName7810, milestoneMappings)

                If special = "ProjectBoard" Then

                    ' jetzt muss die Seite mit den Appearance-Shapes kopiert werden 
                    appInstance.EnableEvents = False
                    CType(appInstance.Workbooks(myCustomizationFile).Worksheets(arrWsNames(7)), _
                    Global.Microsoft.Office.Interop.Excel.Worksheet).Copy(After:=projectBoardSheet)

                    ' hier wird die Datei Projekt Tafel Customizations als aktives workbook wieder geschlossen ....
                    appInstance.Workbooks(myCustomizationFile).Close(SaveChanges:=needToBeSaved) ' ur: 6.5.2014 savechanges hinzugefügt; tk 1.3.16 needtobesaved hinzugefügt
                    appInstance.EnableEvents = True


                    ' jetzt muss die apperanceDefinitions wieder neu aufgebaut werden 
                    appearanceDefinitions.Clear()
                    wsName7810 = CType(appInstance.Workbooks(myProjektTafel).Worksheets(arrWsNames(7)), _
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet)
                    Call aufbauenAppearanceDefinitions(wsName7810)


                    ' jetzt werden die ggf vorhandenen detaillierten Ressourcen Kapazitäten ausgelesen 
                    Call readRessourcenDetails()

                    ' jetzt werden die ggf vorhandenen  Urlaubstage berücksichtigt 
                    Call readRessourcenDetails2()

                    ' Auslesen der Rollen aus der Datenbank ! 
                    Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
                    Dim RoleDefinitions2 As clsRollen = request.retrieveRolesFromDB(Date.Now)
                    Dim costDefinitions2 As clsKostenarten = request.retrieveCostsFromDB(Date.Now)

                    If RoleDefinitions.isIdenticalTo(RoleDefinitions2) And _
                            CostDefinitions.isIdenticalTo(costDefinitions2) Then
                        If awinSettings.visboDebug Then
                            Call MsgBox("es gibt keine Unterschiede in den Rollen / Kosten Definitionen")
                        End If

                        'RoleDefinitions = RoleDefinitions2
                        'CostDefinitions = costDefinitions2
                    Else
                        If awinSettings.visboDebug Then
                            Call MsgBox("es gibt Unterschiede in den Rollen / Kosten Definitionen")
                        End If

                    End If


                    ' jetzt werden die Modul-Vorlagen ausgelesen 
                    Call readVorlagen(True)

                    ' jetzt werden die Projekt-Vorlagen ausgelesen 
                    Call readVorlagen(False)

                    Dim a As Integer = Projektvorlagen.Count
                    Dim b As Integer = ModulVorlagen.Count

                    ' jetzt wird die Projekt-Tafel präpariert - Spaltenbreite und -Höhe
                    ' Beschriftung des Kalenders
                    appInstance.EnableEvents = False
                    Call prepareProjektTafel()


                    projectBoardSheet.Activate()
                    appInstance.EnableEvents = True

                    If Not noDB Then
                        ' jetzt werden aus der Datenbank die Konstellationen und Dependencies gelesen 
                        Call readInitConstellations()

                        currentSessionConstellation.constellationName = calcLastSessionScenarioName()

                    End If

                    ' Logfile wird geschlossen
                    Call logfileSchliessen()

                End If ' if special ="ProjectBoard"

            Catch ex As Exception
                If special = "ProjectBoard" Then
                    appInstance.ScreenUpdating = formerSU
                End If
                appInstance.EnableEvents = True
                Throw New ArgumentException(ex.Message)
            End Try

            ' jetzt werden die windowNames noch gesetzt 


            If awinSettings.englishLanguage Then
                windowNames(PTwindows.mpt) = "VISBO Multiproject-Board"
                windowNames(PTwindows.massEdit) = "edit projects: "
                windowNames(PTwindows.meChart) = "project and portfolio Charts: "
                windowNames(PTwindows.mptpf) = "Portfolio Charts: "
                windowNames(PTwindows.mptpr) = "Project Charts"
            Else
                windowNames(PTwindows.mpt) = "VISBO Multiprojekt-Tafel"
                windowNames(PTwindows.massEdit) = "Projekte editieren: "
                windowNames(PTwindows.meChart) = "Projekt und Portfolio Charts: "
                windowNames(PTwindows.mptpf) = "Portfolio Charts: "
                windowNames(PTwindows.mptpr) = "Projekt Charts"
            End If
            

            projectboardViews(PTview.mpt) = Nothing
            projectboardViews(PTview.mptpr) = Nothing
            projectboardViews(PTview.mptprpf) = Nothing
            projectboardViews(PTview.meOnly) = Nothing
            projectboardViews(PTview.meChart) = Nothing

            projectboardWindows(PTwindows.mpt) = Nothing
            projectboardWindows(PTwindows.mptpr) = Nothing
            projectboardWindows(PTwindows.mptpf) = Nothing
            projectboardWindows(PTwindows.massEdit) = Nothing
            projectboardWindows(PTwindows.meChart) = Nothing


        Catch ex As Exception
            Dim msg As String = ""
            If Not ex.Message.StartsWith("LOGIN cancelled") Then
                ' wird an der aufrufenden Stelle gemacht 
                'Call MsgBox("Fehler in awinsettypen " & special & vbLf & ex.Message)
                msg = "Fehler in awinsettypen " & special & vbLf & ex.Message
            Else
                msg = ex.Message
            End If
            Throw New ArgumentException(msg)
        End Try


    End Sub

    ''' <summary>
    ''' setzt alle angezeigten Projekte, also ShowProjekte,  zurück 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub clearProjectBoard()

        Call awinClearPlanTafel()

        ShowProjekte.Clear()
        projectboardShapes.clear()

        selectedProjekte.Clear(False)
        ImportProjekte.Clear(False)


    End Sub

    ''' <summary>
    ''' setzt die komplette Session zurück 
    ''' löscht alle Shapes, sofern noch welche vorhanden sind, löscht Showprojekte, alleprojekte, etc. 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub clearCompleteSession()

        Dim allShapes As Excel.Shapes
        appInstance.EnableEvents = False
        enableOnUpdate = False

        ' jetzt: Löschen der Session 

        Try

            allShapes = CType(appInstance.ActiveSheet, Excel.Worksheet).Shapes
            For Each element As Excel.Shape In allShapes
                element.Delete()
            Next

        Catch ex As Exception
            Call MsgBox("Fehler beim Löschen der Shapes ...")
        End Try

        ShowProjekte.Clear()
        AlleProjekte.Clear()
        writeProtections.Clear()
        selectedProjekte.Clear(False)
        ImportProjekte.Clear(False)
        DiagramList.Clear()
        awinButtonEvents.Clear()
        projectboardShapes.clear()


        ' es gibt ja nix mehr in der Session 
        currentConstellationName = ""

        ' jetzt werden die temporären Schutz Mechanismen rausgenommen ...
        Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, _
                                   dbUsername, dbPasswort)
        If Request.cancelWriteProtections(dbUsername) Then
            If awinSettings.visboDebug Then
                Call MsgBox("Ihre vorübergehenden Schreibsperren wurden aufgehoben")
            End If
        End If

        
        ' tk, 10.11.16 allDependencies darf nicht gelöscht werden, weil das sonst nicht mehr vorhanden ist
        ' allDependencies wird aktull nur beim Start geladen - und das reicht ja auch ... 
        ' beim Laden eines Szenarios, beim Laden von Projekten wird das nicht mehr geladen ...
        ' auch die geladenen Konstellationen bleiben erhalten 
        ' alternativ könnte das Folgende aktiviert werden ..
        ''allDependencies.Clear()
        ''projectConstellations.Liste.Clear()
        ' '' hier werden jetzt wieder die in der Datenbank vorhandenen Abhängigkeiten und Szenarios geladen ...
        ''Call readInitConstellations()


        ' Löschen der Charts

        Try
            If visboZustaende.projectBoardMode = ptModus.graficboard Then
                Call deleteChartsInSheet(arrWsNames(ptTables.mptPfCharts))
                Call deleteChartsInSheet(arrWsNames(ptTables.mptPrCharts))
                Call deleteChartsInSheet(arrWsNames(ptTables.MPT))
                ' jetzt müssen alle Windows bis auf Window(0) = Multiprojekt-Tafel geschlossen werden 
                ' und mache ProjectboardWindows(mpt) great again ...
                Call closeAllWindowsExceptMPT()

            Else
                Call deleteChartsInSheet(arrWsNames(ptTables.meCharts))
            End If
        Catch ex As Exception
            Dim a As String = ex.Message
        End Try

        ' Session gelöscht

        appInstance.EnableEvents = True
        enableOnUpdate = True
    End Sub

    ''' <summary>
    ''' setzt die Messages je nach Sprache 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub setLanguageMessages()
        'ergebnisChartName(0) = "Earned Value"
        'ergebnisChartName(1) = "Earned Value - gewichtet"
        'ergebnisChartName(2) = "Verbesserungs-Potential"
        'ergebnisChartName(3) = "Risiko-Abschlag"

        ergebnisChartName(0) = repMessages.getmsg(54)
        'Call MsgBox(ergebnisChartName(0))
        ergebnisChartName(1) = repMessages.getmsg(55)
        ergebnisChartName(2) = repMessages.getmsg(56)
        ergebnisChartName(3) = repMessages.getmsg(57)

        ' diese Variablen werden benötigt, um die Diagramme gemäß des gewählten Zeitraums richtig zu positionieren
        summentitel1 = repMessages.getmsg(249)
        summentitel2 = repMessages.getmsg(250)
        summentitel3 = repMessages.getmsg(251)
        summentitel4 = repMessages.getmsg(252)
        summentitel5 = repMessages.getmsg(253)
        summentitel6 = repMessages.getmsg(254)
        summentitel7 = repMessages.getmsg(255)
        summentitel8 = repMessages.getmsg(256)
        summentitel9 = repMessages.getmsg(257)
        summentitel10 = repMessages.getmsg(258)
        summentitel11 = repMessages.getmsg(259)



        ReDim portfolioDiagrammtitel(21)
        'portfolioDiagrammtitel(PTpfdk.Phasen) = "Phasen - Übersicht"
        'portfolioDiagrammtitel(PTpfdk.Rollen) = "Rollen - Übersicht"
        'portfolioDiagrammtitel(PTpfdk.Kosten) = "Kosten - Übersicht"
        'portfolioDiagrammtitel(PTpfdk.ErgebnisWasserfall) = summentitel1
        'portfolioDiagrammtitel(PTpfdk.FitRisiko) = summentitel2
        'portfolioDiagrammtitel(PTpfdk.Auslastung) = summentitel9
        'portfolioDiagrammtitel(PTpfdk.UeberAuslastung) = summentitel10
        'portfolioDiagrammtitel(PTpfdk.Unterauslastung) = summentitel11
        'portfolioDiagrammtitel(PTpfdk.ZieleV) = summentitel6
        'portfolioDiagrammtitel(PTpfdk.ZieleF) = summentitel7
        'portfolioDiagrammtitel(PTpfdk.ComplexRisiko) = "Komplexität, Risiko und Volumen"
        'portfolioDiagrammtitel(PTpfdk.ZeitRisiko) = "Zeit, Risiko und Volumen"
        'portfolioDiagrammtitel(PTpfdk.AmpelFarbe) = ""
        'portfolioDiagrammtitel(PTpfdk.ProjektFarbe) = ""
        'portfolioDiagrammtitel(PTpfdk.Meilenstein) = "Meilenstein - Übersicht"
        'portfolioDiagrammtitel(PTpfdk.FitRisikoVol) = "strategischer Fit, Risiko & Volumen"
        'portfolioDiagrammtitel(PTpfdk.Dependencies) = "Abhängigkeiten: Aktive bzw passive Beeinflussung"
        'portfolioDiagrammtitel(PTpfdk.betterWorseL) = "Abweichungen zum letztem Stand"
        'portfolioDiagrammtitel(PTpfdk.betterWorseB) = "Abweichungen zur Beauftragung"
        'portfolioDiagrammtitel(PTpfdk.Budget) = "Budget Übersicht"
        'portfolioDiagrammtitel(PTpfdk.FitRisikoDependency) = "strategischer Fit, Risiko & Ausstrahlung"

        portfolioDiagrammtitel(PTpfdk.Phasen) = repMessages.getmsg(58)
        portfolioDiagrammtitel(PTpfdk.Rollen) = repMessages.getmsg(59)
        portfolioDiagrammtitel(PTpfdk.Kosten) = repMessages.getmsg(60)
        portfolioDiagrammtitel(PTpfdk.ErgebnisWasserfall) = summentitel1
        portfolioDiagrammtitel(PTpfdk.FitRisiko) = summentitel2
        portfolioDiagrammtitel(PTpfdk.Auslastung) = summentitel9
        portfolioDiagrammtitel(PTpfdk.UeberAuslastung) = summentitel10
        portfolioDiagrammtitel(PTpfdk.Unterauslastung) = summentitel11
        portfolioDiagrammtitel(PTpfdk.ZieleV) = summentitel6
        portfolioDiagrammtitel(PTpfdk.ZieleF) = summentitel7
        portfolioDiagrammtitel(PTpfdk.ComplexRisiko) = repMessages.getmsg(61)
        portfolioDiagrammtitel(PTpfdk.ZeitRisiko) = repMessages.getmsg(62)
        portfolioDiagrammtitel(PTpfdk.AmpelFarbe) = ""
        portfolioDiagrammtitel(PTpfdk.ProjektFarbe) = ""
        portfolioDiagrammtitel(PTpfdk.Meilenstein) = repMessages.getmsg(63)
        portfolioDiagrammtitel(PTpfdk.FitRisikoVol) = repMessages.getmsg(64)
        portfolioDiagrammtitel(PTpfdk.Dependencies) = repMessages.getmsg(65)
        portfolioDiagrammtitel(PTpfdk.betterWorseL) = repMessages.getmsg(66)
        portfolioDiagrammtitel(PTpfdk.betterWorseB) = repMessages.getmsg(67)
        portfolioDiagrammtitel(PTpfdk.Budget) = repMessages.getmsg(68)
        portfolioDiagrammtitel(PTpfdk.FitRisikoDependency) = repMessages.getmsg(69)

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
    Private Sub readPhaseDefinitions(ByVal wsname As Excel.Worksheet, Optional ByVal missingDefinitions As Boolean = False)

        Dim hphase As clsPhasenDefinition
        Dim tmpStr As String = ""

        Try

            With wsname

                Dim phaseRange As Excel.Range

                If missingDefinitions Then
                    Try
                        phaseRange = .Range("Missing_Phasen_Definition")
                    Catch ex As Exception
                        Exit Sub
                    End Try

                Else
                    phaseRange = .Range("awin_Phasen_Definition")
                End If

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
                                    If Not IsNothing(CType(c.Offset(0, 2), Excel.Range).Value) Then
                                        If CStr(c.Offset(0, 2).Value).Trim = "LeLe" Then
                                            specialListofPhases.Add(hphase.name, hphase.name)
                                        End If
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
                                If missingDefinitions Then

                                    missingPhaseDefinitions.Add(hphase)

                                Else

                                    PhaseDefinitions.Add(hphase)

                                End If

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
    Private Sub readMilestoneDefinitions(ByVal wsname As Excel.Worksheet, Optional ByVal missingDefinitions As Boolean = False)

        Dim i As Integer = 0
        Dim hMilestone As clsMeilensteinDefinition
        Dim tmpStr As String


        Try

            With wsname

                Dim milestoneRange As Excel.Range

                If missingDefinitions Then
                    Try
                        milestoneRange = .Range("Missing_Meilenstein_Definition")
                    Catch ex As Exception
                        Exit Sub
                    End Try

                Else
                    milestoneRange = .Range("awin_Meilenstein_Definition")
                End If

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
                                If missingDefinitions Then
                                    missingMilestoneDefinitions.Add(hMilestone)
                                Else
                                    MilestoneDefinitions.Add(hMilestone)
                                End If

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
                            .defaultKapa = CDbl(c.Offset(0, 1).Value)
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

                                .kapazitaet(cp) = .defaultKapa
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
    ''' liest die optional vorhandenen Custom Field Definitionen aus 
    ''' </summary>
    ''' <param name="wsname"></param>
    ''' <remarks></remarks>
    Private Sub readCustomFieldDefinitions(wsname As Excel.Worksheet)

        '
        ' Custom Field Definitions Definitionen auslesen - im bereich awin_CustomField_Definitions
        '

        Try


            With wsname

                Dim customFieldRange As Excel.Range = .Range("awin_CustomField_Definitions")
                Dim anzZeilen As Integer = customFieldRange.Rows.Count
                Dim c As Excel.Range


                For i = 2 To anzZeilen - 1
                    c = CType(customFieldRange.Cells(i, 1), Excel.Range)

                    Dim uid As Integer = i - 1
                    Dim cfType As Integer = -1
                    Dim cfName As String = ""
                    Dim ok As Boolean = False
                    Try
                        cfName = CStr(CType(customFieldRange.Cells(i, 1), Excel.Range).Value)
                        cfType = CInt(CType(customFieldRange.Cells(i, 2), Excel.Range).Value)
                        ok = True
                    Catch ex As Exception

                    End Try

                    If ok And cfName <> "" And isValidCustomField(cfType) Then

                        ' jetzt die CustomField Definition hinzufügen 
                        Try
                            customFieldDefinitions.add(cfName, cfType, uid)
                        Catch ex As Exception
                            Call MsgBox(ex.Message)
                        End Try


                    End If

                Next

            End With

        Catch ex As Exception
            Throw New ArgumentException("Fehler im Customization-File: Custom Field Definitions")
        End Try




    End Sub

    ''' <summary>
    ''' gibt zurück, ob die übergebene Zahl ein gültiger CustomField Typ ist
    ''' </summary>
    ''' <param name="id"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function isValidCustomField(ByVal id As Integer) As Boolean

        If id = ptCustomFields.bool Or _
            id = ptCustomFields.Str Or
            id = ptCustomFields.Dbl Then
            isValidCustomField = True
        Else
            isValidCustomField = False
        End If

    End Function


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
                    ' ''If awinSettings.missingDefinitionColor = XlRgbColor.rgbWhite Then
                    ' ''    Call MsgBox("leeres missingDefinitionColor - Feld in customizationfile " & awinSettings.missingDefinitionColor.ToString)
                    ' ''End If
                Catch ex As Exception

                End Try

                ergebnisfarbe1 = .Range("Ergebnisfarbe1").Interior.Color
                ergebnisfarbe2 = .Range("Ergebnisfarbe2").Interior.Color
                weightStrategicFit = CDbl(.Range("WeightStrategicFit").Value)
                ' jetzt wird KalenderStart, Zeiteinheit und Datenbank Name ausgelesen 
                awinSettings.kalenderStart = CDate(.Range("Start_Kalender").Value)
                awinSettings.zeitEinheit = CStr(.Range("Zeiteinheit").Value)
                awinSettings.kapaEinheit = CStr(.Range("kapaEinheit").Value)
                If awinSettings.kapaEinheit <> "PT" And _
                    awinSettings.kapaEinheit <> "PD" Then
                    awinSettings.kapaEinheit = "PT"
                    Call MsgBox("Kapa-Einheit: Personen-Tage")
                End If
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

            ' gibt es die Einstellung für ProjectWithNoMPmayPass

            Try
                awinSettings.mppProjectsWithNoMPmayPass = CBool(.Range("passFilterWithNoMPs").Value)
            Catch ex As Exception
                awinSettings.mppProjectsWithNoMPmayPass = False
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

            ' Einstellung, um das Lesen / Schreiben von MissingDefinitions zu steuern 
            Try
                awinSettings.readWriteMissingDefinitions = CBool(.Range("RW_MissingDefinitions").Value)
            Catch ex As Exception
                awinSettings.readWriteMissingDefinitions = False
            End Try

            ' Einstellung, um das Lesen / Schreiben von MissingDefinitions zu steuern 
            Try
                awinSettings.meExtendedColumnsView = CBool(.Range("meExtendedView").Value)
            Catch ex As Exception
                awinSettings.meExtendedColumnsView = False
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


            ' sollen im Massen-Edit bei der Berechnung der auslastungsWerte die externen aus der Kapa-Datei mitberücksichtigt werden ? 
            Try
                awinSettings.meAuslastungIsInclExt = CBool(.Range("KapaIstMitExt").Value)
            Catch ex As Exception
                awinSettings.meAuslastungIsInclExt = True
            End Try

            ' welche Sprache soll verwendet werden: wenn english, alles andere ist deutsch
            Try
                awinSettings.englishLanguage = CBool(.Range("englishLanguage").Value)
                If awinSettings.englishLanguage Then
                    menuCult = ReportLang(PTSprache.englisch)
                    repCult = menuCult
                    awinSettings.kapaEinheit = "PD"
                Else
                    awinSettings.kapaEinheit = "PT"
                    menuCult = ReportLang(PTSprache.deutsch)
                    repCult = menuCult
                End If
            Catch ex As Exception
                awinSettings.englishLanguage = False
                awinSettings.kapaEinheit = "PT"
                menuCult = ReportLang(PTSprache.deutsch)
                repCult = menuCult
            End Try

            ' sollen Sammelrollen immer nur in Summe dargestellt werden, oder aufgeteilt in Platzhalter / Assigned 
            Try
                awinSettings.showPlaceholderAndAssigned = CBool(.Range("ShowPlaceHolderAndAssigned").Value)
            Catch ex As Exception
                awinSettings.showPlaceholderAndAssigned = False
            End Try

            ' sollen die Risiko Kennzahlen bei der Berechnung der Portfolio / Projekt-Ergebnisse mitgerechnet werden ?  
            Try
                awinSettings.considerRiskFee = CBool(.Range("considerRiskFee").Value)
            Catch ex As Exception
                awinSettings.considerRiskFee = False
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
    ''' liest für die definierten Rollen ggf vorhandene Urlaubsplanung ein 
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub readRessourcenDetails2()

        Dim kapaFileName As String
        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim formerSU As Boolean = appInstance.ScreenUpdating
        Dim listOfFiles As Collections.ObjectModel.ReadOnlyCollection(Of String) = Nothing

        If formerEE Then
            appInstance.EnableEvents = False
        End If

        If formerSU Then
            appInstance.ScreenUpdating = False
        End If

        enableOnUpdate = False

        kapaFileName = "Urlaubsplaner*.xlsx"

        ' Dateien mit WildCards lesen
        listOfFiles = My.Computer.FileSystem.GetFiles(awinPath & projektRessOrdner,
                     FileIO.SearchOption.SearchTopLevelOnly, kapaFileName)

        ''listOfFiles = My.Computer.FileSystem.GetFiles(awinPath & projektRessOrdner,
        ''              FileIO.SearchOption.SearchTopLevelOnly, "Urlaubsplaner*.xlsx")

        If listOfFiles.Count = 1 Then
            Call readUrlOfRole(listOfFiles.Item(0))
        ElseIf listOfFiles.Count = 0 Then
            Call logfileSchreiben("Es gibt keine Datei zur Urlaubsplanung" & vbLf _
                         & "Es wurde daher jetzt keine berücksichtigt", _
                         "", anzFehler)
        Else

            Call logfileSchreiben("Es gibt mehrere Dateien zur Urlaubsplanung" & vbLf _
                         & "Es wurde daher jetzt keine berücksichtigt", _
                         "", anzFehler)
        End If

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
    ''' setzt Kalenderleiste und Spaltenbreite sowie -Höhe 
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub prepareProjektTafel()


        ' bestimmen der Spaltenbreite und Spaltenhöhe ...
        Dim testCase As String = appInstance.ActiveWorkbook.Name
        If testCase <> myProjektTafel Then
            CType(appInstance.Workbooks(myProjektTafel), Excel.Workbook).Activate()
        End If

        Dim wsName3 As Excel.Worksheet = CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.MPT)), _
                                                Global.Microsoft.Office.Interop.Excel.Worksheet)

        Dim tmpRange As Excel.Range
        Dim tempWSName As String = CType(appInstance.ActiveSheet, Excel.Worksheet).Name

        Dim tmpStart As Date
        Try
            With wsName3
                Dim rng As Excel.Range
                'Dim colDate As date
                If awinSettings.zeitEinheit = "PM" Then
                    ' die Kalender-Leiste schreiben 
                    CType(.Cells(1, 1), Global.Microsoft.Office.Interop.Excel.Range).Value = StartofCalendar
                    CType(.Cells(1, 2), Global.Microsoft.Office.Interop.Excel.Range).Value = StartofCalendar.AddMonths(1)
                    rng = .Range(.Cells(1, 1), .Cells(1, 2))
                    '' Deutsches Format:
                    'rng.NumberFormat = "[$-407]mmm yy;@"
                    ' Englische Format:
                    rng.NumberFormat = "[$-409]mmm yy;@"

                    Dim destinationRange As Excel.Range = .Range(.Cells(1, 1), .Cells(1, 720))
                    With destinationRange
                        .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                        .VerticalAlignment = Excel.XlVAlign.xlVAlignBottom
                        '' Deutsches Format: 
                        'rng.NumberFormat = "[$-407]mmm yy;@"
                        ' Englische Format:
                        .NumberFormat = "[$-409]mmm yy;@"
                        .WrapText = False
                        .Orientation = 90
                        .AddIndent = False
                        .IndentLevel = 0
                        ' Änderung tk 14.11 - sonst können ja die Spaltenbreiten ud Höhen nicht explizit gesetzt werden 
                        ' das ist vor allem auf der Zeichenfläche notwendig, weil sonst die Berechnung und Positionierung der Grafik Elemente nicht mehr stimmt 
                        .ShrinkToFit = True
                        .ReadingOrder = Excel.Constants.xlContext
                        .MergeCells = False
                        .Interior.Color = noshowtimezone_color
                    End With

                    rng.AutoFill(Destination:=destinationRange, Type:=Excel.XlAutoFillType.xlFillMonths)

                ElseIf awinSettings.zeitEinheit = "PW" Then
                    For i As Integer = 1 To 210
                        CType(.Cells(1, i), Global.Microsoft.Office.Interop.Excel.Range).Value = StartofCalendar.AddDays((i - 1) * 7)
                    Next
                ElseIf awinSettings.zeitEinheit = "PT" Then
                    Dim workOnSat As Boolean = False
                    Dim workOnSun As Boolean = False


                    If Weekday(StartofCalendar, FirstDayOfWeek.Monday) > 3 Then
                        tmpStart = StartofCalendar.AddDays(8 - Weekday(StartofCalendar, FirstDayOfWeek.Monday))
                    Else
                        tmpStart = StartofCalendar.AddDays(Weekday(StartofCalendar, FirstDayOfWeek.Monday) - 8)
                    End If
                    '
                    ' jetzt ist tmpstart auf Montag ... 
                    Dim tmpDay As Date
                    Dim i As Integer = 1

                    For w As Integer = 1 To 30
                        For d As Integer = 0 To 4
                            ' das sind Montag bis Freitag
                            tmpDay = tmpStart.AddDays(d)
                            If Not feierTage.Contains(tmpDay) Then
                                CType(.Cells(1, i), Global.Microsoft.Office.Interop.Excel.Range).Value = tmpDay.ToString("d")
                                i = i + 1
                            End If
                        Next
                        tmpDay = tmpStart.AddDays(5)
                        If workOnSat Then
                            CType(.Cells(1, i), Global.Microsoft.Office.Interop.Excel.Range).Value = tmpDay.ToString("d")
                            i = i + 1
                        End If
                        tmpDay = tmpStart.AddDays(6)
                        If workOnSun Then
                            CType(.Cells(1, i), Global.Microsoft.Office.Interop.Excel.Range).Value = tmpDay.ToString("d")
                            i = i + 1
                        End If
                        tmpStart = tmpStart.AddDays(7)
                    Next


                End If


                ' hier werden jetzt die Spaltenbreiten und Zeilenhöhen gesetzt 

                Dim maxRows As Integer = .Rows.Count
                Dim maxColumns As Integer = .Columns.Count

                tmpRange = CType(.Rows(1), Global.Microsoft.Office.Interop.Excel.Range)
                CType(.Rows(1), Global.Microsoft.Office.Interop.Excel.Range).RowHeight = awinSettings.zeilenhoehe1
                CType(.Range(.Cells(2, 1), .Cells(maxRows, maxColumns)), Global.Microsoft.Office.Interop.Excel.Range).RowHeight = awinSettings.zeilenhoehe2
                CType(.Columns, Global.Microsoft.Office.Interop.Excel.Range).ColumnWidth = awinSettings.spaltenbreite

                With CType(.Range(.Cells(2, 1), .Cells(maxRows, maxColumns)), Global.Microsoft.Office.Interop.Excel.Range)
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                    .NumberFormat = "####0"
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    ' Änderung tk 14.11 - sonst können ja die Spaltenbreiten ud Höhen nicht explizit gesetzt werden 
                    ' das ist vor allem auf der Zeichenfläche notwendig, weil sonst die Berechnung und Positionierung der Grafik Elemente nicht mehr stimmt 
                    .ShrinkToFit = True
                    .ReadingOrder = Excel.Constants.xlContext
                    .MergeCells = False
                End With

                boxWidth = CDbl(CType(.Cells(3, 3), Global.Microsoft.Office.Interop.Excel.Range).Width)
                boxHeight = CDbl(CType(.Cells(3, 3), Global.Microsoft.Office.Interop.Excel.Range).Height)

                topOfMagicBoard = CDbl(CType(.Cells(1, 1), Global.Microsoft.Office.Interop.Excel.Range).Height) + 0.1 * boxHeight
                screen_correct = 0.1 * 19.3 / boxWidth


                Dim laenge As Integer
                laenge = showRangeRight - showRangeLeft

                If laenge > 0 And showRangeLeft > 0 Then
                    .Range(.Cells(1, showRangeLeft), .Cells(1, showRangeLeft + laenge)).Interior.Color = showtimezone_color
                End If

            End With
        Catch ex As Exception

        End Try

    End Sub

    ''' <summary>
    ''' liest die Konstellationen und Abhängigkeiten in der Datenbank 
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub readInitConstellations()

        Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)

        ' Datenbank ist gestartet
        If request.pingMongoDb() Then

            ' alle Konstellationen laden 
            projectConstellations = request.retrieveConstellationsFromDB()

            ' hier werden jetzt auch alle Abhängigkeiten geladen 
            allDependencies = request.retrieveDependenciesFromDB()

            Dim axt As Integer = 9

        Else
            Throw New ArgumentException("Datenbank - Verbindung ist unterbrochen ...")
        End If

    End Sub
    '
    '
    '
    Public Sub awinChangeTimeSpan(ByVal von As Integer, ByVal bis As Integer, _
                                  Optional ByVal noFurtherActions As Boolean = False)

        'Dim k As Integer

        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim formerSU As Boolean = appInstance.ScreenUpdating
        Dim noTimeFrame As Boolean = False

        appInstance.EnableEvents = False



        If von < 1 Then
            von = 1
        End If

        If bis < von + minColumns - 1 Then
            noTimeFrame = True
        End If

        ' damit es nicht flackert, wenn zweimal hintereinander Zeitzone aufgehoben wird  
        If noTimeFrame Then
            If showRangeRight <> showRangeLeft Then
                appInstance.ScreenUpdating = False
            End If
        Else
            appInstance.ScreenUpdating = False
        End If


        If showRangeLeft <> von Or showRangeRight <> bis Or _
            AlleProjekte.Count = 0 Then
            '
            ' wenn roentgenblick.ison , werden Bedarfe angezeigt - die müssen hier ausgeblendet werden - nachher mit den neuen Werten eingeblendet werden
            '

            If roentgenBlick.isOn And ShowProjekte.Count > 0 Then
                Call awinNoshowProjectNeeds()
            End If



            '
            ' aktualisieren der Showtime zone, erst die alte ausblenden , dann die neue einblenden
            '
            Call awinShowtimezone(showRangeLeft, showRangeRight, False)

            If noTimeFrame Then

                showRangeLeft = 0
                showRangeRight = 0

                ' jetzt prüfen, ob Röntgenblick an ist
                ' wenn ja: ausscalten 
                With roentgenBlick
                    If roentgenBlick.isOn Then
                        .isOn = False
                        .name = ""
                        .myCollection = Nothing
                        .type = ""
                        Call awinNoshowProjectNeeds()
                    End If
                End With


            Else
                Call awinShowtimezone(von, bis, True)


                showRangeLeft = von
                showRangeRight = bis


                ' jetzt werden - falls nötig die Projekte nachgeladen ... 
                Try
                    If awinSettings.applyFilter Then
                        ' vorher hiess das loadprojectsonChange - jetzt ist es so: 
                        ' wenn applyFilter = true, dann soll nachgeladen werden unter Anwendung 
                        ' des Filters "Last"
                        Dim filter As New clsFilter
                        filter = filterDefinitions.retrieveFilter("Last")
                        Call awinProjekteImZeitraumLaden(awinSettings.databaseName, filter)

                        '' jetzt sind wieder alle Projekte des Zeitraums da - deswegen muss nicht ggf nachgeladen werden 
                        'DeletedProjekte.Clear()

                        '
                        '   wenn "selectedRoleNeeds" ungleich Null ist, werden Bedarfe angezeigt - die müssen hier wieder - mit den neuen Daten für show_range_lefet, .._right eingeblendet werden
                        '
                        If roentgenBlick.isOn Then
                            With roentgenBlick
                                Call awinShowProjectNeeds1(mycollection:=.myCollection, type:=.type)
                            End With
                        End If



                        '
                        ' wenn diagramme angezeigt sind - aktualisieren dieser Diagramme
                        '



                    End If

                    ' betrachteter Zeitraum wurde geändert - typus = 4
                    Call awinNeuZeichnenDiagramme(4)


                Catch ex As Exception
                    Call MsgBox(ex.Message)
                End Try
            End If






        End If



        appInstance.EnableEvents = formerEE
        If appInstance.ScreenUpdating <> formerSU Then
            appInstance.ScreenUpdating = formerSU
        End If



    End Sub

    ''' <summary>
    '''speziell auf BMW Rplan Output angepasstes Inventur Import File 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub bmwImportProjektInventur(ByRef myCollection As Collection)

        Dim zeile As Integer, spalte As Integer
        Dim pName As String = " "

        Dim lastRow As Integer

        Dim hproj As clsProjekt
        Dim vproj As clsProjektvorlage
        Dim geleseneProjekte As Integer
        Dim projektFarbe As Object
        Dim anfang As Integer, ende As Integer
        Dim cphase As clsPhase
        Dim cresult As clsMeilenstein
        Dim cbewertung As clsBewertung
        Dim ix As Integer
        Dim tmpStr(20) As String
        Dim aktuelleZeile As String
        Dim nameSopTyp As String = " "
        Dim nameBU As String
        Dim sopDate As Date
        Dim tmpStartSop As Date ' wird benutzt , um eine Hilfsphase zu machen 
        Dim startDate As Date, endDate As Date
        Dim startoffset As Long, duration As Long
        Dim vorlagenName As String
        Dim phaseName As String
        Dim itemName As String
        Dim zufall As New Random(10)
        Dim farbKennung As Integer
        Dim responsible As String


        ' Vorbedingung: das Excel File. das importiert werden soll , ist bereits geöffnet 

        zeile = 2
        spalte = 1
        geleseneProjekte = 0


        Try
            'Dim activeWSListe As Excel.Worksheet = CType(appInstance.ActiveWorkbook.Worksheets("Tabelle1"), _
            '                                                Global.Microsoft.Office.Interop.Excel.Worksheet)
            Dim activeWSListe As Excel.Worksheet = CType(appInstance.ActiveWorkbook.ActiveSheet, _
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet)
            With activeWSListe

                Dim tstStr As String
                Try
                    tstStr = CStr(CType(activeWSListe.Cells(2, 1), Excel.Range).Value)
                    projektFarbe = CType(activeWSListe.Cells(2, 1), Excel.Range).Interior.Color
                Catch ex As Exception
                    projektFarbe = CType(activeWSListe.Cells(2, 1), Excel.Range).Interior.ColorIndex
                End Try

                ' hier werden jetzt die Columns bestimmt 


                lastRow = System.Math.Max(CType(.Cells(2000, 1), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlUp).Row, _
                                          CType(.Cells(2000, 2), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlUp).Row)

                While zeile <= lastRow

                    anfang = zeile + 1
                    ix = anfang


                    Do While CBool((CType(.Cells(ix, 1), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color IsNot projektFarbe)) And (ix <= lastRow)
                        ix = ix + 1
                    Loop

                    ende = ix - 1

                    ' hier wird Name, Typ, SOP, Business Unit, vname, Start-Datum, Dauer der Phase(1) ausgelesen  
                    aktuelleZeile = CStr(CType(activeWSListe.Cells(zeile, 2), Excel.Range).Value).Trim
                    startDate = CDate(CType(activeWSListe.Cells(zeile, 3), Excel.Range).Value)
                    endDate = CDate(CType(activeWSListe.Cells(zeile, 4), Excel.Range).Value)
                    farbKennung = CInt(CType(activeWSListe.Cells(zeile, 12), Excel.Range).Value)
                    responsible = CStr(CType(activeWSListe.Cells(zeile, 9), Excel.Range).Value)


                    duration = DateDiff(DateInterval.Day, startDate, endDate) + 1
                    If duration < 0 Then
                        startDate = endDate
                        duration = -1 * duration
                        endDate = startDate.AddDays(duration)
                    End If

                    tmpStr = aktuelleZeile.Trim.Split(New Char() {CChar("["), CChar("]")}, 5)


                    nameSopTyp = tmpStr(0).Trim
                    If Not isValidProjectName(nameSopTyp) Then
                        nameSopTyp = makeValidProjectName(nameSopTyp)
                    End If
                    pName = nameSopTyp
                    Try
                        nameBU = tmpStr(1)
                        tmpStr = nameBU.Split(New Char() {CChar(" ")}, 3)
                        nameBU = tmpStr(0)
                    Catch ex1 As Exception
                        nameBU = ""
                    End Try


                    Dim foundIX As Integer = -1

                    tmpStr = nameSopTyp.Trim.Split(New Char() {CChar(" ")}, 15)
                    Dim k As Integer = 0

                    Do While foundIX < 0 And k <= tmpStr.Length - 2
                        If tmpStr(k).Trim = "SOP" And k < tmpStr.Length - 1 Then
                            Try
                                sopDate = CDate(tmpStr(k + 1)).AddMonths(1).AddDays(-1)
                                tmpStartSop = CDate(tmpStr(k + 1))
                            Catch ex As Exception
                                Dim tmp1Str(3) As String
                                tmp1Str = tmpStr(k + 1).Split(New Char() {CChar("/")}, 8)

                                If CInt(tmp1Str(1)) < 50 Then
                                    tmp1Str(1) = CStr(2000 + CInt(tmp1Str(1)))
                                End If
                                tmpStr(k + 1) = tmp1Str(0) & "-" & tmp1Str(1)
                                sopDate = CDate(tmpStr(k + 1)).AddMonths(1).AddDays(-1)
                                tmpStartSop = CDate(tmpStr(k + 1))
                            End Try

                            foundIX = k + 2
                        Else
                            k = k + 1
                        End If
                    Loop

        If foundIX < 0 Then
            ' SOP Date konnte nicht bestimmt werden 
            sopDate = endDate
            tmpStartSop = sopDate.AddDays(-28)
            foundIX = tmpStr.Length - 1
        End If

        Select Case tmpStr(foundIX).Trim
            Case "eA"
                vorlagenName = "Enge Ableitung"
            Case "wA"
                vorlagenName = "Weite Ableitung"
            Case "E"
                vorlagenName = "Erstanläufer"
            Case Else
                vorlagenName = "Erstanläufer"
        End Select

        '
        ' jetzt wird das Projekt angelegt 
        '
        hproj = New clsProjekt

        Try
            vproj = Projektvorlagen.getProject(vorlagenName)


            hproj.farbe = vproj.farbe
            hproj.Schrift = vproj.Schrift
            hproj.Schriftfarbe = vproj.Schriftfarbe
            hproj.name = ""
            hproj.VorlagenName = vorlagenName
            hproj.earliestStart = vproj.earliestStart
            hproj.latestStart = vproj.latestStart
            hproj.ampelStatus = farbKennung
            hproj.leadPerson = responsible

        Catch ex As Exception
            Throw New Exception("es gibt keine entsprechende Vorlage mit Namen  " & vorlagenName & vbLf & ex.Message)
        End Try


        Try

            hproj.name = pName
            hproj.startDate = startDate
            hproj.earliestStartDate = hproj.startDate.AddMonths(hproj.earliestStart)
            hproj.latestStartDate = hproj.startDate.AddMonths(hproj.latestStart)
            ' immer als beauftragtes PRojekt importieren 
            hproj.Status = ProjektStatus(1)
            'If DateDiff(DateInterval.Month, startDate, Date.Now) <= 0 Then
            '    hproj.Status = ProjektStatus(0)
            'Else
            '    hproj.Status = ProjektStatus(1)
            'End If

            hproj.StrategicFit = zufall.NextDouble * 10
            hproj.Risiko = zufall.NextDouble * 10
            hproj.volume = zufall.NextDouble * 1000000
            hproj.complexity = zufall.NextDouble
            hproj.businessUnit = nameBU
            hproj.description = nameSopTyp

            hproj.Erloes = 0.0


        Catch ex As Exception
            Throw New Exception("in erstelle InventurProjekte: " & vbLf & ex.Message)
        End Try

        ' jetzt werden all die Phasen angelegt , beginnend mit der ersten 
        cphase = New clsPhase(parent:=hproj)
        cphase.nameID = rootPhaseName
        startoffset = 0
        duration = DateDiff(DateInterval.Day, startDate, endDate) + 1
        cphase.changeStartandDauer(startoffset, duration)

        cresult = New clsMeilenstein(parent:=cphase)
        cresult.nameID = calcHryElemKey("SOP", True)
        cresult.setDate = sopDate

        cbewertung = New clsBewertung
        cbewertung.colorIndex = farbKennung
        cbewertung.description = " .. es wurde  keine Erläuterung abgegeben .. "
        cresult.addBewertung(cbewertung)

        Try
            cphase.addMilestone(cresult)
        Catch ex As Exception

        End Try


        hproj.AddPhase(cphase)


        Dim phaseIX As Integer = PhaseDefinitions.Count + 1


        Dim pStartDate As Date
        Dim pEndDate As Date
        Dim ok As Boolean = True
        Dim lastPhaseName As String = cphase.nameID

        Dim i As Integer
        For i = anfang To ende

            Try
                itemName = CStr(CType(.Cells(i, 2), Excel.Range).Value).Trim
            Catch ex As Exception
                itemName = ""
                ok = False
            End Try

            If ok Then

                pStartDate = CDate(CType(.Cells(i, 3), Excel.Range).Value)
                pEndDate = CDate(CType(.Cells(i, 4), Excel.Range).Value)
                startoffset = DateDiff(DateInterval.Day, hproj.startDate, pStartDate)
                duration = DateDiff(DateInterval.Day, pStartDate, pEndDate) + 1

                If duration > 1 Then
                    ' es handelt sich um eine Phase 
                    phaseName = itemName
                    cphase = New clsPhase(parent:=hproj)
                    cphase.nameID = hproj.hierarchy.findUniqueElemKey(phaseName, False)

                    If PhaseDefinitions.Contains(phaseName) Then
                        ' nichts tun 
                    Else
                        ' in die Phase-Definitions aufnehmen 

                        Dim hphase As clsPhasenDefinition
                        hphase = New clsPhasenDefinition

                        'hphase.farbe = CLng(CType(.Cells(i, 1), Excel.Range).Interior.Color)
                        hphase.name = phaseName
                        hphase.UID = phaseIX
                        phaseIX = phaseIX + 1

                        Try
                            PhaseDefinitions.Add(hphase)
                        Catch ex As Exception

                        End Try

                    End If

                    cphase.changeStartandDauer(startoffset, duration)
                    hproj.AddPhase(cphase)
                    lastPhaseName = cphase.nameID

                ElseIf duration = 1 Then

                    Try
                        ' es handelt sich um einen Meilenstein 

                        Dim bewertungsAmpel As Integer
                        Dim explanation As String

                        bewertungsAmpel = CInt(CType(.Cells(i, 12), Excel.Range).Value)
                        explanation = CStr(CType(.Cells(i, 1), Excel.Range).Value)

                        cphase = hproj.getPhaseByID(lastPhaseName)
                        cresult = New clsMeilenstein(parent:=cphase)
                        cbewertung = New clsBewertung



                        If bewertungsAmpel < 0 Or bewertungsAmpel > 3 Then
                            ' es gibt keine Bewertung
                            bewertungsAmpel = 0
                        End If

                        ' damit Kriterien auch eingelesen werden, wenn noch keine Bewertung existiert ...
                        With cbewertung
                            '.bewerterName = resultVerantwortlich
                            .colorIndex = bewertungsAmpel
                            .datum = Date.Now
                            .description = explanation
                        End With

                        With cresult
                            .nameID = hproj.hierarchy.findUniqueElemKey(itemName, True)
                            .setDate = pEndDate
                            If Not cbewertung Is Nothing Then
                                .addBewertung(cbewertung)
                            End If
                        End With

                        Try
                            With cphase
                                .addMilestone(cresult)
                            End With
                        Catch ex As Exception

                        End Try

                    Catch ex As Exception

                    End Try




                End If




                ' handelt es sich um eine Phase oder um einen Meilenstein ? 


            End If


        Next


        ' jetzt muss das Projekt eingetragen werden 
        ImportProjekte.Add(hproj, False)
        myCollection.Add(hproj.name)


        zeile = ende + 1

        Do While CBool(CType(.Cells(zeile, 1), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color IsNot projektFarbe) And zeile <= lastRow
            zeile = zeile + 1
        Loop

                End While





            End With
        Catch ex As Exception
            Throw New Exception("Fehler in Datei BMW Projekt-Inventur " & vbLf & ex.Message & vbLf & pName)
        End Try



    End Sub
    Sub awinImportMSProject(ByVal modus As String, ByVal filename As String, ByRef hproj As clsProjekt, ByRef importdate As Date)

        Dim prj As MSProject.Application
        Dim msproj As MSProject.Project
        Dim i As Integer = 1
        Dim lastphase As clsPhase
        Dim lasthrchyNode As clsHierarchyNode
        Dim lastelemID As String = ""
        Dim lastlevel As Integer = 0
        Dim Xwerte() As Double
        Dim active_proj As String = ""      ' Name des aktuell aktiven Projektes

        ' hier wird eingetragen, welches vordefinierte Flag das customized Field VISBO usw. repräsentiert
        Dim visboflag As MSProject.PjField = Nothing
        Dim visbo_taskclass As MSProject.PjField = Nothing
        Dim visbo_abbrev As MSProject.PjField = Nothing
        Dim visbo_ampel As MSProject.PjField = Nothing

        ' Liste, die aufgebaut wird beim Einlesen der Tasks. Hier wird vermerkt, welche Task das Visbo-Flag mit YES und welche mit NO
        ' gesetzt hat d.h. berücksichtigt werden soll
        ' Diese Liste enthält keine Elemente, wenn das VISBO-Flag nicht definiert ist
        Dim visboFlagListe As New SortedList(Of String, Boolean)

        Try

            'On Error Resume Next
            Try
                prj = CType(GetObject(, "msproject.application"), MSProject.Application)
            Catch ex As Exception
                prj = CType(CreateObject("msproject.application"), MSProject.Application)

                If IsNothing(prj) Then
                    Call MsgBox("MSproject ist nicht installiert")
                    Exit Sub
                End If
            End Try

            If modus <> "BHTC" Then

                ' ''prj.FileOpen(Name:="\\KOYTEK-NAS\backup\Ute\VISBO\MS Project Beispiele\ute.mpp", _
                ' ''             ReadOnly:=True, FormatID:="MSProject.MPP")

                prj.FileOpen(Name:=filename, _
                            ReadOnly:=True, FormatID:="MSProject.MPP")


            End If


            Dim anzProj As Integer = prj.Projects.Count

            If anzProj > 0 Then


                ' VISBO-Flag dient dazu, Tasks, die nicht benötigt werden in der MultiprojektPlanung nicht mit einzulesen
                ' in die Projekt-Tafel

                ' Ist dieses VISBO-Flag definiert?
                Dim pjFlag As String = ""

                Try
                    visboflag = CType(prj.FieldNameToFieldConstant("VISBO", MSProject.PjFieldType.pjTask), MSProject.PjField)
                    pjFlag = prj.FieldConstantToFieldName(visboflag)

                Catch ex As Exception
                    visboflag = 0
                End Try

                Try
                    visbo_taskclass = CType(prj.FieldNameToFieldConstant(awinSettings.visboTaskClass, MSProject.PjFieldType.pjTask), MSProject.PjField)
                Catch ex As Exception
                    visbo_taskclass = 0
                End Try
                Try
                    visbo_abbrev = CType(prj.FieldNameToFieldConstant(awinSettings.visboAbbreviation, MSProject.PjFieldType.pjTask), MSProject.PjField)
                Catch ex As Exception
                    visbo_abbrev = 0
                End Try
                Try
                    visbo_ampel = CType(prj.FieldNameToFieldConstant(awinSettings.visboAmpel, MSProject.PjFieldType.pjTask), MSProject.PjField)
                Catch ex As Exception
                    visbo_ampel = 0
                End Try

                If modus = "BHTC" Then
                    ' In Missing..Definitions sind noch die Definitionen des vorausgegangenen Projekts definiert.
                    ' Diese sollen nicht mehr aktiv sein.
                    missingPhaseDefinitions.Clear()
                    missingMilestoneDefinitions.Clear()
                    '' Einlesen des aktiven Projekts
                    msproj = prj.ActiveProject
                Else
                    '' Einlesen des zuletzt gelesenen Projekts
                    msproj = prj.Projects.Item(anzProj)

                End If

                ' '' '' Einlesen der diversen Projekte, die geladen wurden (gilt nur für BHTC), sonst immer nur das zuletzt geladene
                '' ''For proj_i = beginnProjekt To endeProjekt



                hproj = New clsProjekt(CDate(msproj.ProjectStart), CDate(msproj.ProjectStart), CDate(msproj.ProjectStart))

                Dim ProjektdauerIndays As Integer = calcDauerIndays(hproj.startDate, CDate(msproj.Finish))
                Dim startOffset As Long = DateDiff(DateInterval.Day, hproj.startDate, hproj.startDate.AddMonths(0))

                ' Projektname ohne "."
                Dim hhstr() As String
                hhstr = Split(msproj.Name, ".", -1)
                ' alle evtl auftretenden #, (, ) werden ersetzt durch unkritische Zeichen ... 
                hproj.name = makeValidProjectName(hhstr(0))
                'hproj.idauer = DateDiff(DateInterval.Month, CType(msproj.DefaultFinishTime, Date), CType(msproj.DefaultStartTime, Date))

                '' '' merken für BHTC, da hier der Report für das aktive Projekt gemacht werden soll 
                ' ''If prj.ActiveProject.Name = msproj.Name Then
                ' ''    active_proj = hproj.name
                ' ''End If

                Dim anzSubprojects As Integer = msproj.Subprojects.Count

                hproj.description = msproj.ProjectNotes
                hproj.UID = msproj.UniqueID
                Dim hrsPerDay As Double = msproj.HoursPerDay

                Dim projUID As Object = msproj.DatabaseProjectUniqueID

                ' ------------------------------------------------------------------------------------------------------
                ' Erzeugen und eintragen der Projekt-Phase (= erste Phase mit Dauer des Projekts)
                ' ------------------------------------------------------------------------------------------------------
                Try
                    Dim cphase As New clsPhase(hproj)

                    ' ProjektPhase wird erzeugt
                    cphase = New clsPhase(parent:=hproj)
                    cphase.nameID = rootPhaseName

                    ' Phasen Dauer wird gleich der Dauer des Projekts gesetzt
                    With cphase
                        .nameID = rootPhaseName
                        Dim cphaseStartOffset As Integer = 0
                        .changeStartandDauer(cphaseStartOffset, ProjektdauerIndays)
                    End With
                    ' rootPhaseName - Phase wird hinzugefügt
                    hproj.AddPhase(cphase)

                Catch ex1 As Exception
                    Throw New ArgumentException("Fehler in awinImportMSProject, Erzeugen ProjektPhase")
                End Try




                Dim anzTasks As Integer = msproj.Tasks.Count
                anzTasks = msproj.NumberOfTasks


                Dim resPool As MSProject.Resources = msproj.Resources

                Dim res(resPool.Count) As Object
                For i = 1 To resPool.Count
                    res(i) = resPool.Item(i)
                Next


                For i = 1 To anzTasks

                    Dim msTask As MSProject.Task

                    Dim cphase As New clsPhase(parent:=hproj)


                    msTask = msproj.Tasks.Item(i)




                    ' hier: evt. Prüfung ob eine VISBO Projekt-Tafel relevante Task
                    ' oder: ob eine Task auf dem kritischen Pfad liegt

                    ' Wenn sumTask = true, dann ist die aktuelle Task eine Summary-Task
                    Dim sumTask As Boolean = CType(msTask.Summary, Boolean)

                    ' Herausfinden der Hierarchiestufe
                    Dim hstr() As String = Split(msTask.WBS, ".", -1)
                    Dim tasklevel As Integer = hstr.Count


                    ' hier muss der Uniquename(ID) erzeugt werden evt. aus PhaseDefinitions

                    If Not CType(msTask.Milestone, Boolean) _
                        Or _
                        (CType(msTask.Milestone, Boolean) And CType(msTask.Summary, Boolean)) Then

                        ' Ergänzung tk für Demo BHTC 
                        ' falls Synonyme definiert sind, ersetzen durch Std-Name, sonst bleibt Name unverändert 
                        Dim origPhName As String = msTask.Name
                        msTask.Name = phaseMappings.mapToStdName("", msTask.Name)

                        ' nachsehen, ob msTask.Name in PhaseDefinitions definiert ist
                        If Not PhaseDefinitions.Contains(msTask.Name) Then
                            Dim newPhaseDef As New clsPhasenDefinition
                            newPhaseDef.name = msTask.Name
                            ' Abbreviation, falls Customfield visbo_abbrev definiert ist
                            If visbo_abbrev <> 0 Then          ' VISBO-Abbrev ist definiert
                                newPhaseDef.shortName = msTask.GetField(visbo_abbrev)
                            Else
                                newPhaseDef.shortName = msTask.Name
                            End If
                            ' Task Class, falls Customfield visbo_taskclass definiert ist
                            If visbo_taskclass <> 0 Then          ' VISBO-TaskClass ist definiert
                                newPhaseDef.darstellungsKlasse = msTask.GetField(visbo_taskclass)
                            Else
                                newPhaseDef.darstellungsKlasse = ""
                            End If

                            newPhaseDef.UID = PhaseDefinitions.Count + 1
                            'PhaseDefinitions.Add(newPhaseDef)
                            missingPhaseDefinitions.Add(newPhaseDef)
                        End If

                        With cphase

                            If Not istElemID(msTask.Name) Then
                                .nameID = hproj.hierarchy.findUniqueElemKey(msTask.Name, False)
                            End If

                            If visboflag <> 0 Then          ' VISBO-Flag ist definiert

                                Dim hflag As Boolean = readCustomflag(msTask, visboflag)

                                ' Liste, ob Task in Projekt für die Projekt-Tafel aufgenommen werden soll, oder nicht
                                'visboFlagListe.Add(.nameID, msTask.GetField(visboflag) = pbYes)
                                visboFlagListe.Add(.nameID, hflag)

                            End If

                            ' Änderung 28.11.13: jetzt wird die Phasen Länge exakt bestimmt , über startoffset in Tagen und dauerinDays als Länge
                            Dim cphaseStartOffset As Long
                            Dim dauerIndays As Long
                            cphaseStartOffset = DateDiff(DateInterval.Day, hproj.startDate, CDate(msTask.Start))
                            dauerIndays = calcDauerIndays(CDate(msTask.Start), CDate(msTask.Finish))
                            .changeStartandDauer(cphaseStartOffset, dauerIndays)
                            .offset = 0

                            ' hier muss eine Routine aufgerufen werden, die die Dauer in Tagen berechnet !!!!!!
                            Dim phaseStartdate As Date = .getStartDate
                            Dim phaseEnddate As Date = .getEndDate


                            Dim anzRessources As Integer = msTask.Resources.Count

                            ' Resourcen je MSTask durchgehen
                            Dim j As Integer = 0
                            Dim ccost As clsKostenart = Nothing
                            Dim crole As clsRolle = Nothing



                            Dim ass As MSProject.Assignment

                            For Each ass In msTask.Assignments


                                Dim msRess As MSProject.Resource = ass.Resource

                                Select Case ass.Resource.Type
                                    Case MSProject.PjResourceTypes.pjResourceTypeMaterial To _
                                       MSProject.PjResourceTypes.pjResourceTypeCost
                                        Try

                                            Dim k As Integer = 0

                                            If CostDefinitions.containsName(ass.ResourceName) Then
                                                k = CInt(CostDefinitions.getCostdef(ass.ResourceName).UID)
                                            Else
                                                ' Kostenart existiert noch nicht
                                                ' wird hier neu aufgenommen
                                                Dim newCostDef As New clsKostenartDefinition
                                                newCostDef.name = ass.ResourceName
                                                newCostDef.farbe = RGB(120, 120, 120)   ' Farbe: grau
                                                newCostDef.UID = CostDefinitions.Count + 1
                                                If Not missingCostDefinitions.containsName(newCostDef.name) Then
                                                    missingCostDefinitions.Add(newCostDef)
                                                End If

                                                CostDefinitions.Add(newCostDef)

                                                ' Änderung tk: muss auf costdefinitions gesetzt werden 
                                                ' k = CInt(missingCostDefinitions.getCostdef(ass.ResourceName).UID)
                                                k = CInt(CostDefinitions.getCostdef(ass.ResourceName).UID)
                                            End If

                                            Dim work As Double = CType(ass.Work, Double)
                                            Dim cost As Double = CType(ass.Cost, Double)

                                            Dim startdate As Date = CDate(msTask.Start)
                                            Dim endedate As Date = CDate(msTask.Finish)

                                            Dim anzmonth As Integer = CInt(DateDiff(DateInterval.Month, startdate, endedate))
                                            Dim anzdays As Integer = CInt(DateDiff(DateInterval.Day, startdate, endedate))
                                            Dim anzhours As Integer = CInt(DateDiff(DateInterval.Hour, startdate, endedate))

                                            If anzhours > 0 And anzdays = 0 And anzmonth = 0 Then
                                                anzdays = 1
                                                anzmonth = 1
                                            End If
                                            If anzdays > 0 And anzmonth = 0 Then
                                                anzmonth = 1
                                            End If


                                            ReDim Xwerte(anzmonth - 1)

                                            Dim m As Integer
                                            For m = 1 To anzmonth

                                                Try
                                                    Xwerte(m - 1) = CType(cost / anzmonth, Double)
                                                Catch ex As Exception
                                                    Xwerte(m - 1) = 0.0
                                                End Try

                                            Next m

                                            ccost = New clsKostenart(anzmonth - 1)

                                            With ccost
                                                .KostenTyp = k
                                                .Xwerte = Xwerte
                                            End With


                                            With cphase
                                                .AddCost(ccost)
                                            End With
                                        Catch ex As Exception
                                            '
                                            ' handelt es sich um die Kostenart Definition?
                                            '
                                        End Try
                                        'Call MsgBox("Kosten = " & ass.ResourceName)

                                    Case MSProject.PjResourceTypes.pjResourceTypeWork

                                        Try
                                            Dim r As Integer = 0


                                            If RoleDefinitions.containsName(ass.ResourceName) Then
                                                r = CInt(RoleDefinitions.getRoledef(ass.ResourceName).UID)
                                            Else
                                                ' Rolle existiert noch nicht
                                                ' wird hier neu aufgenommen

                                                Dim newRoleDef As New clsRollenDefinition
                                                newRoleDef.name = ass.ResourceName
                                                newRoleDef.farbe = RGB(120, 120, 120)
                                                newRoleDef.defaultKapa = 200000

                                                ' OvertimeRate in Tagessatz umrechnen
                                                Dim hoverstr() As String = Split(CStr(ass.Resource.OvertimeRate), "/", -1)
                                                hoverstr = Split(hoverstr(0), "CHF", -1)
                                                newRoleDef.tagessatzExtern = CType(hoverstr(1), Double) * msproj.HoursPerDay

                                                ' StandardRate in Tagessatz umrechnen
                                                Dim hstdstr() As String = Split(CStr(ass.Resource.StandardRate), "/", -1)
                                                hstdstr = Split(hstdstr(0), "CHF", -1)
                                                newRoleDef.tagessatzIntern = CType(hstdstr(1), Double) * msproj.HoursPerDay

                                                newRoleDef.UID = RoleDefinitions.Count + 1
                                                If Not missingRoleDefinitions.containsName(newRoleDef.name) Then
                                                    missingRoleDefinitions.Add(newRoleDef)
                                                End If

                                                RoleDefinitions.Add(newRoleDef)

                                                ' Änderung tk: das muss von roledefinitions geholt werden ...
                                                ' r = CInt(missingRoleDefinitions.getRoledef(ass.ResourceName).UID)
                                                r = CInt(RoleDefinitions.getRoledef(ass.ResourceName).UID)
                                            End If



                                            Dim work As Double = CType(ass.Work, Double)
                                            'Dim duration As Double = CType(ass.Duration, Double)
                                            Dim unit As Double = CType(ass.Units, Double)


                                            Dim startdate As Date = CDate(msTask.Start)
                                            Dim endedate As Date = CDate(msTask.Finish)

                                            ' Änderung tk: wurde ersetzt durch tk Anpassung: keine Gleichverteilung auf die Monate, sondern 
                                            ' entsprechend der Lage der Monate ; es muss auch beachtet werden, dass anzmonth von 3.5 - 1.6 2 Monate sind; 
                                            ' die Berechnung Datediff ergibt aber nur 1 Monat '
                                            'Dim anzmonth As Integer = CInt(DateDiff(DateInterval.Month, startdate, endedate))
                                            'Dim anzdays As Integer = CInt(DateDiff(DateInterval.Day, startdate, endedate))
                                            'Dim anzhours As Integer = CInt(DateDiff(DateInterval.Hour, startdate, endedate))

                                            'If anzhours > 0 And anzdays = 0 And anzmonth = 0 Then
                                            '    anzdays = 1
                                            '    anzmonth = 1
                                            'End If
                                            'If anzdays > 0 And anzmonth = 0 Then
                                            '    anzmonth = 1
                                            'End If


                                            'ReDim Xwerte(anzmonth - 1)
                                            ' Ende Auskommentierung tk  

                                            ' tk Anpassung ...
                                            Dim oldWerte(0) As Double
                                            Dim anzmonth As Integer = getColumnOfDate(endedate) - getColumnOfDate(startdate) + 1
                                            oldWerte(0) = work
                                            ReDim Xwerte(anzmonth - 1)
                                            Call cphase.berechneBedarfe(startdate, endedate, oldWerte, 1.0, Xwerte)


                                            For m As Integer = 1 To anzmonth
                                                Xwerte(m - 1) = Xwerte(m - 1) / 60 / 8
                                            Next

                                            ' Ende tk Anpassung


                                            ' Änderung tk: wieder auskommentieren - alter Code: hier wurde gleichverteilt  
                                            'For m As Integer = 1 To anzmonth

                                            '    Try
                                            '        ' Xwerte in Anzahl Tage; in MSProject alle Werte in anz. Minuten
                                            '        Xwerte(m - 1) = CType(work / anzmonth / 60 / 8, Double)

                                            '    Catch ex As Exception
                                            '        Xwerte(m - 1) = 0.0
                                            '    End Try

                                            'Next m

                                            ' Check , um Unterschiede in der Summe herausfinden zu können
                                            ' die waren immer 0 ... 
                                            'Dim aChck As Double = Xwerte1.Sum - Xwerte.Sum

                                            crole = New clsRolle(anzmonth - 1)
                                            With crole
                                                .RollenTyp = r
                                                .Xwerte = Xwerte
                                            End With

                                            With cphase
                                                .addRole(crole)
                                            End With
                                        Catch ex As Exception

                                        End Try

                                        'Call MsgBox("Work = " & ass.ResourceName & " mit " & CStr(ass.Work) & "Arbeit")
                                End Select
                            Next ass


                            ' Hierarchie-Aufbau
                            Dim cphaseParent As Object = msTask.Parent

                            Dim hrchynode As New clsHierarchyNode
                            hrchynode.elemName = cphase.name

                            If tasklevel = 0 Then
                                hrchynode.parentNodeKey = ""

                            ElseIf tasklevel = 1 Then
                                hrchynode.parentNodeKey = rootPhaseName

                            ElseIf tasklevel - lastlevel = 1 Then
                                hrchynode.parentNodeKey = lastelemID

                            ElseIf tasklevel - lastlevel = 0 Then
                                hrchynode.parentNodeKey = hproj.hierarchy.getParentIDOfID(lastelemID)

                            ElseIf lastlevel - tasklevel >= 1 Then
                                Dim hilfselemID As String = lastelemID
                                For l As Integer = 1 To lastlevel - tasklevel
                                    hilfselemID = hproj.hierarchy.getParentIDOfID(hilfselemID)
                                Next l
                                hrchynode.parentNodeKey = hproj.hierarchy.getParentIDOfID(hilfselemID)
                            Else
                                Throw New ArgumentException("Fehler beim Import! Hierarchie kann nicht richtig aufgebaut werden")
                            End If

                            hproj.AddPhase(cphase, origName:=origPhName, parentID:=hrchynode.parentNodeKey)

                            ' '' ''hproj.hierarchy.addNode(hrchynode, cphase.nameID)
                            hrchynode.indexOfElem = hproj.AllPhases.Count
                            ' merken von letzem Element (Knoten,Phase,Meilenstein)
                            lasthrchyNode = hrchynode
                            lastelemID = cphase.nameID
                            lastphase = cphase
                            lastlevel = tasklevel
                        End With


                        Dim oBreadCrumb As String = hproj.hierarchy.getBreadCrumb(lastelemID)

                    Else
                        ' mstask ist ein Meilenstein und kein Summary-Meilenstein


                        ' Ergänzung tk für Demo BHTC 
                        ' falls Synonyme definiert sind, ersetzen durch Std-Name, sonst bleibt Name unverändert 
                        Dim origMsName As String = msTask.Name
                        msTask.Name = milestoneMappings.mapToStdName("", msTask.Name)
                        '


                        Dim hierarchy As String = msTask.WBS
                        'Dim oBreadCrumb As String = hproj.hierarchy.getBreadCrumb(lastelemID)
                        Dim msPhase As clsPhase = Nothing
                        Dim parentID As String = rootPhaseName

                        lastlevel = hproj.hierarchy.getIndentLevel(lastelemID)

                        If lastlevel = -1 Then          ' lastelemID existiert in der hierarchy nicht, also wird Meilenstein der Rootphase zugeordnet
                            parentID = rootPhaseName

                        ElseIf tasklevel = lastlevel Then
                            parentID = hproj.hierarchy.getParentIDOfID(lastelemID)

                        ElseIf tasklevel > lastlevel Then
                            parentID = lastelemID

                        ElseIf tasklevel = 1 And tasklevel < lastlevel Then
                            parentID = rootPhaseName

                        ElseIf lastlevel - tasklevel >= 1 Then
                            Dim hilfselemID As String = lastelemID
                            For l As Integer = 1 To lastlevel - tasklevel
                                hilfselemID = hproj.hierarchy.getParentIDOfID(hilfselemID)
                            Next l
                            parentID = hproj.hierarchy.getParentIDOfID(hilfselemID)

                        End If

                        msPhase = hproj.getPhaseByID(parentID)

                        Dim cmilestone As New clsMeilenstein(msPhase)


                        ' prüfen, ob MeilensteinDefinition bereits vorhanden
                        If Not MilestoneDefinitions.Contains(msTask.Name) Then
                            Dim msDef As New clsMeilensteinDefinition
                            msDef.belongsTo = msPhase.name
                            msDef.name = msTask.Name
                            ' Abbreviation, falls Customfield visbo_abbrev definiert ist
                            If visbo_abbrev <> 0 Then          ' VISBO-Abbrev ist definiert
                                msDef.shortName = msTask.GetField(visbo_abbrev)
                            Else
                                msDef.shortName = ""
                            End If
                            ' Task Class, falls Customfield visbo_taskclass definiert ist
                            If visbo_taskclass <> 0 Then          ' VISBO-TaskClass ist definiert
                                msDef.darstellungsKlasse = msTask.GetField(visbo_taskclass)
                            Else
                                msDef.darstellungsKlasse = ""
                            End If

                            msDef.schwellWert = 0
                            msDef.UID = MilestoneDefinitions.Count + 1
                            'MilestoneDefinitions.Add(msDef)
                            Try
                                missingMilestoneDefinitions.Add(msDef)
                            Catch ex As Exception
                            End Try


                        End If

                        ' MeilensteinDefinition vorhanden?
                        If MilestoneDefinitions.Contains(msTask.Name) _
                            Or missingMilestoneDefinitions.Contains(msTask.Name) Then

                            Dim msBewertung As New clsBewertung
                            cmilestone.setDate = CType(msTask.Start, Date)
                            cmilestone.nameID = hproj.hierarchy.findUniqueElemKey(msTask.Name, True)
                            msBewertung.description = msTask.Notes
                            If visbo_ampel <> 0 Then

                                Dim visboAmpel As String = msTask.GetField(visbo_ampel)

                                Select Case visboAmpel

                                    Case "none"
                                        msBewertung.colorIndex = PTfarbe.none
                                    Case "red"
                                        msBewertung.colorIndex = PTfarbe.red
                                    Case "green"
                                        msBewertung.colorIndex = PTfarbe.green
                                    Case "yellow"
                                        msBewertung.colorIndex = PTfarbe.yellow
                                    Case Else
                                        msBewertung.colorIndex = PTfarbe.none

                                End Select

                            Else
                                msBewertung.colorIndex = PTfarbe.none
                            End If

                            cmilestone.addBewertung(msBewertung)


                            If visboflag <> 0 Then        ' Ist VISBO-flag definiert?

                                Dim hflag As Boolean = readCustomflag(msTask, visboflag)
                                ' Liste, ob Meilenstein in Projekt für die Projekt-Tafel aufgenommen werden soll, oder nicht
                                visboFlagListe.Add(cmilestone.nameID, hflag)
                            End If

                            Try
                                With msPhase
                                    .addMilestone(cmilestone, origName:=origMsName)
                                End With
                            Catch ex1 As Exception
                                Throw New Exception(ex1.Message)
                            End Try
                        Else
                            Throw New ArgumentException("Fehler: Meilenstein konnte nicht gefunden werden")
                        End If
                    End If

                    '' Testweise hier eingetragen

                    Dim anzVorgaenger As Integer = msTask.PredecessorTasks.Count
                    Dim anzNachfolger As Integer = msTask.SuccessorTasks.Count
                    Dim dependencies As MSProject.TaskDependencies = msTask.TaskDependencies

                    Dim startTask As Date = CType(msTask.Start, Date)
                    Dim endeTask As Date = CType(msTask.Finish, Date)




                Next i          ' Ende Schleife über alle Tasks/Phasen eines Projektes

                Dim ele_i As Integer = 0
                Dim msStart As Integer = hproj.hierarchy.getIndexOf1stMilestone

                ' Liste der Phasen/Meilensteine durchgehen und die Phasen/Meilensteine die den visbo-Flag nicht gesetzt haben aus der Hierarchie löschen
                For ele_i = 0 To visboFlagListe.Count - 1

                    Dim elemID As String = visboFlagListe.ElementAt(ele_i).Key
                    If hproj.hierarchy.containsKey(elemID) Then

                        If Not visboFlagListe.ElementAt(ele_i).Value Then

                            If elemIDIstMeilenstein(elemID) Then

                                ' Meilenstein muss entfernt werden

                                Dim hrchynode As clsHierarchyNode = hproj.hierarchy.nodeItem(elemID)
                                If hrchynode.childCount > 0 Then
                                    Call MsgBox("Knoten " & elemNameOfElemID(elemID) & " kann nicht aus der Hierarchie entfernt werden")
                                Else
                                    hproj.removeMeilenstein(elemID)
                                End If

                            Else        ' Element elemID ist eine Phase


                                If isRemovable(elemID, hproj, visboFlagListe) Then

                                    ' es wird die Phase elemID mit allen seinen Kindern gelöscht
                                    hproj.removePhase(elemID, True)

                                    ' ''Call MsgBox("isRemovable = true" & vbLf & _
                                    ' ''            elemID & " kann entfernt werden")
                                Else

                                    '' ''Call MsgBox("isRemovable = false" & vbLf & _
                                    '' ''            elemID & " kann nicht entfernt werden ")

                                End If
                            End If

                        Else
                            ' Phase/Meilenstein bleibt erhalten
                        End If

                    Else
                        ' Element elemID wurde bereits entfernt '
                        ' Call MsgBox("das Element elemID= " & elemID & " wurde bereits entfernt")
                    End If

                Next  ' Schleife über alle Phasen/Meilensteine zum entfernern derer, die VISBO-Flag nicht gesetzt haben

                Dim key As String = calcProjektKey(hproj.name, hproj.variantName)

                ' prüfen, ob AlleProjekte das Projekt bereits enthält 
                ' danach ist sichergestellt, daß AlleProjekte das Projekt bereit enthält 
                If AlleProjekte.Containskey(key) Then
                    AlleProjekte.Remove(key)
                End If

                AlleProjekte.Add(hproj)

                If modus = "BHTC" Then
                    ' Alle Projekte entfernen
                    ShowProjekte.Clear()
                End If

                If Not ShowProjekte.contains(hproj.name) Then
                    ShowProjekte.Add(hproj)
                Else
                    ShowProjekte.Remove(hproj.name)
                    ShowProjekte.Add(hproj)
                    'Call MsgBox("Projekt " & hproj.name & " ist bereits in der Projekt-Liste enthalten")
                End If


                If modus <> "BHTC" Then

                    prj.FileExit(MSProject.PjSaveType.pjDoNotSave)
                    ' ''Else
                    ' ''    ' aktives Projekt durch hproj zurück an anrufende Routine übergeben

                    ' ''    If ShowProjekte.contains(active_proj) Then
                    ' ''        hproj = ShowProjekte.getProject(active_proj)
                    ' ''    End If

                End If

            Else

                Call MsgBox("Bitte zunächst ein Projekt öffnen !")

            End If
        Catch ex As Exception
            Call MsgBox(ex)
        End Try

        enableOnUpdate = True


    End Sub


    ''' <summary>
    ''' Methode trägt alle Projekte aus ImportProjekte in AlleProjekte bzw. Showprojekte ein, sofern die Anzahl mit der myCollection übereinstimmt
    ''' die Projekte werden in der Reihenfolge auf das Board gezeichnet, wie sie in der ImportProjekte aufgeführt sind
    ''' wenn ein importiertes Projekt bereits in der Datenbank existiert und verändert ist, dann wird es markiert und gleichzeitig temporär geschützt 
    ''' wenn ein importiertes Projekt bereits in der Datenbank existiert, verändert wurde und von anderen geschützt ist, dann wird eine Variante angelegt 
    ''' </summary>
    ''' <param name="importDate"></param>
    ''' <remarks></remarks>
    Public Sub importProjekteEintragen(ByVal importDate As Date, ByVal pStatus As String, _
                                       Optional drawPlanTafel As Boolean = True)


        'Public Sub importProjekteEintragen(ByVal myCollection As Collection, ByVal importDate As Date, ByVal pStatus As String, _
        '                                   Optional ByVal scenarioName As String = "")
        Dim hproj As New clsProjekt, cproj As New clsProjekt
        Dim fullName As String, vglName As String
        'Dim pname As String



        Dim anzAktualisierungen As Integer, anzNeuProjekte As Integer
        Dim tafelZeile As Integer = 2
        'Dim shpElement As Excel.Shape
        Dim phaseList As New Collection
        Dim milestoneList As New Collection
        Dim wasNotEmpty As Boolean

        Dim existsInSession As Boolean = False

        Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)

        ' aus der Datenbank alle WriteProtections holen ...
        If Not noDB And AlleProjekte.Count > 0 Then
            writeProtections.adjustListe = request.retrieveWriteProtectionsFromDB(AlleProjekte)
        End If

        If AlleProjekte.Count > 0 Then
            wasNotEmpty = True
            tafelZeile = projectboardShapes.getMaxZeile
        Else
            wasNotEmpty = False
        End If


        Dim differentToPrevious As Boolean = False

        ' Änderung tk 5.6.16: 
        'es wird jetzt getrennt zwischen dem was in einer Constellation gespeichert werden soll und dem , 
        ' was noch in die Session importiert werden muss. 

        ''If myCollection.Count <> ImportProjekte.Count Then
        ''    Throw New ArgumentException("keine Übereinstimmung in der Anzahl gültiger/ímportierter Projekte - Abbruch!")
        ''End If


        anzAktualisierungen = 0
        anzNeuProjekte = 0

        ' jetzt werden alle importierten Projekte bearbeitet 
        For Each kvp As KeyValuePair(Of String, clsProjekt) In ImportProjekte.liste

            ' wenn ein Projekt importiert wird, das durch andere geschützt ist , so wird eine neue Variante angelegt
            ' dann soll das ursprüngliche Projekt , sofern es in de rSession existiert, nicht aus der Session gelöscht werden 
            Dim newVariantGenerated As Boolean = False
            fullName = kvp.Key
            hproj = kvp.Value


            ' jetzt muss überprüft werden, ob dieses Projekt bereits in AlleProjekte / Showprojekte existiert 
            ' wenn ja, muss es um die entsprechenden Werte dieses Projektes (Status, etc)  ergänzt werden
            ' wenn nein, wird es im Show-Modus ergänzt 

            vglName = calcProjektKey(hproj)
            Try
                cproj = AlleProjekte.getProject(vglName)

                If IsNothing(cproj) Then
                    ' jetzt muss geprüft werden, ob das Projekt bereits in der Datenbank existiert ... 
                    existsInSession = False
                    If Not noDB Then
                        cproj = awinReadProjectFromDatabase(hproj.name, hproj.variantName, Date.Now)
                    End If
                Else
                    existsInSession = True
                End If

                ' ist es immer noch Nothing ? 
                If IsNothing(cproj) Then
                    ' wenn es jetzt immer noch Nothing ist, dann existiert es weder in der Datenbank noch in der Session .... 
                    If hproj.VorlagenName = "" Then
                        Try
                            Dim anzVorlagen = Projektvorlagen.Count
                            Dim vproj As clsProjektvorlage
                            hproj.VorlagenName = Projektvorlagen.Liste.Last.Value.VorlagenName

                            For i = 1 To anzVorlagen
                                vproj = Projektvorlagen.Liste.ElementAt(i - 1).Value
                                If vproj.farbe = hproj.farbe Then
                                    hproj.VorlagenName = vproj.VorlagenName
                                End If
                            Next

                        Catch ex1 As Exception

                        End Try
                    End If

                    Try
                        With hproj
                            ' 5.5.2014 ur: soll nicht wieder auf 0 gesetzt werden, sondern Einstellung beibehalten
                            '.earliestStart = 0
                            .earliestStartDate = .startDate
                            .latestStartDate = .startDate
                            .Id = vglName & "#" & importDate.ToString
                            ' 5.5.2014 ur: soll nicht wieder auf 0 gesetzt werden, sondern Einstellung beibehalten
                            '.latestStart = 0

                            ' Änderung tk 12.12.15: LeadPerson darf doch nicht auf leer gesetzt werden ...
                            '.leadPerson = " "
                            .shpUID = ""
                            .StartOffset = 0

                            ' ein importiertes Projekt soll normalerweise immer gleich  auf "beauftragt" gesetzt werden; 
                            ' das kann aber jetzt an der aufrufenden Stelle gesetzt werden 
                            ' Inventur: erst mal auf geplant, sonst beauftragt 
                            .Status = pStatus
                            .tfZeile = tafelZeile
                            .timeStamp = importDate

                        End With

                        ' Workaround: 
                        Dim tmpValue As Integer = hproj.dauerInDays
                        ' tk, Änderung 19.1.17 nicht mehr notwendig ..
                        'Call awinCreateBudgetWerte(hproj)
                        tafelZeile = tafelZeile + 1

                        anzNeuProjekte = anzNeuProjekte + 1
                    Catch ex1 As Exception
                        Throw New ArgumentException("Fehler bei Übernahme der Attribute des alten Projektes" & vbLf & ex1.Message)
                    End Try
                Else

                    ' jetzt sollen bestimmte Werte aus dem cproj übernommen werden 
                    ' das ist dann wichtig, wenn z.Bsp nur Rplan Excel Werte eingelesen werden, die enthalten ja nix ausser Termine ...
                    ' und in dem Fall können ja interaktiv bzw. über Export/Import Visbo Steckbrief Werte gesetzt worden sein 

                    Try
                        Call awinAdjustValuesByExistingProj(hproj, cproj, existsInSession, importDate, tafelZeile)
                    Catch ex As Exception
                        Call MsgBox(ex.Message)
                    End Try


                    If Not hproj.isIdenticalTo(vProj:=cproj) Then
                        ' das heisst, das Projekt hat sich verändert 
                        hproj.diffToPrev = True
                        If hproj.Status = ProjektStatus(1) Then
                            hproj.Status = ProjektStatus(2)
                        End If

                        ' wenn das Projekt bereits von anderen geschützt ist, soll es als Variante angelegt werden 
                        ' andernfalls soll es von mir geschützt werden ; allerdings soll es nur dann einen temporärewn Schutz bekommen, 
                        ' wenn es nicht schon von mir permanent geschützt ist 
                        If Not noDB Then
                            Dim wpItem As clsWriteProtectionItem

                            Dim isProtectedbyOthers As Boolean = Not tryToprotectProjectforMe(hproj.name, hproj.variantName)

                            If isProtectedbyOthers Then

                                ' nicht erfolgreich, weil durch anderen geschützt ... 
                                ' oder aber noch gar nicht in Datenbank: aber das ist noch nicht berücksichtigt  
                                wpItem = request.getWriteProtection(hproj.name, hproj.variantName)
                                writeProtections.upsert(wpItem)

                                ' jetzt Variante anlegen 
                                Dim teilName As String = dbUsername
                                If dbUsername.Length > 4 Then
                                    teilName = dbUsername.Substring(0, 4)
                                End If
                                Dim newVname As String = "I" & teilName
                                hproj.variantName = newVname

                                ' jetzt das Flag setzen 
                                newVariantGenerated = True
                            End If

                        End If


                    Else
                        hproj.diffToPrev = False
                    End If


                    anzAktualisierungen = anzAktualisierungen + 1

                    Try
                        If newVariantGenerated Then
                            ' das alte in AlleProjekte lassen 
                            ' das alte in ShowProjekte rausnehmen  
                            If ShowProjekte.contains(hproj.name) Then
                                ShowProjekte.Remove(hproj.name)
                            End If

                        ElseIf existsInSession Then
                            AlleProjekte.Remove(vglName, False)
                            If ShowProjekte.contains(hproj.name) Then
                                ShowProjekte.Remove(hproj.name, False)
                            End If
                        End If


                    Catch ex1 As Exception
                        Throw New ArgumentException("Fehler beim Update des Projektes " & ex1.Message)
                    End Try

                End If


            Catch ex As Exception



            End Try

                ' in beiden Fällen - sowohl bei neu wie auch Aktualisierung muss jetzt das Projekt 
                ' sowohl auf der Plantafel eingetragen werden als auch in ShowProjekte und in alleProjekte eingetragen 

                ' bringe das neue Projekt in Showprojekte und in AlleProjekte



            Try
                vglName = calcProjektKey(hproj.name, hproj.variantName)
                If existsInSession Then
                    AlleProjekte.Add(hproj, False)
                    ShowProjekte.Add(hproj, False)
                Else
                    AlleProjekte.Add(hproj)
                    ShowProjekte.Add(hproj)
                End If
                

                ' ggf Bedarfe anzeigen 
                If roentgenBlick.isOn Then
                    With roentgenBlick
                        Call awinShowNeedsofProject1(mycollection:=.myCollection, type:=.type, projektname:=hproj.name)
                    End With

                End If


            Catch ex As Exception
                'ur:16.1.2015: Dies ist kein Fehler sondern gewollt: 
                'Call MsgBox("Fehler bei Eintrag Showprojekte / Import " & hproj.name)
            End Try





        Next

        If ImportProjekte.Count < 1 Then
            If awinSettings.englishLanguage Then
                Call MsgBox(" no projects imported ...")
            Else
                Call MsgBox(" es wurden keine Projekte importiert ...")
            End If

        Else

            If awinSettings.englishLanguage Then
                
                Call MsgBox(ImportProjekte.Count & " projects were read " & vbLf & vbLf & _
                        anzNeuProjekte.ToString & " new projects" & vbLf & _
                        anzAktualisierungen.ToString & " project updates")
            Else
                
                Call MsgBox("es wurden " & ImportProjekte.Count & " Projekte bearbeitet!" & vbLf & vbLf & _
                        anzNeuProjekte.ToString & " neue Projekte" & vbLf & _
                        anzAktualisierungen.ToString & " Projekt-Aktualisierungen")
            End If
            
            

            ' Änderung tk: jetzt wird das neu gezeichnet 
            ' wenn anzNeuProjekte > 0, dann hat sich die Konstellataion verändert 
            If currentConstellationName <> calcLastSessionScenarioName() Then
                currentConstellationName = calcLastSessionScenarioName()
            End If


            If drawPlanTafel Then
                If wasNotEmpty Then
                    Call awinClearPlanTafel()
                End If

                'Call awinZeichnePlanTafel(True)
                Call awinZeichnePlanTafel(True)
                Call awinNeuZeichnenDiagramme(2)
            End If

            'Call storeSessionConstellation("Last")

        End If

        ImportProjekte.Clear(False)

    End Sub

    ''' <summary>
    ''' übernimmt vom existierenden Projekt einige Werte 
    ''' ist vor allem dann relevant wenn nur ein RPLAN Excel mit gerademal Terminen eingelesen wird ....
    ''' </summary>
    ''' <param name="hproj"></param>
    ''' <param name="cproj"></param>
    ''' <param name="existsInSession"></param>
    ''' <remarks></remarks>
    Private Sub awinAdjustValuesByExistingProj(ByRef hproj As clsProjekt, ByVal cproj As clsProjekt, _
                                               ByVal existsInSession As Boolean, ByVal importDate As Date, _
                                               ByRef tafelZeile As Integer)
        ' es existiert schon - deshalb müssen alle restlichen Werte aus dem cproj übernommen werden 
        Dim vglName As String = calcProjektKey(hproj)

        Try
            With hproj
                .farbe = cproj.farbe
                .Schrift = cproj.Schrift
                .Schriftfarbe = cproj.Schriftfarbe

                ' Änderung tk: das wird mit 28.12.16 nicht mehr benötigt ...  
                '.earliestStart = cproj.earliestStart
                '.earliestStartDate = cproj.earliestStartDate
                '.latestStart = cproj.latestStart
                '.latestStartDate = cproj.latestStartDate
                .earliestStartDate = .startDate
                .latestStartDate = .startDate

                .Id = vglName & "#" & importDate.ToString

                .StartOffset = 0

                ' Änderung 28.1.14: bei einem bereits existierenden Projekt muss der Status mitübernommen werden 
                ' tk 7.3.17 das soll jetzt nicht mehr gemacht werden 
                '.Status = cproj.Status ' wird evtl , falls sich Änderungen ergeben haben, noch geändert ...

                If existsInSession Then
                    .shpUID = cproj.shpUID
                    ' in diesem Fall heisst es ja genaus, dann ist es auch in der sortListe der Constellations bereits vorhanden ...
                    '.tfZeile = cproj.tfZeile
                Else
                    .shpUID = ""
                    .tfZeile = tafelZeile
                    tafelZeile = tafelZeile + 1
                End If


                .timeStamp = importDate
                .UID = cproj.UID

                ' tk 7.3.17 das soll jetzt nicht mehr gemacht werden  
                .VorlagenName = cproj.VorlagenName

                If .Erloes > 0 Then
                    ' Workaround: 
                    Dim tmpValue As Integer = hproj.dauerInDays
                    ' tk, Änderung 19.1.17 nicht mehr notwendig ..
                    ' Call awinCreateBudgetWerte(hproj)

                End If



            End With

        Catch ex As Exception
            Throw New ArgumentException("Fehler bei Übernahme der Attribute des alten Projektes" & vbLf & ex.Message)
        End Try

    End Sub


    ''' <summary>
    ''' wenn das Projekt mit Namen pName und Varianten-Name vName und einem TimeStamp kleiner/gleich datum in der Datenbank existiert, 
    ''' wird das Projekt als Ergebnis zurückgegeben
    ''' Nothing sonst 
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <param name="vName"></param>
    ''' <param name="datum"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function awinReadProjectFromDatabase(ByVal pName As String, ByVal vName As String, ByVal datum As Date) As clsProjekt
        Dim tmpResult As clsProjekt = Nothing

        '
        ' prüfen, ob es in der Datenbank existiert ... wenn ja,  laden und anzeigen
        Try

            If Not noDB Then
                Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
                If request.pingMongoDb() Then

                    If request.projectNameAlreadyExists(pName, vName, datum) Then

                        ' Projekt ist noch nicht im Hauptspeicher geladen, es muss aus der Datenbank geholt werden.
                        tmpResult = request.retrieveOneProjectfromDB(pName, vName, datum)

                    Else
                        ' nichts tun, tmpResult ist bereits Nothing 
                    End If
                Else
                    ' nichts tun, tmpResult ist bereits Nothing 
                End If
            End If


        Catch ex As Exception

        End Try

        awinReadProjectFromDatabase = tmpResult

    End Function

    ''' <summary>
    ''' liest den Wert eines Cusomized Flag. Das Ergebnis ist True oder False
    ''' </summary>
    ''' <param name="msTask"></param>
    ''' <param name="visboflag"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function readCustomflag(ByVal msTask As MSProject.Task, ByVal visboflag As MSProject.PjField) As Boolean

        Dim tskflag As Boolean = True
        Select Case visboflag
            Case MSProject.PjField.pjTaskFlag1
                tskflag = msTask.Flag1
            Case MSProject.PjField.pjTaskFlag2
                tskflag = msTask.Flag2
            Case MSProject.PjField.pjTaskFlag3
                tskflag = msTask.Flag3
            Case MSProject.PjField.pjTaskFlag4
                tskflag = msTask.Flag4
            Case MSProject.PjField.pjTaskFlag5
                tskflag = msTask.Flag5
            Case MSProject.PjField.pjTaskFlag6
                tskflag = msTask.Flag6
            Case MSProject.PjField.pjTaskFlag7
                tskflag = msTask.Flag7
            Case MSProject.PjField.pjTaskFlag8
                tskflag = msTask.Flag8
            Case MSProject.PjField.pjTaskFlag9
                tskflag = msTask.Flag9
            Case MSProject.PjField.pjTaskFlag10
                tskflag = msTask.Flag10
            Case MSProject.PjField.pjTaskFlag11
                tskflag = msTask.Flag11
            Case MSProject.PjField.pjTaskFlag12
                tskflag = msTask.Flag12
            Case MSProject.PjField.pjTaskFlag13
                tskflag = msTask.Flag13
            Case MSProject.PjField.pjTaskFlag14
                tskflag = msTask.Flag14
            Case MSProject.PjField.pjTaskFlag15
                tskflag = msTask.Flag15
            Case MSProject.PjField.pjTaskFlag16
                tskflag = msTask.Flag16
            Case MSProject.PjField.pjTaskFlag17
                tskflag = msTask.Flag17
            Case MSProject.PjField.pjTaskFlag18
                tskflag = msTask.Flag18
            Case MSProject.PjField.pjTaskFlag19
                tskflag = msTask.Flag19
            Case MSProject.PjField.pjTaskFlag20
                tskflag = msTask.Flag230

        End Select
        readCustomflag = tskflag
    End Function

    ''' <summary>
    ''' Prüft, ob eine Phase (elemID) aus dem Projekt hproj gelöscht werden kann, 
    ''' da weder sie selbst betrachtet werden soll, noch all ihre Kinder
    ''' </summary>
    ''' <param name="elemID"></param>
    ''' <param name="hproj"></param>
    ''' <param name="liste"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    '''
    Public Function isRemovable(ByVal elemID As String, ByVal hproj As clsProjekt, ByVal liste As SortedList(Of String, Boolean)) As Boolean

        Dim ind As Integer = 1
        Dim hrchynode As clsHierarchyNode = Nothing
        Dim result As Boolean

        result = True

        Try

            hrchynode = hproj.hierarchy.nodeItem(elemID)
            If hrchynode.childCount = 0 Then
                result = result And Not liste(elemID)
            End If
            If hrchynode.childCount > 0 And result Then

                While result And ind <= hrchynode.childCount

                    Dim nodeID As String = hrchynode.getChild(ind)
                    result = result And liste.ContainsKey(nodeID) And Not liste(nodeID)
                    result = result And isRemovable(nodeID, hproj, liste)
                    ind = ind + 1

                End While

            End If

        Catch ex As Exception
            Call MsgBox("Fehler bei der Prüfung, ob das Element elemID= " & elemID & " entfernt werden kann")
            Throw New ArgumentException("Fehler bei der Prüfung, ob das Element elemID entfernt werden kann")
        End Try

        isRemovable = result

    End Function

    ''' <summary>
    ''' liest alle in der Massen-Edit referenzierten Projekte ein und ersetzt die Werte dafür  
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub importiereMassenEdit()

        Dim projectName As String = ""
        Dim variantName As String = ""

        Dim phaseName As String = ""
        Dim phaseNameID As String = ""
        Dim rcName As String = ""

        Dim isRole As Boolean = False
        Dim isCost As Boolean = False

        Dim zeile As Integer = 2
        Dim spalte As Integer = 1
        Dim lastRow As Integer

        Dim startColumnData As Integer
        Dim endColumnData As Integer

        Dim tmpValues() As Double = Nothing

        Dim von As Integer, bis As Integer
        Dim vonDate As Date, bisDate As Date
        Dim ok As Boolean = False
        Dim hproj As clsProjekt = Nothing
        Dim vproj As clsProjekt = Nothing

        Try
            Dim activeWSListe As Excel.Worksheet = CType(appInstance.ActiveWorkbook.Worksheets("VISBO"), _
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet)

            With activeWSListe

                startColumnData = CType(.Range("StartData"), Excel.Range).Column
                endColumnData = CType(.Range("EndData"), Excel.Range).Column

                vonDate = CType(CType(.Range("StartData"), Excel.Range).Value, Date)
                bisDate = CType(CType(.Range("EndData"), Excel.Range).Value, Date)

                von = getColumnOfDate(vonDate)
                bis = getColumnOfDate(bisDate)

                ' jetzt die TimeZone markieren , ohne die sonstigen Konsequenzen .. 
                ' überlegen, ob hier nicht awinchangeTimeSpan aufgerufen werden sollte ...

                Call awinShowtimezone(von, bis, True)
                showRangeLeft = von
                showRangeRight = bis

                lastRow = CType(.Cells(2000, 1), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlUp).Row

                ' jetzt Zeile für Zeile auslesen 
                While zeile <= lastRow

                    Dim valuesDidChange As Boolean = False

                    Try

                    Catch ex As Exception
                        ' dann ist irgendwo was schief gegangen ... 

                    End Try

                    ' die Farben in der Zeile zurücksetzen , aber nicht in den Datenbereichen, weil sonst die Info zu den Phasen weg ist 
                    CType(.Range(.Cells(zeile, 1), .Cells(zeile, startColumnData - 1)), Excel.Range).Interior.ColorIndex = XlColorIndex.xlColorIndexNone
                    Dim namesOK As Boolean = True

                    Try
                        projectName = CStr(CType(.Cells(zeile, 2), Global.Microsoft.Office.Interop.Excel.Range).Value).Trim
                    Catch ex As Exception
                        projectName = ""
                        namesOK = False
                    End Try

                    Try
                        variantName = CStr(CType(.Cells(zeile, 3), Global.Microsoft.Office.Interop.Excel.Range).Value).Trim
                    Catch ex As Exception
                        variantName = ""
                    End Try

                    Try
                        phaseName = CStr(CType(.Cells(zeile, 4), Global.Microsoft.Office.Interop.Excel.Range).Value).Trim
                    Catch ex As Exception
                        phaseName = ""
                        namesOK = False
                    End Try


                    Try
                        Dim cellComment As Excel.Comment = CType(.Cells(zeile, 4), Global.Microsoft.Office.Interop.Excel.Range).Comment
                        If Not IsNothing(cellComment) Then
                            phaseNameID = cellComment.Text
                        Else
                            phaseNameID = calcHryElemKey(phaseName, False)
                        End If
                    Catch ex As Exception
                        phaseNameID = calcHryElemKey(phaseName, False)
                    End Try

                    Try
                        rcName = CStr(CType(.Cells(zeile, 5), Global.Microsoft.Office.Interop.Excel.Range).Value)
                    Catch ex As Exception
                        rcName = ""
                        namesOK = False
                    End Try

                    If namesOK Then

                        ok = False

                        Dim pKey As String = calcProjektKey(projectName, variantName)
                        If AlleProjekte.Containskey(pKey) Then
                            hproj = AlleProjekte.getProject(pKey)
                            ok = True
                        Else
                            ' in der Datenbank nachsehen und laden ... 
                            If Not noDB Then

                                '
                                ' prüfen, ob es in der Datenbank existiert ... wenn ja,  laden und anzeigen
                                Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
                                If request.pingMongoDb() Then

                                    If request.projectNameAlreadyExists(projectName, variantName, Date.Now) Then

                                        ' Projekt ist noch nicht im Hauptspeicher geladen, es muss aus der Datenbank geholt werden.
                                        hproj = request.retrieveOneProjectfromDB(projectName, variantName, Date.Now)
                                        ' jetzt in AlleProjekte eintragen ... 
                                        If Not IsNothing(hproj) Then
                                            AlleProjekte.Add(hproj)
                                            ok = True
                                        End If

                                    Else
                                        ' nicht in Session, nicht in Datenbank: nicht ok !
                                        ok = False
                                    End If
                                Else
                                    Throw New ArgumentException("Datenbank-Verbindung ist unterbrochen!" & vbLf & "Massen-Edit ..")
                                End If


                            Else
                                ' nicht in Session, keine Datenbank aktiv: nicht ok !
                                ok = False

                            End If


                        End If

                        If ok Then

                            If Not ImportProjekte.Containskey(pKey) Then
                                ImportProjekte.Add(hproj, False)
                            End If

                            ' hier kommt die eigentliche Behandlung , andernfalls Zeile rot einfärben ... 
                            ' hier ist das hproj gelesen 
                            ' jetzt prüfen, ob es die Phase gibt 
                            Dim cphase As clsPhase = hproj.getPhaseByID(phaseNameID)
                            If Not IsNothing(cphase) Then
                                ' es gibt die Phase

                                If RoleDefinitions.containsName(rcName) Then
                                    isRole = True
                                    isCost = False

                                ElseIf CostDefinitions.containsName(rcName) Then
                                    isCost = True
                                    isRole = False
                                Else
                                    isCost = False
                                    isRole = False
                                End If

                                ' jetzt werden die Werte ausgelesen ... 
                                ' die müssen an der Stelle ausgelesen werden, weil eine fehlende Rolle/kostenart nur angemeckert werden soll, 
                                ' wenn auch tmpValues.sum > 0 
                                ReDim tmpValues(bis - von)
                                Dim i As Integer

                                For i = 0 To bis - von

                                    Try
                                        tmpValues(i) = CDbl(CType(.Cells(zeile, startColumnData + 2 * i), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                        If tmpValues(i) < 0 Then
                                            tmpValues(i) = 0
                                            CType(.Cells(zeile, startColumnData + 2 * i), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color = awinSettings.AmpelRot
                                        End If
                                    Catch ex As Exception
                                        CType(.Cells(zeile, startColumnData + 2 * i), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color = awinSettings.AmpelRot
                                    End Try

                                Next


                                ' nur weitermachen, wenn es entweder eine gültige Rolle oder gültige Kostenart ist 
                                If isRole Or isCost Then


                                    If tmpValues.Sum > 0 Then

                                        Dim ixZeitraum As Integer, ix As Integer, anzLoops As Integer
                                        Call awinIntersectZeitraum(getColumnOfDate(cphase.getStartDate), getColumnOfDate(cphase.getEndDate), _
                                                                   ixZeitraum, ix, anzLoops)

                                        If anzLoops > 0 Then
                                            ' es gibt eine Überdeckung
                                            If isRole Then
                                                Dim tmpRole As clsRolle = cphase.getRole(rcName)
                                                ' wenn die Rolle in diesem Projekt noch nicht da war, dann wird eine neue Instanz angelegt 
                                                Dim didntExist As Boolean = False

                                                If IsNothing(tmpRole) Then
                                                    didntExist = True
                                                    Dim dimension As Integer = cphase.relEnde - cphase.relStart
                                                    tmpRole = New clsRolle(dimension)

                                                    With tmpRole
                                                        .RollenTyp = RoleDefinitions.getRoledef(rcName).UID
                                                    End With
                                                End If

                                                Dim xWerte() As Double = tmpRole.Xwerte

                                                ' jetzt werden die Werte überschrieben ...
                                                For al As Integer = 1 To anzLoops
                                                    If xWerte(ix + al - 1) <> tmpValues(ixZeitraum + al - 1) Then
                                                        valuesDidChange = True
                                                    End If
                                                    xWerte(ix + al - 1) = tmpValues(ixZeitraum + al - 1)
                                                Next

                                                If didntExist Then
                                                    cphase.addRole(tmpRole)
                                                End If

                                            ElseIf isCost Then
                                                Dim tmpCost As clsKostenart = cphase.getCost(rcName)
                                                ' wenn die Kostenart in diesem Projekt noch nicht da war, dann wird eine neue Instanz angelegt 
                                                Dim didntExist As Boolean = False

                                                If IsNothing(tmpCost) Then
                                                    didntExist = True
                                                    Dim dimension As Integer = cphase.relEnde - cphase.relStart
                                                    tmpCost = New clsKostenart(dimension)

                                                    With tmpCost
                                                        .KostenTyp = CostDefinitions.getCostdef(rcName).UID
                                                    End With
                                                End If

                                                Dim xWerte() As Double = tmpCost.Xwerte

                                                ' jetzt werden die Werte überschrieben ...
                                                For al As Integer = 1 To anzLoops
                                                    If xWerte(ix + al - 1) <> tmpValues(ixZeitraum + al - 1) Then
                                                        valuesDidChange = True
                                                    End If
                                                    xWerte(ix + al - 1) = tmpValues(ixZeitraum + al - 1)
                                                Next

                                                If didntExist Then
                                                    cphase.AddCost(tmpCost)
                                                End If

                                            End If

                                        End If
                                    Else
                                        ' Löschen der Rolle bzw. Kostenart aus dieser Phase
                                        valuesDidChange = True
                                        If isRole Then
                                            Call cphase.removeRoleByName(rcName)
                                        ElseIf isCost Then
                                            Call cphase.removeCostByName(rcName)
                                        End If
                                    End If


                                Else
                                    ' es gibt die Rolle / Kostenart nicht 
                                    If tmpValues.Sum > 0 Then
                                        CType(.Cells(zeile, 5), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color = awinSettings.AmpelRot
                                    Else
                                        ' keine Aktion notwendig 
                                    End If

                                End If

                            Else
                                ' es gibt die Phase nicht 
                                CType(.Cells(zeile, 4), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color = awinSettings.AmpelRot
                            End If
                        Else
                            ' Projekt- Variante existiert nicht !
                            CType(.Range(.Cells(zeile, 2), .Cells(zeile, 3)), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color = awinSettings.AmpelRot
                        End If

                    End If



                    If valuesDidChange Then
                        hproj.diffToPrev = True
                    End If

                    zeile = zeile + 1

                End While


            End With


        Catch ex As Exception
            Call MsgBox("Fehler beim Import der Massen-Edit Datei" & vbLf & ex.Message)
        End Try



    End Sub
    ''' <summary>
    ''' erzeugt eine Szenario Definition
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Public Function importScenarioDefinition(ByVal scenarioName As String) As clsConstellation

        Dim zeile As Integer, spalte As Integer
        

        Dim tfZeile As Integer = 2
        Dim listOfpNames As New SortedList(Of String, String)
        Dim pName As String = ""
        Dim variantName As String = ""

        Dim lastRow As Integer
        Dim lastColumn As Integer
        'Dim startSpalte As Integer
        
        Dim geleseneProjekte As Integer


        Dim firstZeile As Excel.Range

        Dim newC As New clsConstellation
        newC.constellationName = scenarioName
        newC.sortCriteria = ptSortCriteria.customTF


        zeile = 2
        spalte = 1
        geleseneProjekte = 0




        Try
            Dim activeWSListe As Excel.Worksheet
            Try
                activeWSListe = CType(appInstance.ActiveWorkbook.Worksheets("VISBO"), _
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet)
            Catch ex As Exception
                activeWSListe = CType(appInstance.ActiveWorkbook.Worksheets("Liste"), _
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet)
            End Try
            
            With activeWSListe

                firstZeile = CType(.Rows(1), Excel.Range)
                lastRow = CType(.Cells(2000, 1), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlUp).Row

                While zeile <= lastRow

                    ' Kommentare zurücksetzen ...
                    Try
                        CType(.Range(.Cells(zeile, 1), .Cells(zeile, lastColumn)), Global.Microsoft.Office.Interop.Excel.Range).ClearComments()
                    Catch ex As Exception

                    End Try

                    ' hier muss jetzt alles zurückgesetzt werden 
                    ' ansonsten könnten alte Werte übernommen werden aus der Projekt-Information von vorher ..
                    pName = CStr(CType(.Cells(zeile, spalte), Global.Microsoft.Office.Interop.Excel.Range).Value)

                    If IsNothing(pName) Then
                        CType(.Cells(zeile, spalte), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color = awinSettings.AmpelGelb
                        CType(.Cells(zeile, spalte), Global.Microsoft.Office.Interop.Excel.Range).AddComment(Text:="Projekt-Name fehlt ..")
                    ElseIf pName.Trim.Length < 2 Then

                        Try
                            CType(.Cells(zeile, spalte), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color = awinSettings.AmpelGelb
                            CType(.Cells(zeile, spalte), Global.Microsoft.Office.Interop.Excel.Range).AddComment(Text:="Projekt-Name muss mindestens 2 Buchstaben haben und eindeutig sein ..")
                        Catch ex As Exception

                        End Try


                    Else
                        variantName = CStr(CType(.Cells(zeile, spalte + 1), Global.Microsoft.Office.Interop.Excel.Range).Value)
                        If IsNothing(variantName) Then
                            variantName = ""
                        End If


                        Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, _
                                                   dbUsername, dbPasswort)

                        If request.projectNameAlreadyExists(pName, variantName, Date.Now) Then
                            ' als Constellation Item aufnehmen 
                            Dim cItem As New clsConstellationItem

                            With cItem
                                .projectName = pName
                                .variantName = variantName
                                .show = True
                                .zeile = zeile
                            End With

                            newC.add(cItem)

                        End If

                    End If

                    geleseneProjekte = geleseneProjekte + 1
                    zeile = zeile + 1

                End While


            End With
        Catch ex As Exception

            Throw New Exception("Fehler in Portfolio-Datei" & ex.Message)
        End Try



        Call MsgBox("gelesen: " & geleseneProjekte & vbLf & _
                    "Portfolio erzeugt: " & scenarioName)

        importScenarioDefinition = newC

    End Function

    ''' <summary>
    ''' erzeugt die Projekte, die in der Batch-Datei angegeben sind
    ''' stellt sie in ImportProjekte 
    ''' erstellt ein Szenario mit Namen der Batch-Datei; die Sortierung erfolgt über die Reihenfolge in der Batch-Datei 
    ''' das wird sichergestellt über Eintrag der tfzeile in hproj ... 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub awinImportProjektInventur()
        Dim zeile As Integer, spalte As Integer

        Dim tfZeile As Integer = 2
        Dim listOfpNames As New SortedList(Of String, String)
        Dim pName As String = ""
        Dim variantName As String = ""
        Dim vorlageName As String = ""
        Dim start As Date, inputStart As Date
        Dim startElem As String = ""
        Dim endElem As String = ""
        Dim ende As Date, inputEnde As Date
        Dim budget As Double
        Dim budgetInput As String = ""
        Dim dauer As Integer = 0
        Dim sfit As Double, risk As Double
        Dim capacityNeeded As String = ""
        Dim externCostInput As String = ""

        Dim description As String = ""
        Dim businessUnit As String = ""
        Dim createdProjects As Integer = 0
        Dim responsiblePerson As String = ""
        Dim custFields As New Collection
        ' wieviele Spalten müssen mindesten drin sein ... also was ist der standard 
        Dim nrOfStdColumns As Integer = 15

        Dim lastRow As Integer
        Dim lastColumn As Integer
        'Dim startSpalte As Integer
        Dim vglName As String = ""
        Dim hproj As clsProjekt
        Dim vproj As clsProjektvorlage
        Dim geleseneProjekte As Integer
        Dim ProjektdauerIndays As Integer = 0
        Dim ok As Boolean = False
        Dim refDauer As Double
        Dim vorgabeDauer As Double
        Dim abstandAnfang As Double
        Dim abstandEnde As Double
        Dim lastSpaltenValue As Integer

        Dim dauerFaktor As Double = 1.0
        Dim refProj As New clsProjekt

        Dim firstZeile As Excel.Range
        ' Änderung tk 5.6.16 wird jetzt an der Aufruf Schnittstelle gemacht 
        ''Dim scenarioName As String = appInstance.ActiveWorkbook.Name
        ''Dim tmpName As String = ""

        ' ''Dim namesForConstellation As New Collection
        ' '' bestimme den Namen des Szenarios - das ist gleich der Name der Excel Datei 
        ''Dim positionIX As Integer = scenarioName.IndexOf(".xls") - 1
        ''tmpName = ""
        ''For ih As Integer = 0 To positionIX
        ''    tmpName = tmpName & scenarioName.Chars(ih)
        ''Next
        ''scenarioName = tmpName.Trim

        ' Vorbedingung: das Excel File. das importiert werden soll , ist bereits geöffnet 

        zeile = 2
        spalte = 1
        geleseneProjekte = 0

        ' später, um mal das Einlesen einigermaßen intelligent zu machen .... 
        'Dim suchstr(1) As String
        'suchstr(ptInventurSpalten.Name) = "Name"
        'suchstr(ptInventurSpalten.Vorlage) = "Vorlage"
        'suchstr(ptInventurSpalten.Start) = "Start-Datum"
        'suchstr(ptInventurSpalten.Ende) = "Ende-Datum"
        'suchstr(ptInventurSpalten.startElement) = "Bezug Start"
        'suchstr(ptInventurSpalten.endElement) = "Bezug Ende"
        'suchstr(ptInventurSpalten.Dauer) = "Dauer [Tage]"
        'suchstr(ptInventurSpalten.Budget) = "Budget [T€]"
        'suchstr(ptInventurSpalten.Risiko) = "Risiko"
        'suchstr(ptInventurSpalten.Strategie) = "Strategie"
        'suchstr(ptInventurSpalten.Kapazitaet) = "benötigte Kapazität"
        'suchstr(ptInventurSpalten.Businessunit) = "Business Unit"
        'suchstr(ptInventurSpalten.Beschreibung) = "Beschreibung"


        'Dim inputColumns(11) As Integer



        Try
            Dim activeWSListe As Excel.Worksheet = CType(appInstance.ActiveWorkbook.Worksheets("Liste"), _
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet)
            With activeWSListe

                firstZeile = CType(.Rows(1), Excel.Range)

                ' für später ... siehe oben, intelligent ...
                '' jetzt werden die Spalten bestimmt 
                'Try
                '    For i As Integer = 0 To 13
                '        inputColumns(i) = firstZeile.Find(What:=suchstr(i)).Column
                '    Next
                'Catch ex As Exception

                'End Try

                'lastColumn = firstZeile.End(XlDirection.xlToLeft).Column
                lastColumn = firstZeile.Columns.Count
                lastColumn = CType(firstZeile, Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlToLeft).Column
                lastColumn = CType(.Cells(1, 2000), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlToLeft).Column
                lastRow = CType(.Cells(2000, 1), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlUp).Row

                While zeile <= lastRow
                    ok = False
                    Dim sMilestone As clsMeilenstein = Nothing
                    Dim eMilestone As clsMeilenstein = Nothing
                    ' Kommentare zurücksetzen ...
                    Try
                        CType(.Range(.Cells(zeile, 1), .Cells(zeile, lastColumn)), Global.Microsoft.Office.Interop.Excel.Range).ClearComments()
                    Catch ex As Exception

                    End Try

                    ' hier muss jetzt alles zurückgesetzt werden 
                    ' ansonsten könnten alte Werte übernommen werden aus der Projekt-Information von vorher ..
                    pName = CStr(CType(.Cells(zeile, spalte), Global.Microsoft.Office.Interop.Excel.Range).Value)
                    If IsNothing(pName) Then
                        CType(.Cells(zeile, spalte), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color = awinSettings.AmpelGelb
                        CType(.Cells(zeile, spalte), Global.Microsoft.Office.Interop.Excel.Range).AddComment(Text:="Projekt-Name fehlt ..")
                    ElseIf pName.Trim.Length < 2 Then

                        Try
                            CType(.Cells(zeile, spalte), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color = awinSettings.AmpelGelb
                            CType(.Cells(zeile, spalte), Global.Microsoft.Office.Interop.Excel.Range).AddComment(Text:="Projekt-Name muss mindestens 2 Buchstaben haben und eindeutig sein ..")
                        Catch ex As Exception

                        End Try

                    ElseIf Not isValidProjectName(pName) And Not pName.Contains("#") Then
                        Try
                            CType(.Cells(zeile, spalte), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color = awinSettings.AmpelGelb
                            CType(.Cells(zeile, spalte), Global.Microsoft.Office.Interop.Excel.Range).AddComment(Text:="Name darf keine ( oder ) Zeichen enthalten ..")
                        Catch ex As Exception

                        End Try
                    Else
                        variantName = ""
                        custFields.Clear()
                        capacityNeeded = ""

                        ' falls ein Varianten-Name mit angegeben wurde: pname#variantNAme 
                        Try
                            Dim tmpStr() As String = CStr(CType(.Cells(zeile, spalte), Global.Microsoft.Office.Interop.Excel.Range).Value).Split(New Char() {CChar("#")}, 2)
                            If tmpStr.Length > 1 Then
                                pName = makeValidProjectName(tmpStr(0))
                                variantName = tmpStr(1).Trim
                            End If
                        Catch ex As Exception
                            CType(.Cells(zeile, spalte), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color = awinSettings.AmpelGelb
                            variantName = ""
                        End Try

                        vorlageName = CStr(CType(.Cells(zeile, spalte + 1), Global.Microsoft.Office.Interop.Excel.Range).Value)
                        lastSpaltenValue = spalte + 1

                        If IsNothing(vorlageName) Then
                            CType(.Cells(zeile, lastSpaltenValue), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color = awinSettings.AmpelGelb
                        ElseIf vorlageName.Trim.Length = 0 Then
                            CType(.Cells(zeile, lastSpaltenValue), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color = awinSettings.AmpelGelb
                        Else
                            If Projektvorlagen.Liste.ContainsKey(vorlageName) Then

                                vproj = Projektvorlagen.getProject(vorlageName)
                                refProj = New clsProjekt
                                vproj.copyTo(refProj)
                                refProj.startDate = Date.Now

                                Try

                                    lastSpaltenValue = spalte + 2
                                    responsiblePerson = CStr(CType(.Cells(zeile, spalte + 2), Global.Microsoft.Office.Interop.Excel.Range).Value)

                                    lastSpaltenValue = spalte + 3
                                    start = CDate(CType(.Cells(zeile, spalte + 3), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                    ' eines der beiden Daten Start bzw Ende darf ohne Angabe bleiben ...
                                    'If start < StartofCalendar Then
                                    '    Throw New ArgumentException("Datum vor Kalender-Start")
                                    'End If

                                    lastSpaltenValue = spalte + 4
                                    ende = CDate(CType(.Cells(zeile, spalte + 4), Global.Microsoft.Office.Interop.Excel.Range).Value)


                                    If start < StartofCalendar And ende < StartofCalendar Then
                                        Throw New ArgumentException("sowohl Start wie Ende-Datum liegen vor dem Kalender-Start")
                                    End If

                                    lastSpaltenValue = spalte + 5
                                    startElem = CStr(CType(.Cells(zeile, spalte + 5), Global.Microsoft.Office.Interop.Excel.Range).Value)

                                    lastSpaltenValue = spalte + 6
                                    endElem = CStr(CType(.Cells(zeile, spalte + 6), Global.Microsoft.Office.Interop.Excel.Range).Value)

                                    lastSpaltenValue = spalte + 7
                                    dauer = CInt(CType(.Cells(zeile, spalte + 7), Global.Microsoft.Office.Interop.Excel.Range).Value)

                                    ' Konsistenzprüfung 
                                    If start > StartofCalendar And ende > StartofCalendar And dauer > 0 Then
                                        Throw New ArgumentException("Überbestimmt: es kann nicht Start, Ende und Dauer angegeben werden .. ")
                                    End If

                                    lastSpaltenValue = spalte + 8
                                    budgetInput = CStr(CType(.Cells(zeile, spalte + 8), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                    If budgetInput <> "calcNeeded" And IsNumeric(budgetInput) Then
                                        budget = CDbl(CType(.Cells(zeile, spalte + 8), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                        If budget < 0 Then
                                            Throw New ArgumentException("negative Werte nicht zugelassen!")
                                        End If
                                    ElseIf budgetInput = "calcNeeded" Then
                                        ' das bedeutet, dass das Budget errechnet werden soll ... 
                                        budget = -999
                                    ElseIf budgetInput = "" Then
                                        budget = 0
                                    Else
                                        Throw New ArgumentException("mit dieser Angabe konnte nichts angefangen werden ...")
                                    End If


                                    lastSpaltenValue = spalte + 9
                                    capacityNeeded = CStr(CType(.Cells(zeile, spalte + 9), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                    If Not isValidRoleCostInput(capacityNeeded, True) Then
                                        Throw New ArgumentException("ungültige Kapa-Angabe")
                                    End If

                                    lastSpaltenValue = spalte + 10
                                    externCostInput = CStr(CType(.Cells(zeile, spalte + 10), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                    If Not isValidRoleCostInput(externCostInput, False) Then
                                        Throw New ArgumentException("ungültige Kosten-Angabe")
                                    End If

                                    ' Konsistenzprüfung ...
                                    ' es darf nicht sein, dass Budget und externe Kosten berechnet werden sollen ...
                                    If budget = -999 And externCostInput = "filltobudget" Then
                                        Throw New ArgumentException("unterbestimmt: es können nicht sowohl Budget als auch externe Kosten berechnet werden")
                                    End If

                                    lastSpaltenValue = spalte + 11
                                    risk = CDbl(CType(.Cells(zeile, spalte + 11), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                    If risk < 0 Or risk > 10.0 Then
                                        Throw New ArgumentException("Kennzahl muss zwischen [0 und 10] liegen")
                                    End If

                                    lastSpaltenValue = spalte + 12
                                    sfit = CDbl(CType(.Cells(zeile, spalte + 12), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                    If sfit < 0 Or risk > 10.0 Then
                                        Throw New ArgumentException("Kennzahl muss zwischen [0 und 10] liegen")
                                    End If


                                    lastSpaltenValue = spalte + 13
                                    businessUnit = CStr(CType(.Cells(zeile, spalte + 13), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                    If Not IsNothing(businessUnit) Then
                                        Dim bi As Integer = 0
                                        Dim found As Boolean = False
                                        While bi <= businessUnitDefinitions.Count - 1 And Not found
                                            If businessUnitDefinitions.ElementAt(bi).Value.name = businessUnit Then
                                                found = True
                                            Else
                                                bi = bi + 1
                                            End If
                                        End While

                                        If Not found Then
                                            Throw New ArgumentException("Business Unit unbekannt ..")
                                        End If
                                    End If


                                    lastSpaltenValue = spalte + 14
                                    description = CStr(CType(.Cells(zeile, spalte + 14), Global.Microsoft.Office.Interop.Excel.Range).Value)

                                    If lastColumn > nrOfStdColumns Then
                                        ' es gibt evtl Custom fields 
                                        For i As Integer = nrOfStdColumns To lastColumn - 1

                                            Try
                                                Dim cfName As String = CStr(CType(.Cells(1, spalte + i), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                                Dim uniqueID As Integer = customFieldDefinitions.getUid(cfName)

                                                If uniqueID > 0 Then
                                                    ' es ist eine Custom Field

                                                    Dim cfType As Integer = customFieldDefinitions.getTyp(uniqueID)
                                                    Dim cfValue As Object = Nothing
                                                    Dim tstStr As String

                                                    Select Case cfType
                                                        Case ptCustomFields.Str
                                                            lastSpaltenValue = spalte + i
                                                            cfValue = CStr(CType(.Cells(zeile, spalte + i), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                                        Case ptCustomFields.Dbl
                                                            lastSpaltenValue = spalte + i
                                                            cfValue = CDbl(CType(.Cells(zeile, spalte + i), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                                        Case ptCustomFields.bool
                                                            lastSpaltenValue = spalte + i
                                                            cfValue = CBool(CType(.Cells(zeile, spalte + i), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                                    End Select

                                                    Dim cfObj As New clsCustomField
                                                    With cfObj
                                                        .uid = uniqueID
                                                        .wert = cfValue
                                                        tstStr = CStr(.wert)
                                                    End With
                                                    custFields.Add(cfObj)
                                                End If
                                            Catch ex As Exception
                                                CType(.Cells(zeile, lastSpaltenValue), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color = awinSettings.AmpelGelb
                                            End Try

                                        Next
                                    End If

                                    vglName = calcProjektKey(pName.Trim, variantName)
                                    inputStart = start
                                    inputEnde = ende

                                    If DateDiff(DateInterval.Day, StartofCalendar, start) >= 0 Then

                                        If DateDiff(DateInterval.Day, start, ende) > 0 Then
                                            ' nichts tun , Ende-Datum ist ein gültiges Datum
                                            ok = True
                                        ElseIf DateDiff(DateInterval.Day, StartofCalendar, ende) >= 0 Then
                                            ' auch Ende ist ein gültiges Datum , liegt nur vor Start
                                            ' also vertauschen der beiden 
                                            Dim tmpDate As Date = ende
                                            ende = start
                                            start = tmpDate
                                            ok = True
                                        Else
                                            ' Ende Datum wird anhand der Laufzeit der Vorlage oder der Dauer berechnet
                                            If dauer > 0 Then
                                                ProjektdauerIndays = dauer
                                            Else
                                                ProjektdauerIndays = vproj.dauerInDays
                                            End If
                                            ende = calcDatum(start, ProjektdauerIndays)
                                            ok = True
                                        End If

                                    ElseIf DateDiff(DateInterval.Day, StartofCalendar, ende) >= 0 Then
                                        ' hier ist Start kein gültiges Datum innerhalb der Projekt-Tafel 
                                        ' Start Datum wird anhand der Laufzeit der Vorlage berechnet
                                        If dauer > 0 Then
                                            ProjektdauerIndays = -1 * dauer
                                        Else
                                            ProjektdauerIndays = -1 * vproj.dauerInDays
                                        End If

                                        start = calcDatum(ende, ProjektdauerIndays)

                                        If DateDiff(DateInterval.Day, StartofCalendar, start) >= 0 Then
                                            ' Start ist ein korrektes Datum 
                                            ok = True
                                        Else
                                            CType(.Cells(zeile, spalte + 1), Global.Microsoft.Office.Interop.Excel.Range).Value = "Start liegt vor Kalender-Start "
                                            ok = False
                                        End If

                                    Else
                                        CType(.Cells(zeile, spalte + 1), Global.Microsoft.Office.Interop.Excel.Range).Value = "ungültiges Start- und Ende-Datum"
                                        ok = False
                                    End If

                                Catch ex As Exception

                                    ok = False
                                    'Call MsgBox(ex.Message)
                                    CType(.Cells(zeile, lastSpaltenValue), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color = awinSettings.AmpelGelb
                                    CType(.Cells(zeile, lastSpaltenValue), Global.Microsoft.Office.Interop.Excel.Range).AddComment(Text:=ex.Message)
                                End Try

                                ' jetzt die Daten richtig berechnen, falls Bezug Start , Bezug Ende angegeben ist 

                                vorgabeDauer = calcDauerIndays(start, ende)
                                Try

                                    If Not IsNothing(startElem) Then
                                        If startElem.Trim.Length > 0 Then
                                            sMilestone = refProj.getMilestone(startElem)
                                        End If
                                    End If

                                    If Not IsNothing(endElem) Then
                                        If endElem.Trim.Length > 0 Then
                                            eMilestone = refProj.getMilestone(endElem)
                                        End If
                                    End If

                                    ' jetzt werden Start und Ende ggf neu bestimmt, so dass die Bezugs-Elemente genau so liegen 
                                    If Not IsNothing(sMilestone) Then
                                        abstandAnfang = DateDiff(DateInterval.Day, refProj.startDate, sMilestone.getDate) * -1
                                        If Not IsNothing(eMilestone) Then
                                            abstandEnde = DateDiff(DateInterval.Day, eMilestone.getDate, refProj.endeDate)
                                            refDauer = calcDauerIndays(sMilestone.getDate, eMilestone.getDate)
                                        Else
                                            refDauer = calcDauerIndays(sMilestone.getDate, refProj.endeDate)
                                        End If
                                    Else
                                        If Not IsNothing(eMilestone) Then
                                            abstandEnde = DateDiff(DateInterval.Day, eMilestone.getDate, refProj.endeDate)
                                            refDauer = calcDauerIndays(refProj.startDate, eMilestone.getDate)
                                        Else
                                            refDauer = vorgabeDauer
                                        End If
                                    End If

                                    If refDauer < 0 Then
                                        refDauer = -1 * refDauer
                                    ElseIf refDauer = 0 Then
                                        refDauer = vorgabeDauer
                                    End If

                                    dauerFaktor = vorgabeDauer / refDauer

                                    ' rechne den neuen Start aus 
                                    If Not IsNothing(sMilestone) Then
                                        start = start.AddDays(CInt(dauerFaktor * abstandAnfang))
                                        ende = start.AddDays(CInt(dauerFaktor * vproj.dauerInDays - 1))
                                    ElseIf Not IsNothing(eMilestone) Then
                                        ende = start.AddDays(CInt(dauerFaktor * vproj.dauerInDays - 1))
                                    End If

                                Catch ex As Exception
                                    ' nichts tn 
                                End Try


                            Else
                                'CType(.Cells(zeile, spalte + 1), Global.Microsoft.Office.Interop.Excel.Range).Value = ".?."
                                CType(.Cells(zeile, lastSpaltenValue), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color = awinSettings.AmpelGelb
                                ok = False
                            End If

                            ' jetzt die Aktion durchführen, wenn alles ok 
                            If ok Then

                                'Projekt anlegen ,Verschiebung um 
                                hproj = New clsProjekt(start, start.AddMonths(-1), start.AddMonths(1))

                                ' #####################################################################
                                ' Erstellen des Projekts nach den Angaben aus der Batch-Datei 
                                '
                                hproj = erstelleInventurProjekt(pName, vorlageName, variantName, _
                                                             start, ende, budget, zeile, sfit, risk, _
                                                             capacityNeeded, externCostInput, businessUnit, description, custFields, _
                                                             responsiblePerson)


                                If Not IsNothing(hproj) Then

                                    ' immer als Fixiertes Projekt darstellen ..
                                    hproj.Status = ProjektStatus(1)

                                    'prüfen ob Rundungsfehler bei Setzen Meilenstein passiert sind ... 
                                    If Not IsNothing(sMilestone) Then
                                        If DateDiff(DateInterval.Day, hproj.getMilestone(startElem).getDate, inputStart) <> 0 Then
                                            'Call MsgBox("Differenz Start:" & DateDiff(DateInterval.Day, hproj.getMilestone(startElem).getDate, inputStart))
                                            hproj.getMilestone(startElem).setDate = inputStart
                                        End If
                                    End If

                                    If Not IsNothing(eMilestone) Then
                                        If DateDiff(DateInterval.Day, hproj.getMilestone(endElem).getDate, inputEnde) <> 0 Then
                                            'Call MsgBox("Differenz Ende:" & DateDiff(DateInterval.Day, hproj.getMilestone(endElem).getDate, inputEnde))
                                            hproj.getMilestone(endElem).setDate = inputEnde
                                        End If
                                    End If

                                Else
                                    ok = False
                                    CType(.Range(.Cells(zeile, 1), .Cells(zeile, 15)), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color = awinSettings.AmpelGelb
                                    CType(.Cells(zeile, lastSpaltenValue), Global.Microsoft.Office.Interop.Excel.Range).AddComment(Text:="Projekt konnte nicht erzeugt werden ...")
                                End If


                                If ok Then ' wenn es nicht explizit auf false gesetzt wurde, ist es an dieser Stelle immer noch true 
                                    Dim pkey As String = ""
                                    If Not IsNothing(hproj) Then
                                        Try
                                            pkey = calcProjektKey(hproj)

                                            If ImportProjekte.Containskey(pkey) Then
                                                CType(.Cells(zeile, 1), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color = awinSettings.AmpelGelb
                                                CType(.Cells(zeile, 1), Global.Microsoft.Office.Interop.Excel.Range).AddComment(Text:="Name existiert bereits")
                                            Else

                                                createdProjects = createdProjects + 1
                                                ' jetzt in die Liste der 
                                                If Not listOfpNames.ContainsValue(hproj.name) Then
                                                    hproj.tfZeile = tfZeile
                                                    Dim tmpKey As String = calcSortKeyCustomTF(tfZeile)
                                                    listOfpNames.Add(tmpKey, hproj.name)
                                                    tfZeile = tfZeile + 1
                                                Else
                                                    hproj.tfZeile = CInt(listOfpNames.ElementAt(listOfpNames.IndexOfValue(hproj.name)).Key)
                                                End If

                                                ImportProjekte.Add(hproj, False)
                                            End If


                                            'myCollection.Add(calcProjektKey(hproj))
                                        Catch ex As Exception
                                            CType(.Cells(zeile, 1), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color = awinSettings.AmpelGelb
                                            CType(.Cells(zeile, 1), Global.Microsoft.Office.Interop.Excel.Range).AddComment(Text:=ex.Message)
                                        End Try

                                    End If

                                End If

                            End If
                        End If


                    End If


                    geleseneProjekte = geleseneProjekte + 1
                    zeile = zeile + 1

                End While


            End With
        Catch ex As Exception

            Throw New Exception("Fehler in Portfolio-Datei" & ex.Message)
        End Try


        Call MsgBox("gelesen: " & geleseneProjekte & vbLf & _
                    "erzeugt: " & createdProjects & vbLf & _
                    "importiert: " & ImportProjekte.Count)

    End Sub

    ''' <summary>
    ''' bestimmt, ob es sich um einen gültigen Kapazitäts- bzw Kosten-Input String handelt
    ''' alle Rollen- bzw Kostenart Namen bekannt, alle Werte >= 0 
    ''' </summary>
    ''' <param name="inputStr"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function isValidRoleCostInput(ByVal inputStr As String, ByVal checkRoles As Boolean) As Boolean
        Dim resultValue As Boolean = True
        Dim anzDefinitions As Integer

        If checkRoles Then
            anzDefinitions = RoleDefinitions.Count
        Else
            anzDefinitions = CostDefinitions.Count
        End If


        If Not IsNothing(inputStr) Then
            If inputStr.Trim.Length > 0 Then

                Dim completeStr() As String = inputStr.Split(New Char() {CType("#", Char)}, 100)


                ' jetzt die ganzen Rollen bzw. Kosten abarbeiten 
                Dim i As Integer = 1
                While i <= completeStr.Length And resultValue = True

                    Dim roleCostStr() As String = completeStr(i - 1).Split(New Char() {CType(":", Char)}, 2)

                    If roleCostStr.Length = 2 Then

                        Try
                            Dim roleCostName As String = roleCostStr(0).Trim
                            Dim roleCostSum As Double = CDbl(roleCostStr(1).Trim)
                            If checkRoles Then
                                If RoleDefinitions.containsName(roleCostName) And roleCostSum >= 0 Then
                                    ' ok, nichts tun 
                                Else
                                    resultValue = False
                                End If

                            Else
                                If CostDefinitions.containsName(roleCostName) And roleCostSum >= 0 Then
                                    ' ok, nichts tun 
                                Else
                                    resultValue = False
                                End If

                            End If

                        Catch ex As Exception
                            resultValue = False
                        End Try

                    ElseIf roleCostStr.Length = 1 And anzDefinitions >= 1 Then
                        ' es muss sich um eine Zahl größer 0 handeln, Rolle 1 wird angenommen 

                        Try
                            If IsNumeric(roleCostStr(0).Trim) Then
                                If CDbl(roleCostStr(0).Trim) >= 0 Then
                                    ' ok, nichts tun
                                Else
                                    resultValue = False
                                End If

                            ElseIf Not checkRoles And roleCostStr(0) = "filltobudget" Then
                                ' ok , nichts tun 

                            Else
                                resultValue = False
                            End If
                        Catch ex As Exception
                            resultValue = False
                        End Try

                    Else
                        resultValue = False
                    End If

                    i = i + 1

                End While



            Else
                ' leerer String, ok 
                resultValue = True
            End If
        Else
            ' Nothing, ok 
            resultValue = True
        End If

        isValidRoleCostInput = resultValue

    End Function

    ''' <summary>
    ''' prüft, ob es sich um einen zugelassenen Projekt-Namen handelt ....
    ''' nicht zugelassen: #, (, ), Zeilenvorschub 
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function isValidProjectName(ByVal pName As String) As Boolean
        Dim ergebnis As Boolean = False
        If pName.Contains("#") Or _
            pName.Contains("(") Or _
            pName.Contains(")") Or _
            pName.Contains(vbCr) Or _
            pName.Contains(vbLf) Then
            ergebnis = False
        Else
            ergebnis = True
        End If

        isValidProjectName = ergebnis

    End Function

    ''' <summary>
    ''' macht aus einem evtl ungültigen Namen einen gültigen Projekt-NAmen 
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function makeValidProjectName(ByVal pName As String) As String

        If pName.Contains("#") Then
            pName = pName.Replace("#", "-")
        End If
        If pName.Contains("(") Then
            pName = pName.Replace("(", "/")
        End If
        If pName.Contains(")") Then
            pName = pName.Replace(")", "/")
        End If

        makeValidProjectName = pName

    End Function

    ''' <summary>
    ''' diese Funktion verarbeitet die Import Projekte 
    ''' wenn sie schon in der Datenbank bzw Session existieren und unterschiedlich sind: es wird eine Variante angelegt, die so heisst wie das Scenario 
    ''' wenn sie bereits existieren und identisch sind: in AlleProjekte holen, wenn nicht schon geschehen
    ''' wenn sie noch nicht existieren: in AlleProjekte anlegen
    ''' in jedem Fall: eine Constellation mit dem Namen cName anlegen
    ''' </summary>
    ''' <param name="cName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function verarbeiteImportProjekte(ByVal cName As String, _
                                             Optional ByVal noComparison As Boolean = False) As clsConstellation
        Dim newC As New clsConstellation
        newC.constellationName = cName
        ' das Folgende soll sihcerstellen, dass die Projekte in der Reihenfolge ihres Auftretens in der Excel Datei eingelesen und dargestellt werden ..
        newC.sortCriteria = ptSortCriteria.customListe
        currentSessionConstellation.sortCriteria = ptSortCriteria.customTF

        Dim vglProj As clsProjekt
        Dim lfdZeilenNr As Integer = 2
        Dim ok As Boolean

        Dim importDate As Date = Date.Now

        For Each kvp As KeyValuePair(Of String, clsProjekt) In ImportProjekte.liste

            Dim impProjekt As clsProjekt = kvp.Value

            ' jetzt das Import Datum setzen ...
            impProjekt.timeStamp = importDate

            Dim importKey As String = calcProjektKey(impProjekt)

            vglProj = Nothing

            If noComparison Then
                ' nicht vergleichen, einfach in AlleProjekte rein machen 
                If AlleProjekte.Containskey(importKey) Then
                    AlleProjekte.Remove(importKey)
                End If
                AlleProjekte.Add(impProjekt)
            Else
                ' jetzt muss ggf verglichen werden 
                If AlleProjekte.Containskey(importKey) Then

                    vglProj = AlleProjekte.getProject(importKey)

                Else
                    ' nicht in der Session, aber ist es in der Datenbank ?  

                    If Not noDB Then

                        '
                        ' prüfen, ob es in der Datenbank existiert ... wenn ja,  laden und anzeigen
                        Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
                        If request.pingMongoDb() Then

                            If request.projectNameAlreadyExists(impProjekt.name, impProjekt.variantName, Date.Now) Then

                                ' Projekt ist noch nicht im Hauptspeicher geladen, es muss aus der Datenbank geholt werden.
                                vglProj = request.retrieveOneProjectfromDB(impProjekt.name, impProjekt.variantName, Date.Now)

                                If IsNothing(vglProj) Then
                                    ' kann eigentlich nicht sein 
                                    ok = False
                                Else
                                    ' jetzt in AlleProjekte eintragen ... 
                                    AlleProjekte.Add(vglProj)

                                End If
                            Else
                                ' nicht in der Session, nicht in der Datenbank : also in AlleProjekte eintragen ... 
                                AlleProjekte.Add(impProjekt)
                            End If
                        Else
                            Throw New ArgumentException("Datenbank-Verbindung ist unterbrochen!" & vbLf & "Projekt '" & impProjekt.name & "'konnte nicht geladen werden")
                        End If


                    Else
                        ' nicht in der Session, nicht in der Datenbank : also in AlleProjekte eintragen ... 
                        AlleProjekte.Add(impProjekt)

                    End If


                End If

            End If



            ' wenn jetzt vglProj <> Nothing, dann vergleichen und ggf Variante anlegen ...
            If Not IsNothing(vglProj) And Not noComparison Then

                ' erstezt durch Abfrage auf Identität 
                'Dim unterschiede As Collection = impProjekt.listOfDifferences(vglProj, True, 0)

                If Not impProjekt.isIdenticalTo(vglProj) Then
                    ' es gibt Unterschiede, also muss eine Variante angelegt werden 

                    impProjekt.variantName = cName
                    importKey = calcProjektKey(impProjekt)

                    ' wenn die Variante bereits in der Session existiert ..
                    ' wird die bisherige gelöscht , die neue über ImportProjekte neu aufgenommen  
                    If AlleProjekte.Containskey(importKey) Then
                        AlleProjekte.Remove(importKey)
                    End If

                    ' jetzt das Importierte PRojekt in AlleProjekte aufnehmen 
                    AlleProjekte.Add(impProjekt)
                End If

            End If

            ' Aufnehmen in Constellation
            Dim newCItem As New clsConstellationItem
            newCItem.projectName = impProjekt.name
            newCItem.variantName = impProjekt.variantName
            If newC.containsProject(impProjekt.name) Then
                newCItem.show = False
            Else
                newCItem.show = True
            End If
            newCItem.start = impProjekt.startDate
            newCItem.zeile = lfdZeilenNr
            newC.add(newCItem)

            lfdZeilenNr = lfdZeilenNr + 1

        Next

        verarbeiteImportProjekte = newC

    End Function

    Public Sub awinImportModule(ByRef myCollection As Collection)

        Dim zeile As Integer, spalte As Integer
        Dim pName As String = ""
        Dim vorlagenName As String = ""
        Dim start As Date
        Dim ende As Date
        Dim budget As Double
        Dim dauer As Integer = 0
        Dim sfit As Double, risk As Double
        Dim volume As Double, complexity As Double
        Dim description As String = ""
        Dim businessUnit As String = ""
        Dim lastRow As Integer
        Dim lastColumn As Integer
        'Dim startSpalte As Integer
        Dim vglName As String = ""
        Dim hproj As New clsProjekt
        Dim vproj As clsProjektvorlage
        Dim geleseneProjekte As Integer
        Dim ProjektdauerIndays As Integer = 0
        Dim ok As Boolean = False

        Dim fullProjectNames As New SortedList(Of String, String)
        Dim firstZeile As Excel.Range

        Dim scenarioName As String = appInstance.ActiveWorkbook.Name
        Dim tmpName As String = ""

        ' bestimme den Namen des Szenarios - das ist gleich der NAme der Excel Datei 
        Dim positionIX As Integer = scenarioName.IndexOf(".xls") - 1
        tmpName = ""
        For ih As Integer = 0 To positionIX
            tmpName = tmpName & scenarioName.Chars(ih)
        Next
        scenarioName = tmpName.Trim

        ' Vorbedingung: das Excel File. das importiert werden soll , ist bereits geöffnet 

        zeile = 2
        spalte = 1
        geleseneProjekte = 0

        Dim suchstr(7) As String
        suchstr(ptModuleSpalten.produktlinie) = "Produktlinie"
        suchstr(ptModuleSpalten.name) = "Name"
        suchstr(ptModuleSpalten.projektTyp) = "Projekt-Typ"
        suchstr(ptModuleSpalten.abhaengigVon) = "ist abhängig von"
        suchstr(ptModuleSpalten.strategicFit) = "strat. Bedeutung"
        suchstr(ptModuleSpalten.risiko) = "Risiko der Umsetzung"
        suchstr(ptModuleSpalten.volume) = "Produktions-Volumen"
        suchstr(ptModuleSpalten.budget) = "Budget"


        Dim inputColumns(7) As Integer



        Try
            Dim activeWSListe As Excel.Worksheet = CType(appInstance.ActiveWorkbook.Worksheets("Tabelle1"), _
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet)
            With activeWSListe

                firstZeile = CType(.Rows(1), Excel.Range)

                ' jetzt werden die Spalten bestimmt 
                Try
                    For i As Integer = 0 To 7
                        inputColumns(i) = firstZeile.Find(What:=suchstr(i)).Column
                    Next
                Catch ex As Exception

                End Try

                lastColumn = firstZeile.End(XlDirection.xlToLeft).Column
                lastColumn = CType(.Cells(1, 10000), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlToLeft).Column
                lastRow = CType(.Cells(2000, 1), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlUp).Row





                While zeile <= lastRow
                    ok = False

                    pName = CStr(CType(.Cells(zeile, inputColumns(ptModuleSpalten.name)), Global.Microsoft.Office.Interop.Excel.Range).Value)
                    vorlagenName = "Projekt-Platzhalter"

                    ' jetzt muss das Start bzw. Ende Date für das Projekt bestimmt werden
                    ' es ist bestimmt durch das erste auftretende Datum bzw. das letzte auftretende Datum
                    Dim projectStartDate As Date = StartofCalendar.AddYears(100)
                    Dim projectEndDate As Date = StartofCalendar.AddYears(-100)

                    Dim firstC As Integer = inputColumns.Max + 1
                    Dim lastC As Integer = lastColumn
                    Dim anzahlPhasenToAdd As Integer = CInt((lastC - firstC + 1) / 5)
                    Dim allesOK As Boolean
                    Dim ignore As Boolean

                    For i As Integer = 1 To anzahlPhasenToAdd
                        Dim tmpDate As Date
                        Dim chkName As String
                        tmpDate = CDate(CType(.Cells(zeile, firstC + 1 + (i - 1) * 5), Global.Microsoft.Office.Interop.Excel.Range).Value)

                        Try
                            chkName = CStr(CType(.Cells(zeile, firstC + (i - 1) * 5), Global.Microsoft.Office.Interop.Excel.Range).Value).Trim
                        Catch ex As Exception
                            ignore = True
                            chkName = ""
                        End Try


                        If DateDiff(DateInterval.Day, StartofCalendar, tmpDate) < 0 Or chkName = "-" Then
                            ignore = True
                        Else
                            ignore = False
                        End If

                        If Not ignore Then
                            If DateDiff(DateInterval.Day, projectStartDate, tmpDate) < 0 Then
                                projectStartDate = tmpDate
                            End If

                            tmpDate = CDate(CType(.Cells(zeile, firstC + 2 + (i - 1) * 5), Global.Microsoft.Office.Interop.Excel.Range).Value)
                            If DateDiff(DateInterval.Day, projectEndDate, tmpDate) > 0 Then
                                projectEndDate = tmpDate
                            End If
                        End If

                    Next


                    If Projektvorlagen.Liste.ContainsKey(vorlagenName) Then

                        vproj = Projektvorlagen.getProject(vorlagenName)
                        Try

                            start = projectStartDate
                            ende = projectEndDate
                            dauer = calcDauerIndays(start, ende)
                            budget = CDbl(CType(.Cells(zeile, inputColumns(ptModuleSpalten.budget)), Global.Microsoft.Office.Interop.Excel.Range).Value)
                            risk = CDbl(CType(.Cells(zeile, inputColumns(ptModuleSpalten.risiko)), Global.Microsoft.Office.Interop.Excel.Range).Value)
                            sfit = CDbl(CType(.Cells(zeile, inputColumns(ptModuleSpalten.strategicFit)), Global.Microsoft.Office.Interop.Excel.Range).Value)
                            volume = CDbl(CType(.Cells(zeile, inputColumns(ptModuleSpalten.volume)), Global.Microsoft.Office.Interop.Excel.Range).Value)
                            complexity = 0.2
                            businessUnit = CStr(CType(.Cells(zeile, inputColumns(ptModuleSpalten.produktlinie)), Global.Microsoft.Office.Interop.Excel.Range).Value)
                            description = ""
                            'vglName = pName.Trim & "#" & ""
                            vglName = calcProjektKey(pName.Trim, scenarioName)


                            If DateDiff(DateInterval.Day, StartofCalendar, start) >= 0 Then

                                If DateDiff(DateInterval.Day, start, ende) > 0 Then
                                    ' nichts tun , Ende-Datum ist ein gültiges Datum
                                    ok = True
                                ElseIf DateDiff(DateInterval.Day, StartofCalendar, ende) >= 0 Then
                                    ' auch Ende ist ein gültiges Datum , liegt nur vor Start
                                    ' also vertauschen der beiden 
                                    Dim tmpDate As Date = ende
                                    ende = start
                                    start = tmpDate
                                    ok = True
                                Else
                                    ' Ende Datum wird anhand der Laufzeit der Vorlage oder der Dauer berechnet
                                    If dauer > 0 Then
                                        ProjektdauerIndays = dauer
                                    Else
                                        ProjektdauerIndays = vproj.dauerInDays
                                    End If
                                    ende = calcDatum(start, ProjektdauerIndays)
                                    ok = True
                                End If

                            ElseIf DateDiff(DateInterval.Day, StartofCalendar, ende) >= 0 Then
                                ' hier ist Start kein gültiges Datum innerhalb der Projekt-Tafel 
                                ' Start Datum wird anhand der Laufzeit der Vorlage berechnet
                                If dauer > 0 Then
                                    ProjektdauerIndays = -1 * dauer
                                Else
                                    ProjektdauerIndays = -1 * vproj.dauerInDays
                                End If

                                start = calcDatum(ende, ProjektdauerIndays)

                                If DateDiff(DateInterval.Day, StartofCalendar, start) >= 0 Then
                                    ' Start ist ein korrektes Datum 
                                    ok = True
                                Else
                                    CType(.Cells(zeile, spalte + 1), Global.Microsoft.Office.Interop.Excel.Range).Value = "Start liegt vor Kalender-Start "
                                    ok = False
                                End If

                            Else
                                CType(.Cells(zeile, spalte + 1), Global.Microsoft.Office.Interop.Excel.Range).Value = "ungültiges Start- und Ende-Datum"
                                ok = False
                            End If

                        Catch ex As Exception
                            CType(.Cells(zeile, spalte + 1), Global.Microsoft.Office.Interop.Excel.Range).Value = ".?."
                            ok = False
                        End Try


                    Else
                        CType(.Cells(zeile, spalte + 1), Global.Microsoft.Office.Interop.Excel.Range).Value = ".?."
                        ok = False
                    End If

                    ' jetzt die Aktion durchführen, wenn alles ok 
                    If ok Then
                        If AlleProjekte.Containskey(vglName) Then
                            ' nichts tun ...
                            Call MsgBox("Projekt aus Inventur Liste existiert bereits - keine Neuanlage")
                        Else
                            Try
                                fullProjectNames.Add(vglName, vglName)
                                'Projekt anlegen ,Verschiebung um 
                                hproj = New clsProjekt(start, start.AddMonths(-1), start.AddMonths(1))

                                Dim capacityNeeded As String = ""
                                hproj = erstelleInventurProjekt(pName, vorlagenName, scenarioName, _
                                                             start, ende, budget, zeile, sfit, risk, _
                                                             capacityNeeded, Nothing, businessUnit, description)

                                If Not IsNothing(hproj) Then
                                    projectStartDate = start
                                    projectEndDate = ende
                                Else
                                    ok = False
                                End If

                            Catch ex As Exception
                                ok = False
                            End Try


                        End If
                    End If

                    If ok Then

                        Dim phaseName As String = ""
                        Dim scaleRule As Integer
                        Dim moduleNames() As String
                        Dim moduleName As String
                        Dim allNames As String
                        Dim planModul As clsProjektvorlage

                        ' jetzt müssen die Module ergänzt werden 
                        For i As Integer = 1 To anzahlPhasenToAdd

                            start = CDate(CType(.Cells(zeile, firstC + 1 + (i - 1) * 5), Global.Microsoft.Office.Interop.Excel.Range).Value)
                            ende = CDate(CType(.Cells(zeile, firstC + 2 + (i - 1) * 5), Global.Microsoft.Office.Interop.Excel.Range).Value)


                            Dim startOffset As Integer = CInt(DateDiff(DateInterval.Day, projectStartDate, start))
                            Dim endOffset As Integer = CInt(DateDiff(DateInterval.Day, projectStartDate, ende))


                            Try
                                phaseName = CStr(CType(.Cells(zeile, firstC + (i - 1) * 5), Global.Microsoft.Office.Interop.Excel.Range).Value).Trim
                                If phaseName = "-" Or endOffset - startOffset = 0 Then
                                    allesOK = False
                                    phaseName = "-"
                                Else
                                    allesOK = True
                                End If
                            Catch ex As Exception
                                allesOK = False
                            End Try

                            Dim parentPhase As clsPhase = Nothing



                            If allesOK Then

                                '
                                ' jetzt muss die aufnehmende Phase erstmal angelegt werden 
                                '
                                If Not IsNothing(phaseName) Then

                                    If phaseName.Length > 0 Then

                                        parentPhase = New clsPhase(parent:=hproj)
                                        parentPhase.nameID = hproj.hierarchy.findUniqueElemKey(phaseName, False)
                                        parentPhase.changeStartandDauer(startOffset, calcDauerIndays(start, ende))

                                        hproj.AddPhase(parentPhase, origName:=phaseName, _
                                               parentID:=rootPhaseName)

                                    End If

                                End If


                                scaleRule = CInt(CType(.Cells(zeile, firstC + 3 + (i - 1) * 5), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                allNames = CStr(CType(.Cells(zeile, firstC + 4 + (i - 1) * 5), Global.Microsoft.Office.Interop.Excel.Range).Value)

                                ' jetzt müssen die einzelnen Module ausgelesen werden 
                                ' aber nur, wenn überhaupt was drin steht und das auch als Modul existiert ...
                                '
                                If Not IsNothing(allNames) Then

                                    If Not allNames.Trim.Length = 0 Then

                                        moduleNames = allNames.Split(New Char() {CChar("#")}, 20)
                                        Dim anzahl As Integer = moduleNames.Length

                                        For ix As Integer = 1 To anzahl
                                            moduleName = moduleNames(ix - 1)
                                            If ModulVorlagen.Contains(moduleName) Then
                                                planModul = ModulVorlagen.getProject(moduleName)

                                                If Not IsNothing(parentPhase) Then

                                                    planModul.moduleCopyTo(hproj, parentPhase.nameID, moduleName, startOffset, endOffset, True)

                                                End If
                                            End If
                                        Next

                                    End If

                                End If

                            End If
                        Next

                        ' jetzt die Projekt eintragen 
                        If Not hproj Is Nothing Then
                            Try
                                ImportProjekte.Add(hproj, False)
                                myCollection.Add(calcProjektKey(hproj))
                            Catch ex As Exception

                            End Try

                        End If

                    End If

                    zeile = zeile + 1

                End While





            End With
        Catch ex As Exception
            Throw New Exception("Fehler in Datei Module Import ...")
        End Try


        ' jetzt noch ein Szenario anlegen, wenn ImportProjekte was enthält 
        If ImportProjekte.Count > 0 Then
            Call storeSessionConstellation(scenarioName, fullProjectNames)
        End If

        currentConstellationName = scenarioName

    End Sub

    ''' <summary>
    ''' ergänzt das übergebene Projekt um die im Ruleset angegebenen Phasen und Meilensteine
    ''' Wenn die Phase mit Namen ruleset.name schon existiert, werden die Elemente hinzugefügt, sofern sie nicht mit demselben Namen in dieser Phase bereits auftreten
    ''' Andernfalls wird bestimmt, wie lange die Phase sein muss
    ''' </summary>
    ''' <param name="hproj"></param>
    ''' <param name="addElementSet"></param>
    ''' <remarks></remarks>
    Public Sub awinApplyAddOnRules(ByRef hproj As clsProjekt, ByVal addElementSet As clsAddElements)

        Dim phaseName As String = ""
        Dim topPhaseName As String = ""
        Dim breadCrumb As String = ""
        Dim milestoneName As String = ""
        Dim elemID As String

        Dim topPhase As clsPhase
        Dim cMilestone As clsMeilenstein


        ' erst bestimmen, ob die Phase schon existiert 
        topPhaseName = addElementSet.name
        topPhase = hproj.getPhase(topPhaseName)

        Dim minDate As Date = Date.Now.AddYears(100)
        Dim maxDate As Date = Date.Now.AddYears(-100)

        Dim currentDate As Date
        Dim currentElem As clsAddElementRules
        Dim currentMS As clsMeilenstein
        Dim currentPH As clsPhase
        Dim index As Integer = 1

        ' hier muss bestimmt werden,wie groß die aufnehmende Phase werden soll 
        ' es werden Mindate und MAxdate bestimmt 
        '
        Do While index <= addElementSet.count
            currentElem = addElementSet.getRule(index)

            Dim anzRules As Integer = currentElem.count
            Dim currentRule As clsAddElementRuleItem
            Dim found As Boolean = False

            Dim i As Integer = 1

            Do While i <= currentElem.count And Not found
                currentRule = currentElem.getItem(i)

                With currentRule
                    If .referenceIsPhase Then
                        ' existiert die Phase überhaupt? wenn nicht , weiter zu nächster Regel
                        Dim type As Integer = -1
                        Dim pvName As String = ""
                        Call splitHryFullnameTo2(.referenceName, phaseName, breadCrumb, type, pvName)

                        If type = -1 Or _
                            (type = PTProjektType.projekt And pvName = hproj.name) Or _
                            (type = PTProjektType.vorlage And pvName = hproj.VorlagenName) Then

                            currentPH = hproj.getPhase(name:=phaseName, breadcrumb:=breadCrumb, lfdNr:=1)

                            If Not IsNothing(currentPH) Then
                                found = True
                                If .referenceDateIsStart Then
                                    currentDate = currentPH.getStartDate.AddDays(currentRule.offset)
                                Else
                                    currentDate = currentPH.getEndDate.AddDays(currentRule.offset)
                                End If

                                If DateDiff(DateInterval.Day, minDate, currentDate) < 0 Then
                                    minDate = currentDate
                                End If

                                If currentElem.elemToCreateIsPhase Then
                                    currentDate = currentDate.AddDays(currentElem.duration)
                                End If
                                If DateDiff(DateInterval.Day, maxDate, currentDate) > 0 Then
                                    maxDate = currentDate
                                End If

                            Else
                                i = i + 1
                            End If

                        Else
                            i = i + 1
                        End If

                        
                    Else
                        Dim type As Integer = -1
                        Dim pvName As String = ""
                        Call splitHryFullnameTo2(.referenceName, milestoneName, breadCrumb, type, pvName)

                        If type = -1 Or _
                            (type = PTProjektType.projekt And pvName = hproj.name) Or _
                            (type = PTProjektType.vorlage And pvName = hproj.VorlagenName) Then

                            currentMS = hproj.getMilestone(milestoneName, breadCrumb, 1)

                            If Not IsNothing(currentMS) Then
                                found = True
                                currentDate = currentMS.getDate.AddDays(currentRule.offset)

                                If DateDiff(DateInterval.Day, minDate, currentDate) < 0 Then
                                    minDate = currentDate
                                End If

                                If currentElem.elemToCreateIsPhase Then
                                    currentDate = currentDate.AddDays(currentElem.duration)
                                End If

                                If DateDiff(DateInterval.Day, maxDate, currentDate) > 0 Then
                                    maxDate = currentDate
                                End If
                            Else
                                i = i + 1
                            End If

                        Else
                            i = i + 1
                        End If
                        
                    End If
                End With

            Loop

            index = index + 1

        Loop

        ' jetzt wird die oberste Phase entsprechend aufgenommen 
        '
        Dim startOffset As Integer = DateDiff(DateInterval.Day, hproj.startDate, minDate)
        If startOffset < 0 Then
            minDate = hproj.startDate
            startOffset = 0
        End If

        Dim duration As Integer = DateDiff(DateInterval.Day, minDate, maxDate) + 1

        If IsNothing(topPhase) Then
            ' die Phase existiert noch nicht
            elemID = hproj.hierarchy.findUniqueElemKey(topPhaseName, False)

            topPhase = New clsPhase(parent:=hproj)

            topPhase.nameID = elemID
            topPhase.changeStartandDauer(startOffset, duration)

            ' der Aufbau der Hierarchie erfolgt in addphase
            hproj.AddPhase(topPhase, origName:="", _
                           parentID:=rootPhaseName)

        Else

            elemID = topPhase.nameID
            ' die Phase existiert bereits; aber ist sie auch ausreichend dimensioniert ? 
            ' ggf werden Start und Dauer angepasst 
            If startOffset <> topPhase.startOffsetinDays Or duration <> topPhase.dauerInDays Then
                topPhase.changeStartandDauer(startOffset, duration)
            End If

        End If


        ' jetzt müssen die Meilensteine / anderen Plan-Elemente eingetragen werden 
        '
        index = 1
        Do While index <= addElementSet.count

            Dim offs As Integer = 1
            Dim wasSuccessful As Boolean = False
            Dim newItemDate As Date
            Dim referenceMS As clsMeilenstein = Nothing
            Dim referencePH As clsPhase = Nothing
            Dim referenceDate As Date
            Dim currentRule As clsAddElementRuleItem

            currentElem = addElementSet.getRule(index)

            ' soll ein Meilenstein oder eine Phase erzeugt werden ? 
            If currentElem.elemToCreateIsPhase Then
                ' es soll eine Phase erzeugt werden 
            Else
                ' es soll ein Meilenstein erzeugt werden 
                Dim found As Boolean = False

                If IsNothing(topPhase.getMilestone(currentElem.name)) Then
                    ' nur wenn der nicht schon existiert, soll er auch erzeugt werden ... 

                    Do While offs <= currentElem.count And Not found
                        Dim ok As Boolean = False
                        currentRule = currentElem.getItem(offs)

                        If currentRule.referenceIsPhase Then
                            Dim type As Integer = -1
                            Dim pvName As String = ""
                            Call splitHryFullnameTo2(currentRule.referenceName, phaseName, breadCrumb, type, pvName)
                            referencePH = hproj.getPhase(name:=phaseName, breadcrumb:=breadCrumb)

                            If Not IsNothing(referencePH) Then
                                If currentRule.referenceDateIsStart Then
                                    referenceDate = referencePH.getStartDate
                                Else
                                    referenceDate = referencePH.getEndDate
                                End If

                                ok = True
                            Else
                                ok = False
                            End If

                        Else
                            Dim type As Integer = -1
                            Dim pvName As String = ""
                            Call splitHryFullnameTo2(currentRule.referenceName, milestoneName, breadCrumb, type, pvName)
                            referenceMS = hproj.getMilestone(msName:=milestoneName, breadcrumb:=breadCrumb)
                            If Not IsNothing(referenceMS) Then
                                referenceDate = referenceMS.getDate
                                ok = True
                            Else
                                ok = False
                            End If
                        End If

                        ' wenn es ein Referenz-Datum gibt ....
                        If ok Then
                            newItemDate = referenceDate.AddDays(currentRule.offset)
                            cMilestone = New clsMeilenstein(parent:=topPhase)
                            elemID = hproj.hierarchy.findUniqueElemKey(currentRule.newElemName, True)

                            Dim cbewertung As clsBewertung = New clsBewertung

                            With cbewertung
                                '.bewerterName = resultVerantwortlich
                                .colorIndex = 0
                                .datum = Date.Now
                                Dim abstandsText As String = ""
                                If currentRule.offset >= 0 Then
                                    abstandsText = "+" & currentRule.offset.ToString & " Tage"
                                Else
                                    abstandsText = currentRule.offset.ToString & " Tage"
                                End If
                                .description = " = " & currentRule.referenceName & abstandsText
                                ' Änderung tk 29.5.16 deliverables ist jetzt Bestandteil von clsMeilenstein
                                '.deliverables = currentElem.deliverables
                            End With


                            With cMilestone
                                .nameID = elemID
                                .setDate = newItemDate
                                If Not cbewertung Is Nothing Then
                                    .addBewertung(cbewertung)
                                End If
                            End With

                            Try
                                With topPhase
                                    .addMilestone(cMilestone)
                                End With
                            Catch ex As Exception

                            End Try
                            found = True
                        Else
                            offs = offs + 1
                        End If

                    Loop


                End If


            End If

            index = index + 1
        Loop

    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="ruleSet"></param>
    ''' <remarks></remarks>
    Public Sub awinReadAddOnRules(ByRef ruleSet As clsAddElements)

        Dim zeile As Integer, spalte As Integer
        Dim newName As String
        Dim duration As Integer = 0
        Dim isPhase As Boolean

        Dim referenceNameMS As String = ""
        Dim referenceNamePH As String = ""
        Dim refISStart As Boolean = True
        Dim abstandsRegel As String = ""
        Dim offset As Integer
        Dim deliverables As String = ""
        Dim newRule As clsAddElementRuleItem
        ' faktor = 1 bedeutet Tage; faktor = 7 bedeutet Wochen 
        Dim faktor As Integer = 1

        Dim lastRow As Integer

        Dim ok As Boolean = False

        Dim firstZeile As Excel.Range

        ' der Name des Rule-Sets wird später der Name der Phase, die ergänzt wird 
        Dim fileName As String = appInstance.ActiveWorkbook.Name
        Dim tmpName As String = ""

        ' bestimme den Namen des Szenarios - das ist gleich der Name der Excel Datei 
        Dim positionIX As Integer = fileName.IndexOf(".xls") - 1
        tmpName = ""
        For ih As Integer = 0 To positionIX
            tmpName = tmpName & fileName.Chars(ih)
        Next
        ruleSet.name = tmpName.Trim

        ' Vorbedingung: das Excel File. das importiert werden soll , ist bereits geöffnet 

        zeile = 2
        spalte = 1

        Try
            Dim activeWSListe As Excel.Worksheet = CType(appInstance.ActiveWorkbook.Worksheets("Tabelle1"), _
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet)
            With activeWSListe

                firstZeile = CType(.Rows(1), Excel.Range)
                lastRow = CType(.Cells(2000, 1), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlUp).Row

                While zeile <= lastRow
                    ok = False

                    Try
                        ' Name des neuen Elements lesen  
                        newName = CStr(CType(.Cells(zeile, 1), Global.Microsoft.Office.Interop.Excel.Range).Value).Trim

                        ' Dauer des neuen Elements lesen; bestimmt damit, ob es sich um eine Phase oder einen MEilenstein handelt
                        Try
                            duration = CInt(CType(.Cells(zeile, 2), Global.Microsoft.Office.Interop.Excel.Range).Value)
                            If duration > 0 Then
                                isPhase = True
                            Else
                                isPhase = False
                            End If
                        Catch ex1 As Exception
                            duration = 0
                            isPhase = False
                        End Try

                        ' Ergebnisse des Meilensteins lesen 
                        deliverables = CStr(CType(.Cells(zeile, 3), Global.Microsoft.Office.Interop.Excel.Range).Value)
                        If IsNothing(deliverables) Then
                            deliverables = ""
                        Else
                            If deliverables.Length > 0 Then
                                deliverables = deliverables.Trim
                            End If
                        End If

                        ' Rollenbedarfe der Phase lesen, spalte 4 

                        ' Kostenbedarfe der Phase lesen , spalte 5

                        ' Referenz-Name des Meilensteins lesen 
                        Try
                            referenceNameMS = CStr(CType(.Cells(zeile, 6), Global.Microsoft.Office.Interop.Excel.Range).Value).Trim
                        Catch ex1 As Exception
                            referenceNameMS = ""
                        End Try


                        ' Referenz-Name der Phase  lesen 
                        Try
                            referenceNamePH = CStr(CType(.Cells(zeile, 7), Global.Microsoft.Office.Interop.Excel.Range).Value).Trim
                        Catch ex1 As Exception
                            referenceNamePH = ""
                        End Try


                        ' Start oder Ende der Phase lesen 
                        Try
                            If CStr(CType(.Cells(zeile, 8), Global.Microsoft.Office.Interop.Excel.Range).Value).Trim = "Ende" Then
                                refISStart = False
                            Else
                                refISStart = True
                            End If
                        Catch ex As Exception
                            refISStart = True
                        End Try

                        abstandsRegel = CStr(CType(.Cells(zeile, 9), Global.Microsoft.Office.Interop.Excel.Range).Value).Trim
                        If abstandsRegel.EndsWith("w") Or tmpName.EndsWith("W") Then
                            faktor = 7
                        Else
                            faktor = 1
                        End If
                        Dim tmpstr() As String

                        tmpstr = abstandsRegel.Trim.Split(New Char() {CChar("w"), CChar("W"), CChar("d"), CChar("D")}, 5)
                        offset = CInt(tmpstr(0)) * faktor

                        ' wenn ein Meilenstein - Name angegeben wurde, wird jetzt die Regel für den Meilenstein angelegt
                        If referenceNameMS.Length > 0 Then
                            newRule = New clsAddElementRuleItem
                            With newRule
                                .newElemName = newName
                                .referenceName = referenceNameMS
                                .referenceIsPhase = False
                                .offset = offset
                            End With

                            If ruleSet.containsElement(newName, isPhase) Then
                                ruleSet.addRule(newRule, isPhase)
                            Else
                                Dim newElem As New clsAddElementRules(newName, isPhase, duration, deliverables)
                                ruleSet.addElem(newElem, isPhase)
                                ruleSet.addRule(newRule, isPhase)
                            End If

                        End If

                        '
                        If referenceNamePH.Length > 0 Then
                            newRule = New clsAddElementRuleItem
                            With newRule
                                .newElemName = newName
                                .referenceName = referenceNamePH
                                .referenceIsPhase = True
                                .referenceDateIsStart = refISStart
                                .offset = offset
                            End With

                            If ruleSet.containsElement(newName, isPhase) Then
                                ruleSet.addRule(newRule, isPhase)
                            Else
                                Dim newElem As New clsAddElementRules(newName, isPhase, duration, deliverables)
                                ruleSet.addElem(newElem, isPhase)
                                ruleSet.addRule(newRule, isPhase)
                            End If

                        End If


                    Catch ex As Exception

                    End Try

                    zeile = zeile + 1

                End While

            End With
        Catch ex As Exception
            Throw New Exception("Fehler in Datei Module Import ...")
        End Try


    End Sub

    Public Sub awinImportProject(ByRef hprojekt As clsProjekt, ByRef hprojTemp As clsProjektvorlage, ByVal isTemplate As Boolean, ByVal importDatum As Date)

        Dim zeile As Integer, spalte As Integer
        Dim hproj As New clsProjekt
        Dim hwert As Integer
        Dim anzFehler As Integer = 0
        Dim ProjektdauerIndays As Integer = 0
        Dim endedateProjekt As Date


        ' Vorbedingung: das Excel File. das importiert werden soll , ist bereits geöffnet 

        zeile = 1
        spalte = 1
        ' ------------------------------------------------------------------------------------------------------
        ' Einlesen der Stammdaten
        ' ------------------------------------------------------------------------------------------------------

        Try
            Dim wsGeneralInformation As Excel.Worksheet = CType(appInstance.ActiveWorkbook.Worksheets("Stammdaten"), _
                Global.Microsoft.Office.Interop.Excel.Worksheet)
            With wsGeneralInformation

                .Unprotect(Password:="x")       ' Blattschutz aufheben

                ' Projekt-Name auslesen
                hproj.name = makeValidProjectName(CType(.Range("Projekt_Name").Value, String))
                hproj.farbe = .Range("Projekt_Name").Interior.Color
                hproj.Schriftfarbe = .Range("Projekt_Name").Font.Color
                hproj.Schrift = CInt(.Range("Projekt_Name").Font.Size)


                ' Kurzbeschreibung, kein Problem, wenn nicht da ...
                Try
                    hproj.description = CType(.Range("ProjektBeschreibung").Value, String)
                Catch ex As Exception

                End Try


                ' Verantwortlich - kein Problem wenn nicht da 
                Try
                    hproj.leadPerson = CType(.Range("Projektleiter").Value, String)
                Catch ex As Exception

                End Try


                ' Start
                hproj.startDate = CType(.Range("StartDatum").Value, Date)

                ' Ende

                endedateProjekt = CType(.Range("EndeDatum").Value, Date)  ' Projekt-Ende für spätere Verwendung merken
                ProjektdauerIndays = calcDauerIndays(hproj.startDate, endedateProjekt)
                Dim startOffset As Long = DateDiff(DateInterval.Day, hproj.startDate, hproj.startDate.AddMonths(0))

                ' Budget
                Try
                    hproj.Erloes = CType(.Range("Budget").Value, Double)
                Catch ex1 As Exception

                End Try


                ' Ampel-Farbe
                hwert = CType(.Range("Bewertung").Value, Integer)

                If hwert >= 0 And hwert <= 3 Then
                    hproj.ampelStatus = hwert
                End If

                ' Ampel-Bewertung 
                hproj.ampelErlaeuterung = CType(.Range("BewertgErläuterung").Value, String)


            End With
        Catch ex As Exception
            Throw New ArgumentException("Fehler in awinImportProject, Lesen Stammdaten")
        End Try

        ' ------------------------------------------------------------------------------------------------------
        ' Einlesen der Attribute
        ' ------------------------------------------------------------------------------------------------------

        Try
            Dim wsAttribute As Excel.Worksheet
            Try
                wsAttribute = CType(appInstance.ActiveWorkbook.Worksheets("Attribute"), _
                   Global.Microsoft.Office.Interop.Excel.Worksheet)
            Catch ex As Exception
                wsAttribute = Nothing
            End Try

            If Not IsNothing(wsAttribute) Then

                With wsAttribute

                    .Unprotect(Password:="x")       ' Blattschutz aufheben


                    '   Varianten-Name
                    Try
                        hproj.variantName = CType(.Range("Variant_Name").Value, String)
                        hproj.variantName = hproj.variantName.Trim
                        If hproj.variantName.Length = 0 Then
                            hproj.variantName = ""
                        End If
                    Catch ex1 As Exception
                        hproj.variantName = ""
                    End Try


                    ' Business Unit - kein Problem wenn nicht da   
                    Try
                        hproj.businessUnit = CType(.Range("Business_Unit").Value, String)
                    Catch ex As Exception

                    End Try

                    ' Status    ist ein read-only Feld
                    hproj.Status = ProjektStatus(1)
                    ' hproj.Status = .Range("Status").Value

                    ' Risiko
                    hproj.Risiko = CDbl(.Range("Risiko").Value)


                    ' Strategic Fit
                    hproj.StrategicFit = CDbl(.Range("Strategischer_Fit").Value)


                    '' Komplexitätszahl - kein Problem, wenn nicht da  --- BMW---
                    'Try
                    '    hproj.complexity = CType(.Range("Complexity").Value, Double)
                    'Catch ex As Exception
                    '    hproj.complexity = 0.5 ' Default
                    'End Try

                    '' Volumen - kein Problem, wenn nicht da    --- BMW ---
                    'Try
                    '    hproj.volume = CType(.Range("Volume").Value, Double)
                    'Catch ex As Exception
                    '    hproj.volume = 10 ' Default
                    'End Try



                End With
            End If
        Catch ex As Exception
            Throw New ArgumentException("Fehler in awinImportProject, Lesen Attribute")
        End Try


        ' ------------------------------------------------------------------------------------------------------
        ' Einlesen der Ressourcen
        ' ------------------------------------------------------------------------------------------------------
        Dim wsRessourcen As Excel.Worksheet
        Try
            wsRessourcen = CType(appInstance.ActiveWorkbook.Worksheets("Ressourcen"), _
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet)
        Catch ex As Exception
            wsRessourcen = Nothing
            ' ------------------------------------------------------------------------------------------------------
            ' Erzeugen und eintragen der Projekt-Phase (= erste Phase mit Dauer des Projekts)
            ' ------------------------------------------------------------------------------------------------------
            Try
                Dim cphase As New clsPhase(hproj)

                ' ProjektPhase wird erzeugt
                cphase = New clsPhase(parent:=hproj)
                cphase.nameID = rootPhaseName

                ' Phasen Dauer wird gleich der Dauer des Projekts gesetzt
                With cphase
                    .nameID = rootPhaseName
                    Dim startOffset As Integer = 0
                    .changeStartandDauer(startOffset, ProjektdauerIndays)
                End With
                ' ProjektPhase wird hinzugefügt
                hproj.AddPhase(cphase)

            Catch ex1 As Exception
                Throw New ArgumentException("Fehler in awinImportProject, Erzeugen ProjektPhase")
            End Try

        End Try

        If Not IsNothing(wsRessourcen) Then

            Try
                With wsRessourcen
                    Dim rng As Excel.Range
                    Dim zelle As Excel.Range
                    Dim chkPhase As Boolean = True
                    Dim chkRolle As Boolean = True
                    Dim firsttime As Boolean = False
                    Dim added As Boolean = True
                    Dim Xwerte As Double()
                    Dim crole As clsRolle
                    Dim cphase As New clsPhase(hproj)
                    Dim ccost As clsKostenart
                    Dim phaseName As String = ""

                    Dim anfang As Integer, ende As Integer  ', projDauer As Integer

                    Dim farbeAktuell As Object
                    Dim r As Integer, k As Integer


                    .Unprotect(Password:="x")       ' Blattschutz aufheben


                    Dim tmpws As Excel.Range = CType(wsRessourcen.Range("Phasen_des_Projekts"), Excel.Range)

                    rng = .Range("Phasen_des_Projekts")

                    If Not (CStr(CType(.Range("Phasen_des_Projekts").Cells(1), Excel.Range).Value) = hproj.name Or _
                           CStr(CType(.Range("Phasen_des_Projekts").Cells(1), Excel.Range).Value) = ".") Then

                        ' ProjektPhase wird hinzugefügt
                        cphase = New clsPhase(parent:=hproj)
                        added = False


                        ' Phasen Dauer wird gleich der Dauer des Projekts gesetzt
                        With cphase
                            .nameID = rootPhaseName
                            Dim startOffset As Integer = 0
                            .changeStartandDauer(startOffset, ProjektdauerIndays)
                            Dim phaseStartdate As Date = .getStartDate
                            Dim phaseEnddate As Date = .getEndDate
                            firsttime = True
                        End With
                        'Call MsgBox("Projektnamen/Phasen Konflikt in awinImportProjekt" & vbLf & "Problem wurde behoben")

                    End If

                    zeile = 0

                    For Each zelle In rng

                        zeile = zeile + 1

                        ' nachsehen, ob Phase angegeben oder Rolle/Kosten
                        If Len(CType(zelle.Value, String)) > 0 Then
                            phaseName = CType(zelle.Value, String).Trim
                            If phaseName = "." Then
                                phaseName = rootPhaseName
                            End If
                        Else
                            phaseName = ""
                        End If

                        ' hier wird die Rollen bzw Kosten Information ausgelesen
                        Dim hname As String
                        Try
                            hname = CType(zelle.Offset(0, 1).Value, String).Trim
                        Catch ex1 As Exception
                            hname = ""
                        End Try

                        If Len(phaseName) > 0 And Len(hname) <= 0 Then
                            chkPhase = True
                            chkRolle = False
                            If Not firsttime Then
                                firsttime = True
                            End If
                        End If

                        If Len(phaseName) <= 0 And Len(hname) > 0 Then
                            If zeile = 1 Then
                                Call MsgBox(" es fehlt die ProjektPhase")
                            Else
                                chkPhase = False
                                chkRolle = True
                            End If
                        Else
                        End If

                        If Len(phaseName) > 0 And Len(hname) > 0 Then
                            chkPhase = True
                            chkRolle = True
                        End If

                        If Len(phaseName) <= 0 And Len(hname) <= 0 Then
                            chkPhase = False
                            chkRolle = False
                            ' beim 1.mal: abspeichern der letzten Phase mit Ihren Rollen
                            ' beim 2.mal: for - Schleife abbrechen
                        End If

                        Select Case chkPhase
                            Case True
                                If Not added Then
                                    hproj.AddPhase(cphase)
                                End If

                                cphase = New clsPhase(parent:=hproj)
                                added = False

                                ' Auslesen der Phasen Dauer
                                anfang = 1  ' anfang enthält den rel.Anfang einer Phase
                                Try
                                    While CInt(zelle.Offset(0, anfang + 1).Interior.ColorIndex) = -4142 And
                                        Not (CType(zelle.Offset(0, anfang + 1).Value, String) = "x")
                                        anfang = anfang + 1
                                    End While
                                Catch ex As Exception
                                    Throw New ArgumentException("Es wurden keine oder falsche Angaben zur Phasendauer der Phase '" & phaseName & "' gemacht." & vbLf &
                                                                "Bitte überprüfen Sie dies.")
                                End Try

                                ende = anfang + 1

                                If CInt(zelle.Offset(0, anfang + 1).Interior.ColorIndex) = -4142 Then
                                    While CType(zelle.Offset(0, ende + 1).Value, String) = "x"
                                        ende = ende + 1
                                    End While
                                    ende = ende - 1
                                Else
                                    farbeAktuell = zelle.Offset(0, anfang + 1).Interior.Color
                                    While CInt(zelle.Offset(0, ende + 1).Interior.Color) = CInt(farbeAktuell)

                                        ende = ende + 1
                                    End While
                                    ende = ende - 1
                                End If

                                With cphase
                                    If phaseName = hproj.name Or phaseName = rootPhaseName Then
                                        .nameID = rootPhaseName
                                        ' nichts tun, die erste Phase hat dann schon ihren richtigen Namen 
                                    Else
                                        .nameID = hproj.hierarchy.findUniqueElemKey(phaseName, False)
                                    End If

                                    ' Änderung 28.11.13: jetzt wird die Phasen Länge exakt bestimmt , über startoffset in Tagen und dauerinDays als Länge
                                    Dim startOffset As Long
                                    Dim dauerIndays As Long
                                    startOffset = DateDiff(DateInterval.Day, hproj.startDate, hproj.startDate.AddMonths(anfang - 1))
                                    dauerIndays = calcDauerIndays(hproj.startDate.AddDays(startOffset), ende - anfang + 1, True)

                                    .changeStartandDauer(startOffset, dauerIndays)
                                    .offset = 0

                                    ' hier muss eine Routine aufgerufen werden, die die Dauer in Tagen berechnet !!!!!!
                                    Dim phaseStartdate As Date = .getStartDate
                                    Dim phaseEnddate As Date = .getEndDate

                                End With
                                Select Case chkRolle
                                    Case True
                                        Throw New ArgumentException("Rollen/Kosten-Bedarfe zur Phase '" & phaseName & "' bitte in die darauffolgenden Zeilen eintragen")
                                    Case False  ' es wurde nur eine Phase angegeben: korrekt

                                End Select

                            Case False ' auslesen Rollen- bzw. Kosten-Information

                                Select Case chkRolle
                                    Case True
                                        ' hier wird die Rollen bzw Kosten Information ausgelesen
                                        '
                                        ' entweder nun Rollen/Kostendefinition oder Ende der Phasen
                                        '
                                        If RoleDefinitions.containsName(hname) Then
                                            Try
                                                r = CInt(RoleDefinitions.getRoledef(hname).UID)

                                                ReDim Xwerte(ende - anfang)


                                                Dim m As Integer
                                                For m = anfang To ende

                                                    Try
                                                        Xwerte(m - anfang) = CDbl(zelle.Offset(0, m + 1).Value)
                                                    Catch ex As Exception
                                                        Xwerte(m - anfang) = 0.0
                                                    End Try

                                                Next m
                                                ' tk: das muss doch eigentlich heissen: end-anfag !? 
                                                'crole = New clsRolle(ende - anfang + 1)
                                                crole = New clsRolle(ende - anfang)
                                                With crole
                                                    .RollenTyp = r
                                                    .Xwerte = Xwerte
                                                End With

                                                With cphase
                                                    .addRole(crole)
                                                End With
                                            Catch ex As Exception
                                                '
                                                ' handelt es sich um die Kostenart Definition?
                                                '
                                            End Try

                                        ElseIf CostDefinitions.containsName(hname) Then

                                            Try

                                                k = CInt(CostDefinitions.getCostdef(hname).UID)

                                                ReDim Xwerte(ende - anfang)

                                                Dim m As Integer
                                                For m = anfang To ende
                                                    Try
                                                        Xwerte(m - anfang) = CDbl(zelle.Offset(0, m + 1).Value)
                                                    Catch ex As Exception
                                                        Xwerte(m - anfang) = 0.0
                                                    End Try

                                                Next m

                                                ' Änderung tk: 26.7 
                                                'ccost = New clsKostenart(ende - anfang + 1)
                                                ccost = New clsKostenart(ende - anfang)
                                                With ccost
                                                    .KostenTyp = k
                                                    .Xwerte = Xwerte
                                                End With


                                                With cphase
                                                    .AddCost(ccost)
                                                End With

                                            Catch ex As Exception

                                            End Try

                                        End If

                                    Case False  ' es wurde weder Phase noch Rolle angegeben. 
                                        If firsttime Then
                                            firsttime = False
                                        Else 'beim 2. mal: letzte Phase hinzufügen; ENDE von For-Schleife for each Zelle
                                            hproj.AddPhase(cphase)
                                            Exit For
                                        End If

                                End Select

                        End Select

                    Next zelle


                End With
            Catch ex As Exception
                Throw New ArgumentException("Fehler in awinImportProject, Lesen Ressourcen von '" & hproj.name & "' " & vbLf & ex.Message)
            End Try

        End If

        '' hier wurde jetzt die Reihenfolge geändert - erst werden die Phasen Definitionen eingelesen ..

        '' jetzt werden die Daten für die Phasen sowie die Termine/Deliverables eingelesen 

        Try
            Dim wsTermine As Excel.Worksheet
            Try
                wsTermine = CType(appInstance.ActiveWorkbook.Worksheets("Termine"), _
                                                             Global.Microsoft.Office.Interop.Excel.Worksheet)
            Catch ex As Exception
                wsTermine = Nothing
            End Try

            If Not IsNothing(wsTermine) Then
                Try
                    With wsTermine
                        Dim lastrow As Integer
                        Dim lastcolumn As Integer
                        Dim phaseNameID As String
                        Dim milestoneName As String
                        Dim milestoneDate As Date
                        Dim resultVerantwortlich As String = ""
                        Dim bewertungsAmpel As Integer
                        Dim explanation As String
                        Dim bewertungsdatum As Date = importDatum
                        Dim Nummer As String
                        Dim tbl As Excel.Range
                        Dim sortBereich As Excel.Range
                        Dim sortKey As Excel.Range
                        Dim rowOffset As Integer
                        Dim columnOffset As Integer


                        .Unprotect(Password:="x")       ' Blattschutz aufheben

                        tbl = .ListObjects("ErgebnTabelle").Range
                        rowOffset = tbl.Row             ' ist die erste Zeile der ErgebnTabelle = Überschriftszeile
                        columnOffset = tbl.Column

                        ' hiermit soll die Tabelle der Termine nach der laufenden Nummer sortiert werden

                        lastrow = CInt(CType(.Cells(2000, columnOffset), Excel.Range).End(XlDirection.xlUp).Row)
                        lastcolumn = CInt(CType(.Cells(rowOffset, 2000), Excel.Range).End(XlDirection.xlToLeft).Column)

                        'sortBereich ist der Inhalt der ErgebnTabelle
                        sortBereich = .Range(.Cells(rowOffset + 1, columnOffset), .Cells(lastrow, lastcolumn))
                        ' sortKey ist die erste Spalte der ErgebnTabelle
                        sortKey = .Range(.Cells(rowOffset + 1, columnOffset), .Cells(lastrow, columnOffset))

                        With .Sort
                            ' Bestehende Sortierebenen löschen
                            .SortFields.Clear()
                            ' Sortierung nach der laufenden Nummer in der ErgebnTabelle also erste Spalte 
                            .SortFields.Add(Key:=sortKey, Order:=XlSortOrder.xlAscending)
                            .SetRange(sortBereich)
                            .Apply()
                        End With

                        For zeile = rowOffset + 1 To lastrow


                            Dim cMilestone As clsMeilenstein
                            Dim cBewertung As clsBewertung
                            Dim cphase As clsPhase
                            Dim objectName As String
                            Dim startDate As Date, endeDate As Date
                            Dim bezug As String
                            Dim errMessage As String = ""


                            Dim isPhase As Boolean = False
                            Dim isMeilenstein As Boolean = False
                            Dim cphaseExisted As Boolean = True

                            Try
                                ' Wenn es keine Phasen gibt in diesem Projekt, so wird trotzdem die Phase1, die ProjektPhase erzeugt.

                                If hproj.AllPhases.Count = 0 Then
                                    Dim duration As Integer
                                    Dim offset As Integer

                                    ' Erzeuge ProjektPhase mit Länge des Projekts
                                    cphase = New clsPhase(parent:=hproj)
                                    cphase.nameID = rootPhaseName
                                    'cphaseExisted = False       ' Phase existiert noch nicht

                                    offset = 0

                                    If ProjektdauerIndays < 1 Or offset < 0 Then
                                        Throw New Exception("unzulässige Angaben für Offset und Dauer: " & _
                                                            offset.ToString & ", " & duration.ToString)
                                    End If

                                    cphase.changeStartandDauer(offset, ProjektdauerIndays)
                                    hproj.AddPhase(cphase)

                                End If                            'Phase 1 ist nun angelegt


                                Try
                                    Nummer = CType(CType(.Cells(zeile, columnOffset), Excel.Range).Value, String).Trim
                                Catch ex As Exception
                                    Nummer = Nothing
                                    Exit For ' Ende der For-Schleife, wenn keine laufende Nummer mehr existiert
                                End Try

                                Try
                                    ' bestimme, worum es sich handelt: Phase oder Meilenstein
                                    objectName = CType(CType(.Cells(zeile, columnOffset + 1), Excel.Range).Value, String).Trim
                                Catch ex As Exception
                                    objectName = Nothing
                                    Throw New Exception("In Tabelle 'Termine' ist der PhasenName nicht angegeben ")
                                    Exit For ' Ende der For-Schleife, wenn keine laufende Nummer mehr existiert
                                End Try


                                If PhaseDefinitions.Contains(objectName) Then
                                    isPhase = True
                                    isMeilenstein = False
                                Else
                                    If objectName = "." Or objectName = hproj.name Then
                                        isPhase = True
                                        isMeilenstein = False
                                    Else
                                        isPhase = False
                                        isMeilenstein = True
                                    End If
                                End If


                                Try
                                    bezug = CType(CType(.Cells(zeile, columnOffset + 2), Excel.Range).Value, String).Trim
                                Catch ex As Exception
                                    bezug = Nothing
                                End Try

                                ' ur: 12.01.2015: Änderung, damit Meilensteine, die den gleichen Namen haben wie Phasen, trotzdem als Meilensteine erkannt werden.
                                '                 gilt aktuell aber nur für den BMW-Import
                                If awinSettings.importTyp = 2 Then
                                    If PhaseDefinitions.Contains(objectName) _
                                        And bezug <> "" _
                                        And Not IsNothing(bezug) Then

                                        isPhase = False
                                        isMeilenstein = True
                                    End If
                                End If

                                Try
                                    startDate = CDate(CType(.Cells(zeile, columnOffset + 3), Excel.Range).Value)
                                Catch ex As Exception
                                    startDate = Date.MinValue
                                End Try

                                Try
                                    endeDate = CDate(CType(.Cells(zeile, columnOffset + 4), Excel.Range).Value)
                                Catch ex As Exception
                                    endeDate = Date.MinValue
                                End Try


                                If DateDiff(DateInterval.Day, hproj.startDate, startDate) < 0 Then
                                    ' kein gültiges Startdatum angegeben

                                    If startDate <> Date.MinValue Then
                                        cphase = Nothing
                                        Throw New Exception("Die Phase '" & objectName & "' beginnt vor dem Projekt !" & vbLf &
                                                     "Bitte korrigieren Sie dies in der Datei'" & hproj.name & ".xlsx'")
                                    Else
                                        ' objectName ist ein Meilenstein
                                        ' Fehlermeldung entfernt ur: 27.05.2014

                                        'If endeDate = Date.MinValue Then
                                        '    Throw New Exception("für den Meilenstein '" & objectName & "'" & vbLf & "wurde im Projekt '" & hproj.name & "' kein Datum eingetragen!")
                                        'End If
                                        If bezug = "." Or bezug = hproj.name Then
                                            cphase = hproj.getPhaseByID(rootPhaseName)
                                        Else
                                            cphase = hproj.getPhase(bezug)
                                        End If

                                        If IsNothing(cphase) Then
                                            If hproj.AllPhases.Count > 0 Then
                                                cphase = hproj.getPhase(1)
                                            Else
                                                ' Erzeuge ProjektPhase mit Länge des Projekts


                                            End If

                                        End If
                                    End If

                                    'isPhase = False

                                Else
                                    'objectName ist eine Phase
                                    'isPhase = True

                                    ' ist der Phasen Name in der Liste der definitionen überhaupt bekannt ? 
                                    If Not PhaseDefinitions.Contains(objectName) Then

                                        ' jetzt noch prüfen, ob es sich um die Phase (1) handelt, dann kann sie ja nicht in der PhaseDefinitions enthalten sein  ..
                                        If Not (hproj.name = objectName Or objectName = ".") Then
                                            Throw New Exception("Phase '" & objectName & "' ist nicht definiert!" & vbLf &
                                                           "Bitte löschen Sie diese Phase aus '" & hproj.name & "'.xlsx, Tabellenblatt 'Termine'")

                                        End If

                                    End If

                                    ' an dieser stelle ist sichergestellt, daß der Phasen Name bekannt ist
                                    ' Prüfen, ob diese Phase bereits in hproj über das ressourcen Sheet angelegt wurde 
                                    ' tk: dieser Befehl holt jetzt die erste Phase mit deisem NAmen, berücksichtigt aber noch nicht die Position ind er Hierarchie; 
                                    ' das muss noch ergänzt werden 
                                    If hproj.name = objectName Or objectName = "." Then
                                        cphase = hproj.getPhaseByID(rootPhaseName)
                                    Else
                                        cphase = hproj.getPhase(objectName)
                                    End If

                                    If IsNothing(cphase) Then
                                        cphase = New clsPhase(parent:=hproj)
                                        cphase.nameID = hproj.hierarchy.findUniqueElemKey(objectName, False)
                                        cphaseExisted = False       ' Phase existiert noch nicht
                                    End If
                                End If

                                If isPhase Then  'xxxx Phase
                                    Try

                                        Dim duration As Long
                                        Dim offset As Long



                                        duration = calcDauerIndays(startDate, endeDate)
                                        offset = DateDiff(DateInterval.Day, hproj.startDate, startDate)


                                        If duration < 1 Or offset < 0 Then
                                            If startDate = Date.MinValue And endeDate = Date.MinValue Then
                                                Throw New Exception(" zu '" & objectName & "' wurde kein Datum eingetragen!")
                                            Else
                                                Throw New Exception("unzulässige Angaben für Offset und Dauer: " & _
                                                                    offset.ToString & ", " & duration.ToString)
                                            End If
                                        End If

                                        cphase.changeStartandDauer(offset, duration)

                                        ' jetzt wird auf Inkonsistenz geprüft 
                                        Dim inkonsistent As Boolean = False

                                        If cphase.countRoles > 0 Or cphase.countCosts > 0 Then
                                            ' prüfen , ob es Inkonsistenzen gibt ? 
                                            Dim r As Integer
                                            For r = 1 To cphase.countRoles
                                                If cphase.getRole(r).Xwerte.Length <> cphase.relEnde - cphase.relStart + 1 Then
                                                    inkonsistent = True
                                                End If
                                            Next

                                            Dim k As Integer
                                            For k = 1 To cphase.countCosts
                                                If cphase.getCost(k).Xwerte.Length <> cphase.relEnde - cphase.relStart + 1 Then
                                                    inkonsistent = True
                                                End If
                                            Next
                                        End If

                                        If inkonsistent Then
                                            anzFehler = anzFehler + 1
                                            Throw New Exception("Der Import konnte nicht fertiggestellt werden. " & vbLf & "Die Dauer der Phase '" & cphase.name & "'  in 'Termine' ist ungleich der in 'Ressourcen' " & vbLf &
                                                                 "Korrigieren Sie bitte gegebenenfalls diese Inkonsistenz in der Datei '" & vbLf & hproj.name & ".xlsx'")
                                        End If
                                        If Not cphaseExisted Then
                                            hproj.AddPhase(cphase)
                                        End If


                                    Catch ex As Exception
                                        Throw New Exception(ex.Message)
                                    End Try

                                Else


                                    phaseNameID = cphase.nameID
                                    cMilestone = New clsMeilenstein(parent:=cphase)
                                    cBewertung = New clsBewertung

                                    milestoneName = objectName.Trim
                                    milestoneDate = endeDate

                                    ' wenn der freefloat nicht zugelassen ist und der Meilenstein ausserhalb der Phasen-Grenzen liegt 
                                    ' muss abgebrochen werden 

                                    If Not awinSettings.milestoneFreeFloat And _
                                        (DateDiff(DateInterval.Day, cphase.getStartDate, milestoneDate) < 0 Or _
                                         DateDiff(DateInterval.Day, cphase.getEndDate, milestoneDate) > 0) Then
                                        Throw New Exception("Der Meilenstein liegt ausserhalb seiner Phase" & vbLf & _
                                                            milestoneName & " nicht innerhalb " & cphase.name & vbLf & _
                                                                 "Korrigieren Sie bitte diese Inkonsistenz in der Datei '" & vbLf & hproj.name & ".xlsx'")
                                    End If


                                    ' wenn kein Datum angegeben wurde, soll das Ende der Phase als Datum angenommen werden 
                                    If DateDiff(DateInterval.Month, hproj.startDate, milestoneDate) < -1 Then
                                        milestoneDate = hproj.startDate.AddDays(cphase.startOffsetinDays + cphase.dauerInDays)
                                    Else
                                        If DateDiff(DateInterval.Day, endedateProjekt, endeDate) > 0 Then
                                            Call MsgBox("der Meilenstein '" & milestoneName & "' liegt später als das Ende des gesamten Projekts" & vbLf &
                                                        "Bitte korrigieren Sie dies im Tabellenblatt Ressourcen der Datei '" & hproj.name & ".xlsx")
                                        End If

                                    End If

                                    ' resultVerantwortlich = CType(.Cells(zeile, 5).value, String)
                                    bewertungsAmpel = CType(CType(.Cells(zeile, columnOffset + 5), Excel.Range).Value, Integer)
                                    explanation = CType(CType(.Cells(zeile, columnOffset + 6), Excel.Range).Value, String)


                                    If bewertungsAmpel < 0 Or bewertungsAmpel > 3 Then
                                        ' es gibt keine Bewertung
                                        bewertungsAmpel = 0
                                    End If
                                    ' damit Kriterien auch eingelesen werden, wenn noch keine Bewertung existiert ...
                                    With cBewertung
                                        '.bewerterName = resultVerantwortlich
                                        .colorIndex = bewertungsAmpel
                                        .datum = importDatum
                                        .description = explanation
                                    End With



                                    With cMilestone
                                        .setDate = milestoneDate
                                        '.verantwortlich = resultVerantwortlich
                                        .nameID = hproj.hierarchy.findUniqueElemKey(milestoneName, True)
                                        If Not cBewertung Is Nothing Then
                                            .addBewertung(cBewertung)
                                        End If
                                    End With


                                    Try
                                        With hproj.getPhaseByID(phaseNameID)
                                            .addMilestone(cMilestone)
                                        End With
                                    Catch ex1 As Exception

                                    End Try



                                End If

                            Catch ex As Exception
                                ' letzte belegte Zeile wurde bereits bearbeitet.
                                zeile = lastrow + 1 ' erzwingt das Ende der For - Schleife
                                Nummer = Nothing
                                Throw New Exception(ex.Message)
                            End Try

                        Next

                    End With
                Catch ex As Exception
                    Throw New Exception(ex.Message)
                End Try

            End If
            If anzFehler > 0 Then
                Call MsgBox("Anzahl Fehler bei Import der Termine von " & hproj.name & " : " & anzFehler)
            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

        If isTemplate Then
            ' hier müssen die Werte für die Vorlage übergeben werden.
            Dim projVorlage As New clsProjektvorlage
            projVorlage.VorlagenName = hproj.name
            projVorlage.Schrift = hproj.Schrift
            projVorlage.Schriftfarbe = hproj.Schriftfarbe
            projVorlage.farbe = hproj.farbe
            projVorlage.earliestStart = -6
            projVorlage.latestStart = 6
            projVorlage.AllPhases = hproj.AllPhases
            hprojTemp = projVorlage

        Else
            hprojekt = hproj
        End If

    End Sub

    ''' <summary>
    ''' liest einen ProjektSteckbrief mit Hierarchie ein 
    ''' Außerdem gibt es die Spalte Summe, in der die Summe der Kosten enthalten sein kann.
    ''' </summary>
    ''' <param name="hprojekt"></param>
    ''' <param name="hprojTemp"></param>
    ''' <param name="isTemplate"></param>
    ''' <param name="importDatum"></param>
    ''' <remarks></remarks>
    Public Sub awinImportProjectmitHrchy(ByRef hprojekt As clsProjekt, ByRef hprojTemp As clsProjektvorlage, ByVal isTemplate As Boolean, ByVal importDatum As Date)

        Dim zeile As Integer, spalte As Integer
        Dim hproj As New clsProjekt
        Dim ProjektdauerIndays As Integer = 0
        Dim endedateProjekt As Date

        Dim projektAmpelFarbe As Integer
        Dim projektAmpelText As String

        ' Vorbedingung: das Excel File. das importiert werden soll , ist bereits geöffnet

        Try

            zeile = 1
            spalte = 1
            ' ------------------------------------------------------------------------------------------------------
            ' Einlesen der Stammdaten
            ' ------------------------------------------------------------------------------------------------------

            Try
                Dim wsGeneralInformation As Excel.Worksheet = CType(appInstance.ActiveWorkbook.Worksheets("Stammdaten"), _
                    Global.Microsoft.Office.Interop.Excel.Worksheet)
                With wsGeneralInformation

                    .Unprotect(Password:="x")       ' Blattschutz aufheben

                    ' Projekt-Name auslesen
                    hproj.name = makeValidProjectName(CType(.Range("Projekt_Name").Value, String))
                    hproj.farbe = .Range("Projekt_Name").Interior.Color
                    hproj.Schriftfarbe = .Range("Projekt_Name").Font.Color
                    hproj.Schrift = CInt(.Range("Projekt_Name").Font.Size)


                    ' Kurzbeschreibung, kein Problem, wenn nicht da ...
                    Try
                        hproj.description = CType(.Range("ProjektBeschreibung").Value, String)
                    Catch ex As Exception

                    End Try


                    ' Verantwortlich - kein Problem wenn nicht da 
                    Try
                        hproj.leadPerson = CType(.Range("Projektleiter").Value, String)
                    Catch ex As Exception

                    End Try


                    ' Start
                    hproj.startDate = CType(.Range("StartDatum").Value, Date)

                    ' Ende

                    endedateProjekt = CType(.Range("EndeDatum").Value, Date)  ' Projekt-Ende für spätere Verwendung merken
                    ProjektdauerIndays = calcDauerIndays(hproj.startDate, endedateProjekt)
                    Dim startOffset As Long = DateDiff(DateInterval.Day, hproj.startDate, hproj.startDate.AddMonths(0))

                    ' Budget
                    Try
                        hproj.Erloes = CType(.Range("Budget").Value, Double)
                    Catch ex1 As Exception

                    End Try


                    ' Ampel-Farbe
                    projektAmpelFarbe = CType(.Range("Bewertung").Value, Integer)
                    If projektAmpelFarbe >= 0 And projektAmpelFarbe <= 3 Then
                        ' zulässiger Wert
                    Else
                        projektAmpelFarbe = 0
                    End If


                    ' Ampel-Bewertung 
                    projektAmpelText = CType(.Range("BewertgErläuterung").Value, String)
                    ' das kann jetzt noch gar nicht zugewiesen werden, weil es noch keine Phasen gibt
                    ' Ampel-Beschreibung und Farbe ist jetzt Attribut der Phase(1), der Projekt-Phase
                    'hproj.ampelErlaeuterung = ampelText


                End With
            Catch ex As Exception
                Call logfileSchreiben("Fehler in awinImportProjectmitHrchy, Lesen Stammdaten", hproj.name, anzFehler)
                Throw New ArgumentException("Fehler in awinImportProjectmitHrchy, Lesen Stammdaten")
            End Try

            ' ------------------------------------------------------------------------------------------------------
            ' Einlesen der Attribute
            ' ------------------------------------------------------------------------------------------------------

            Try
                Dim wsAttribute As Excel.Worksheet
                Try
                    wsAttribute = CType(appInstance.ActiveWorkbook.Worksheets("Attribute"), _
                       Global.Microsoft.Office.Interop.Excel.Worksheet)
                Catch ex As Exception
                    wsAttribute = Nothing
                End Try

                If Not IsNothing(wsAttribute) Then

                    With wsAttribute

                        .Unprotect(Password:="x")       ' Blattschutz aufheben


                        '   Varianten-Name
                        Try
                            hproj.variantName = CType(.Range("Variant_Name").Value, String)
                            hproj.variantName = hproj.variantName.Trim
                            If hproj.variantName.Length = 0 Then
                                hproj.variantName = ""
                            End If
                        Catch ex1 As Exception
                            hproj.variantName = ""
                        End Try

                        '   Varianten-Beschreibung
                        Try
                            hproj.variantDescription = ""
                            Dim tmprng As Excel.Range = CType(.Range("Variant_Description"), Excel.Range)
                            If Not IsNothing(tmprng) Then
                                If Not IsNothing(tmprng.Value) Then
                                    hproj.variantDescription = CType(.Range("Variant_Description").Value, String)
                                End If
                            End If

                            'If Not IsNothing(hproj.variantDescription) Then
                            '    hproj.variantDescription = hproj.variantDescription.Trim
                            'Else
                            '    hproj.variantDescription = ""
                            'End If

                        Catch ex1 As Exception
                            hproj.variantDescription = ""
                        End Try

                        ' Business Unit - kein Problem wenn nicht da   
                        Try
                            hproj.businessUnit = CType(.Range("Business_Unit").Value, String)
                        Catch ex As Exception

                        End Try

                        ' Status    ist ein read-only Feld
                        hproj.Status = ProjektStatus(1)
                        ' hproj.Status = .Range("Status").Value

                        ' Risiko
                        hproj.Risiko = CDbl(.Range("Risiko").Value)


                        ' Strategic Fit
                        hproj.StrategicFit = CDbl(.Range("Strategischer_Fit").Value)


                        ' Ergänzung tk 19.5 es können hier auch sogenannte Custom Fields eingelesen werden ...
                        Try
                            Dim cfRange As Excel.Range = CType(.Range("IndivName2"), Excel.Range)
                            Dim startzeile As Integer = cfRange.Row
                            Dim cfValueColumn As Integer = cfRange.Column
                            Dim lastZeile As Integer = CInt(CType(.Cells(10000, 2), Excel.Range).End(XlDirection.xlUp).Row)

                            ' jetzt die Custom-Fields einlesen 
                            For i As Integer = startzeile To lastZeile

                                Try

                                    Dim cfName As String = CStr(CType(.Cells(i, cfValueColumn - 1), Excel.Range).Value).Trim
                                    Dim cfUid As Integer = customFieldDefinitions.getUid(cfName)

                                    If cfUid > -1 Then ' dann existiert diese Custom Field Definition 
                                        Dim cfType As Integer = customFieldDefinitions.getTyp(cfUid)

                                        If Not IsNothing(cfType) Then
                                            Select Case cfType
                                                Case ptCustomFields.Str
                                                    Dim cfvalue As String = CStr(CType(.Cells(i, cfValueColumn), Excel.Range).Value)
                                                    hproj.addSetCustomSField(cfUid, cfvalue)
                                                Case ptCustomFields.Dbl
                                                    Dim cfvalue As Double = CDbl(CType(.Cells(i, cfValueColumn), Excel.Range).Value)
                                                    hproj.addSetCustomDField(cfUid, cfvalue)
                                                Case ptCustomFields.bool
                                                    Dim cfvalue As Boolean = CBool(CType(.Cells(i, cfValueColumn), Excel.Range).Value)
                                                    hproj.addSetCustomBField(cfUid, cfvalue)
                                                Case Else
                                                    ' Custom Field Type nicht bekannt ...
                                                    Call logfileSchreiben("unbekanntes Custom-Field, wird ignoriert: ", hproj.name & " " & cfName & "," & cfType, anzFehler)
                                            End Select
                                        Else
                                            ' Custom Field UID nicht existent ...
                                            Call logfileSchreiben("uid von Custom-Field existiert nicht ...", hproj.name & " " & cfName & "," & cfUid, anzFehler)
                                        End If
                                    Else
                                        ' Custom Field Definition nicht bekannt ...
                                        Call logfileSchreiben("unbekanntes Custom-Field, wird ignoriert: ", hproj.name & " " & cfName, anzFehler)
                                    End If

                                Catch ex As Exception

                                End Try

                            Next


                        Catch ex As Exception

                        End Try



                    End With
                End If
            Catch ex As Exception
                Call logfileSchreiben("Fehler in awinImportProjectmitHrchy, Lesen Attribute", hproj.name, anzFehler)
                Throw New ArgumentException("Fehler in awinImportProjectmitHrchy, Lesen Attribute")
            End Try


            ' ------------------------------------------------------------------------------------------------------
            ' Einlesen der Termine ur: 06.10.2015: nun vor dem Einlesen der Phasen
            ' ------------------------------------------------------------------------------------------------------


            Try
                Dim wsTermine As Excel.Worksheet
                Try
                    wsTermine = CType(appInstance.ActiveWorkbook.Worksheets("Termine"), _
                                                                 Global.Microsoft.Office.Interop.Excel.Worksheet)
                Catch ex As Exception
                    wsTermine = Nothing
                End Try

                If Not IsNothing(wsTermine) Then
                    Try
                        With wsTermine
                            Dim lastrow As Integer
                            Dim phaseNameID As String
                            Dim milestoneName As String
                            Dim milestoneDate As Date
                            Dim resultVerantwortlich As String = ""
                            Dim bewertungsAmpel As Integer
                            Dim explanation As String
                            Dim deliverables As String
                            Dim bewertungsdatum As Date = importDatum
                            Dim tbl As Excel.Range
                            Dim rowOffset As Integer
                            Dim columnOffset As Integer


                            .Unprotect(Password:="x")       ' Blattschutz aufheben

                            tbl = .Range("ErgebnTabelle")
                            rowOffset = tbl.Row
                            columnOffset = tbl.Column

                            lastrow = CInt(CType(.Cells(2000, columnOffset), Excel.Range).End(XlDirection.xlUp).Row)

                            ' ur: 12.05.2015: hier wurde die Sortierung der ErgebnTabelle entfernt

                            Dim cphase As New clsPhase(parent:=hproj)
                            Dim lastPhase As New clsPhase(parent:=hproj)
                            Dim breadCrumb As String = ""
                            Dim lastLevel As Integer = 0
                            Dim lasthrchynode As New clsHierarchyNode
                            Dim lastelemID As String = ""
                            Dim hilfselemID As String = ""


                            ' die beiden ersten Spalten verbinden, sofern nicht schon gemacht und abspeichern
                            Dim verbRange As Excel.Range
                            Dim iv As Integer

                            For iv = 0 To lastrow - rowOffset + 1
                                verbRange = .Range(.Cells(rowOffset + iv, columnOffset), .Cells(rowOffset + iv, columnOffset + 1))
                                verbRange.Merge()
                            Next


                            For zeile = rowOffset To lastrow


                                Dim cMilestone As clsMeilenstein
                                Dim cBewertung As clsBewertung

                                Dim objectName As String
                                Dim startDate As Date, endeDate As Date
                                ' 
                                Dim errMessage As String = ""
                                Dim aktLevel As Integer = 0

                                Dim isPhase As Boolean = False
                                Dim isMeilenstein As Boolean = False
                                Dim cphaseExisted As Boolean = True

                                Dim duration As Long
                                Dim offset As Long


                                Try
                                    ' String aus erster Spalte der Tabelle lesen

                                    objectName = CStr(CType(.Cells(zeile, columnOffset), Excel.Range).Value).Trim

                                    ' Level abfragen

                                    Dim x As Integer = CInt(CType(.Cells(zeile, columnOffset), Excel.Range).IndentLevel)
                                    If x Mod einrückTiefe <> 0 Then
                                        Call logfileSchreiben("Fehler, Lesen Termine: die Einrückung ist keine durch '" & CStr(einrückTiefe) & "' teilbare Zahl", hproj.name, anzFehler)
                                        Throw New ArgumentException("Fehler, Lesen Termine: die Einrückung ist keine durch '" & CStr(einrückTiefe) & "' teilbare Zahl")
                                    End If
                                    aktLevel = CInt(x / einrückTiefe)

                                Catch ex As Exception
                                    objectName = Nothing
                                    Call logfileSchreiben("Fehler, Lesen Termine: In Tabelle 'Termine' ist der PhasenName nicht angegeben ", hproj.name, anzFehler)
                                    Throw New Exception("Fehler, Lesen Termine: In Tabelle 'Termine' ist der PhasenName nicht angegeben ")
                                    Exit For ' Ende der For-Schleife, wenn keine laufende Nummer mehr existiert
                                End Try

                                ' erste Zeile gelesen; muss RootPhasename sein
                                If zeile = rowOffset Then

                                    If (aktLevel <> 0 And objectName <> elemNameOfElemID(rootPhaseName)) Then
                                        Call logfileSchreiben("Fehler, Lesen Termine: In Tabelle 'Termine' fehlt die ProjektPhase '.' !", hproj.name, anzFehler)
                                        Throw New Exception("Fehler, Lesen Termine: In Tabelle 'Termine' fehlt die ProjektPhase '.' !")
                                        Exit For ' Ende der For-Schleife, wenn keine laufende Nummer mehr existiert
                                    Else
                                        ' erzeuge ProjektPhase rootPhaseName
                                        isPhase = True
                                        isMeilenstein = False
                                        Try
                                            startDate = CDate(CType(.Cells(zeile, columnOffset + 2), Excel.Range).Value)
                                        Catch ex As Exception
                                            startDate = Date.MinValue
                                        End Try
                                        Try
                                            endeDate = CDate(CType(.Cells(zeile, columnOffset + 3), Excel.Range).Value)
                                        Catch ex As Exception
                                            endeDate = Date.MinValue
                                        End Try

                                        ' ProjektPhase wird erzeugt
                                        cphase = New clsPhase(parent:=hproj)


                                        ' Phasen Dauer wird gleich der Dauer des Projekts gesetzt
                                        With cphase
                                            .nameID = rootPhaseName

                                            duration = calcDauerIndays(startDate, endeDate)
                                            offset = DateDiff(DateInterval.Day, hproj.startDate, startDate)

                                            If duration < 1 Or offset < 0 Then
                                                If startDate = Date.MinValue And endeDate = Date.MinValue Then
                                                    Call logfileSchreiben("Fehler, Lesen Termine:  zu '" & objectName & "' wurde kein Datum eingetragen!", hproj.name, anzFehler)
                                                    Throw New Exception("Fehler, Lesen Termine:  zu '" & objectName & "' wurde kein Datum eingetragen!")
                                                Else
                                                    Dim exMsg As String = "Fehler, Lesen Termine: unzulässige Angaben für Offset (>=0) und Dauer (>=1): " & _
                                                                        "Offset= " & offset.ToString & _
                                                                        ", Duration= " & duration.ToString & " " & objectName & "; "

                                                    Call logfileSchreiben(exMsg, hproj.name, anzFehler)
                                                    Throw New Exception(exMsg)
                                                End If
                                            End If

                                            ' für die rootPhase muss gelten: offset = startoffset = 0 und duration = ProjektdauerIndays
                                            If duration <> ProjektdauerIndays Or offset <> 0 Then

                                                Dim exMsg As String = "Fehler, Lesen Termine: unzulässige Angaben für Offset und Dauer: der ProjektPhase " & _
                                                                        "Offset= " & offset.ToString & _
                                                                        ", Duration=" & duration.ToString & " " & objectName & "; " & _
                                                                        ", ProjektDauer=" & ProjektdauerIndays.ToString
                                                Call logfileSchreiben(exMsg, hproj.name, anzFehler)
                                                Throw New Exception(exMsg)
                                            Else
                                                Dim startOffset As Integer = 0
                                                .changeStartandDauer(startOffset, ProjektdauerIndays)
                                                Dim phaseStartdate As Date = .getStartDate
                                                Dim phaseEnddate As Date = .getEndDate

                                            End If

                                        End With

                                        ' ProjektPhase wird hinzugefügt
                                        Dim hrchynode As New clsHierarchyNode
                                        hrchynode.elemName = cphase.name
                                        hrchynode.parentNodeKey = ""
                                        hproj.AddPhase(cphase, parentID:=hrchynode.parentNodeKey)
                                        lastPhase = cphase
                                        lastelemID = cphase.nameID
                                    End If

                                Else
                                    ' alle weiteren Phasen oder Meilensteine
                                    Try
                                        startDate = CDate(CType(.Cells(zeile, columnOffset + 2), Excel.Range).Value)
                                    Catch ex As Exception
                                        startDate = Date.MinValue
                                    End Try
                                    Try
                                        endeDate = CDate(CType(.Cells(zeile, columnOffset + 3), Excel.Range).Value)
                                    Catch ex As Exception
                                        endeDate = Date.MinValue
                                    End Try

                                    If startDate = Date.MinValue And endeDate <> Date.MinValue Then
                                        isPhase = False
                                        isMeilenstein = True
                                    ElseIf startDate <> Date.MinValue And endeDate <> Date.MinValue Then

                                        duration = calcDauerIndays(startDate, endeDate)
                                        offset = DateDiff(DateInterval.Day, hproj.startDate, startDate)

                                        If duration < 1 Or offset < 0 Then
                                            If startDate = Date.MinValue And endeDate = Date.MinValue Then
                                                Call logfileSchreiben(("Fehler, Lesen Termine:  zu '" & objectName & "' wurde kein Datum eingetragen!"), hproj.name, anzFehler)
                                                Throw New Exception("Fehler, Lesen Termine:  zu '" & objectName & "' wurde kein Datum eingetragen!")
                                            Else
                                                Dim exmsg As String = "Fehler, Lesen Termine: unzulässige Angaben für Offset und Dauer: " & _
                                                                    offset.ToString & ", " & duration.ToString & ": " & objectName

                                                Call logfileSchreiben(exmsg, hproj.name, anzFehler)
                                                Throw New Exception(exmsg)
                                            End If
                                        End If

                                        isPhase = True
                                        isMeilenstein = False

                                    End If

                                    ' eingelesener String objectname ist eine Phase

                                    If isPhase Then

                                        cphase = New clsPhase(parent:=hproj)

                                        If PhaseDefinitions.Contains(objectName) Or isMissingDefinitionOK(objectName, isTemplate, False) Then

                                            With cphase
                                                .nameID = hproj.hierarchy.findUniqueElemKey(objectName, False)

                                                duration = calcDauerIndays(startDate, endeDate)
                                                offset = DateDiff(DateInterval.Day, hproj.startDate, startDate)

                                                .changeStartandDauer(offset, duration)
                                                Dim phaseStartdate As Date = .getStartDate
                                                Dim phaseEnddate As Date = .getEndDate
                                            End With


                                            Dim hrchynode As New clsHierarchyNode
                                            hrchynode.elemName = cphase.name

                                            If aktLevel = 0 Then
                                                hrchynode.parentNodeKey = ""

                                            ElseIf aktLevel = 1 Then
                                                hrchynode.parentNodeKey = rootPhaseName

                                            ElseIf aktLevel - lastLevel = 1 Then
                                                hrchynode.parentNodeKey = lastelemID

                                            ElseIf aktLevel - lastLevel = 0 Then
                                                hrchynode.parentNodeKey = hproj.hierarchy.getParentIDOfID(lastelemID)

                                            ElseIf lastLevel - aktLevel >= 1 Then
                                                hilfselemID = lastelemID
                                                For l As Integer = 1 To lastLevel - aktLevel
                                                    hilfselemID = hproj.hierarchy.getParentIDOfID(hilfselemID)
                                                Next l
                                                hrchynode.parentNodeKey = hproj.hierarchy.getParentIDOfID(hilfselemID)
                                            Else
                                                Call logfileSchreiben(("Fehler, Lesen Termine: Hierarchie kann nicht richtig aufgebaut werden:" & cphase.nameID), hproj.name, anzFehler)
                                                Throw New ArgumentException("Fehler, Lesen Termine:  Hierarchie kann nicht richtig aufgebaut werden" & cphase.nameID)
                                            End If

                                            hproj.AddPhase(cphase, parentID:=hrchynode.parentNodeKey)
                                            '' ''hproj.hierarchy.addNode(hrchynode, cphase.nameID)
                                            hrchynode.indexOfElem = hproj.AllPhases.Count
                                            ' merken von letzem Element (Knoten,Phase,Meilenstein)
                                            lasthrchynode = hrchynode
                                            lastelemID = cphase.nameID
                                            lastPhase = cphase
                                            lastLevel = aktLevel

                                        Else
                                            ' objectname existiert nicht in den PhaseDefinitions
                                            ' muss in missingPhaseDefinitions noch eingetragen werden
                                            Call logfileSchreiben(("Fehler, Lesen Termine: Phase '" & objectName & "' existiert im CustomizationFile nicht!"), hproj.name, anzFehler)
                                            Throw New ArgumentException("Fehler, Lesen Termine:Phase '" & objectName & "' existiert im CustomizationFile nicht!")
                                        End If

                                    ElseIf isMeilenstein Then

                                        If MilestoneDefinitions.Contains(objectName) Or isMissingDefinitionOK(objectName, isTemplate, True) Then

                                            Dim hrchynode As New clsHierarchyNode
                                            hrchynode.elemName = cphase.name

                                            If aktLevel = 0 Then
                                                ' Fehler, denn Meilenstein kann nicht parallel zu Rootphase sein??
                                                Call logfileSchreiben(("Fehler, Lesen Termine: Hierarchie kann nicht richtig aufgebaut werden:" & vbLf & "Level des Meilensteins ist nicht akzeptabel" & objectName), hproj.name, anzFehler)
                                                Throw New ArgumentException("Fehler, Lesen Termine: Hierarchie kann nicht richtig aufgebaut werden:" & vbLf & "Level des Meilensteins ist nicht akzeptabel" & objectName)

                                            ElseIf aktLevel = 1 Then
                                                phaseNameID = rootPhaseName

                                            ElseIf aktLevel - lastLevel = 1 Then
                                                phaseNameID = lastelemID

                                            ElseIf aktLevel - lastLevel = 0 Then
                                                phaseNameID = hproj.hierarchy.getParentIDOfID(lastelemID)

                                            ElseIf lastLevel - aktLevel >= 1 Then
                                                hilfselemID = lastelemID
                                                For l As Integer = 1 To lastLevel - aktLevel
                                                    hilfselemID = hproj.hierarchy.getParentIDOfID(hilfselemID)
                                                Next l
                                                phaseNameID = hproj.hierarchy.getParentIDOfID(hilfselemID)
                                            Else
                                                Call logfileSchreiben(("Fehler, Lesen Termine: Hierarchie kann nicht richtig aufgebaut werden: Meilenstein " & objectName), hproj.name, anzFehler)
                                                Throw New ArgumentException("Fehler, Lesen Termine:  Hierarchie kann nicht richtig aufgebaut werden: Meilenstein " & objectName)
                                            End If


                                            Dim hilfsPhase As clsPhase = hproj.getPhaseByID(phaseNameID)
                                            cMilestone = New clsMeilenstein(parent:=hproj.getPhaseByID(phaseNameID))
                                            cBewertung = New clsBewertung

                                            milestoneName = objectName.Trim
                                            milestoneDate = endeDate

                                            ' wenn der freefloat nicht zugelassen ist und der Meilenstein ausserhalb der Phasen-Grenzen liegt 
                                            ' muss abgebrochen werden 

                                            If Not awinSettings.milestoneFreeFloat And _
                                                (DateDiff(DateInterval.Day, hilfsPhase.getStartDate, milestoneDate) < 0 Or _
                                                 DateDiff(DateInterval.Day, hilfsPhase.getEndDate, milestoneDate) > 0) Then

                                                Call logfileSchreiben(("Fehler, Lesen Termine: Der Meilenstein liegt ausserhalb seiner Phase" & vbLf & _
                                                                    milestoneName & " nicht innerhalb " & hilfsPhase.name & vbLf & _
                                                                         "Korrigieren Sie bitte diese Inkonsistenz in der Datei '"), hproj.name, anzFehler)
                                                Throw New Exception("Fehler, Lesen Termine: Der Meilenstein liegt ausserhalb seiner Phase" & vbLf & _
                                                                    milestoneName & " nicht innerhalb " & hilfsPhase.name & vbLf & _
                                                                         "Korrigieren Sie bitte diese Inkonsistenz in der Datei '" & vbLf & hproj.name & ".xlsx'")
                                            End If


                                            ' wenn kein Datum angegeben wurde, soll das Ende der Phase als Datum angenommen werden 
                                            If DateDiff(DateInterval.Month, hproj.startDate, milestoneDate) < -1 Then
                                                milestoneDate = hproj.startDate.AddDays(hilfsPhase.startOffsetinDays + hilfsPhase.dauerInDays)
                                            Else
                                                If DateDiff(DateInterval.Day, endedateProjekt, endeDate) > 0 Then
                                                    Call logfileSchreiben(("Fehler, Lesen Termine: der Meilenstein '" & milestoneName & "' liegt später als das Ende des gesamten Projekts" & vbLf &
                                                                "Bitte korrigieren Sie dies im Tabellenblatt Ressourcen der Datei '"), hproj.name & ".xlsx", anzFehler)
                                                End If

                                            End If

                                            ' resultVerantwortlich = CType(.Cells(zeile, 5).value, String)
                                            Try
                                                bewertungsAmpel = CType(CType(.Cells(zeile, columnOffset + 4), Excel.Range).Value, Integer)
                                                If IsNothing(bewertungsAmpel) Then
                                                    bewertungsAmpel = 0
                                                End If
                                            Catch ex As Exception
                                                bewertungsAmpel = 0
                                            End Try

                                            Try
                                                explanation = CType(CType(.Cells(zeile, columnOffset + 5), Excel.Range).Value, String)
                                                If IsNothing(explanation) Then
                                                    explanation = ""
                                                End If
                                            Catch ex As Exception
                                                explanation = ""
                                            End Try

                                            Try
                                                ' Ergänzung tk 2.11 deliverables ergänzt 
                                                deliverables = CType(CType(.Cells(zeile, columnOffset + 6), Excel.Range).Value, String)
                                                If IsNothing(deliverables) Then
                                                    deliverables = ""
                                                End If
                                            Catch ex As Exception
                                                deliverables = ""
                                            End Try



                                            ' tk 29.5.16
                                            ' hier müssen die Deliverables jetzt auseinander dividiert werden in die einzelnen Items
                                            Try
                                                If deliverables.Trim.Length > 0 Then
                                                    Dim splitStr() As String = deliverables.Split(New Char() {CChar(vbLf), CChar(vbCr)}, 100)

                                                    ' tk 29.5.16 Deliverables jetzt als einzelnen Items 
                                                    For ix As Integer = 1 To splitStr.Length
                                                        cMilestone.addDeliverable(splitStr(ix - 1))
                                                    Next
                                                End If
                                            Catch ex As Exception

                                            End Try

                                            If bewertungsAmpel < 0 Or bewertungsAmpel > 3 Then
                                                ' es gibt keine Bewertung
                                                bewertungsAmpel = 0
                                            End If
                                            ' damit Kriterien auch eingelesen werden, wenn noch keine Bewertung existiert ...
                                            With cBewertung
                                                '.bewerterName = resultVerantwortlich
                                                .colorIndex = bewertungsAmpel
                                                .datum = importDatum
                                                .description = explanation
                                            End With



                                            With cMilestone
                                                .setDate = milestoneDate
                                                '.verantwortlich = resultVerantwortlich
                                                .nameID = hproj.hierarchy.findUniqueElemKey(milestoneName, True)
                                                If Not cBewertung Is Nothing Then
                                                    .addBewertung(cBewertung)
                                                End If
                                            End With


                                            Try
                                                With hproj.getPhaseByID(phaseNameID)
                                                    .addMilestone(cMilestone)
                                                End With
                                            Catch ex1 As Exception
                                                Throw New Exception(ex1.Message)
                                            End Try

                                        Else
                                            ' objectname existiert nicht in den PhaseDefinitions
                                            ' muss in missingPhaseDefinitions noch eingetragen werden
                                            Call logfileSchreiben(("Fehler, Lesen Termine: Meilenstein '" & objectName & "' existiert im CustomizationFile nicht!"), hproj.name, anzFehler)
                                            Throw New ArgumentException("Fehler, Lesen Termine:Meilenstein '" & objectName & "' existiert im CustomizationFile nicht!")
                                        End If

                                    End If



                                End If

                            Next zeile
                        End With

                    Catch ex As Exception
                        Call logfileSchreiben("Fehler in awinImportProjectmitHrchy, Lesen Termine: '" & ex.Message, hproj.name, anzFehler)
                        'Throw New ArgumentException("Fehler in awinImportProjectmitHrchy, Lesen Termine von '" & hproj.name & "' " & vbLf & ex.Message)
                        Throw New ArgumentException(ex.Message)

                    End Try


                Else

                    Call MsgBox("keine Termine definiert")
                    Throw New ArgumentException("Es wurden keine Termine definiert! Projekt " & hproj.name & " kann nicht eingelesen werden")
                End If
            Catch ex As Exception
                Call logfileSchreiben("Fehler in awinImportProjectmitHrchy, Lesen Termine: '" & ex.Message, hproj.name, anzFehler)
                Throw New ArgumentException("Fehler in awinImportProjectmitHrchy, Lesen Termine von '" & hproj.name & "' " & vbLf & ex.Message)

            End Try


            ' ------------------------------------------------------------------------------------------------------
            ' Einlesen der Ressourcen
            ' ------------------------------------------------------------------------------------------------------
            Dim wsRessourcen As Excel.Worksheet
            Try
                wsRessourcen = CType(appInstance.ActiveWorkbook.Worksheets("Ressourcen"), _
                                                                Global.Microsoft.Office.Interop.Excel.Worksheet)
            Catch ex As Exception
                wsRessourcen = Nothing
                ' '' '' '' ------------------------------------------------------------------------------------------------------
                ' '' '' '' Erzeugen und eintragen der Projekt-Phase (= erste Phase mit Dauer des Projekts)
                ' '' '' '' ------------------------------------------------------------------------------------------------------
                '' '' ''Try
                '' '' ''    Dim cphase As New clsPhase(hproj)

                '' '' ''    ' ProjektPhase wird erzeugt
                '' '' ''    cphase = New clsPhase(parent:=hproj)
                '' '' ''    cphase.nameID = rootPhaseName

                '' '' ''    ' Phasen Dauer wird gleich der Dauer des Projekts gesetzt
                '' '' ''    With cphase
                '' '' ''        .nameID = rootPhaseName
                '' '' ''        Dim startOffset As Integer = 0
                '' '' ''        .changeStartandDauer(startOffset, ProjektdauerIndays)
                '' '' ''    End With
                '' '' ''    ' ProjektPhase wird hinzugefügt
                '' '' ''    hproj.AddPhase(cphase)

                '' '' ''Catch ex1 As Exception
                '' '' ''    Throw New ArgumentException("Fehler in awinImportProject, Erzeugen ProjektPhase")
                '' '' ''End Try

            End Try

            If Not IsNothing(wsRessourcen) Then

                Try
                    With wsRessourcen
                        Dim rng As Excel.Range
                        Dim zelle As Excel.Range
                        Dim ressSumOffset As Integer = 1
                        Dim ressOff As Integer = 2
                        Dim chkPhase As Boolean = True
                        Dim chkRolle As Boolean = True
                        Dim firsttime As Boolean = False
                        Dim fertig As Boolean = True
                        Dim summe As Double = -1        ' summe = -1: bedeutet, Summe wird nicht verwendet, oder hat einen unsinnigen Wert
                        Dim Xwerte As Double() = Nothing
                        Dim oldXwerte As Double()
                        Dim crole As clsRolle
                        Dim cphase As clsPhase = Nothing
                        Dim lastphase As clsPhase = Nothing
                        Dim lastelemID As String = ""
                        Dim ccost As clsKostenart
                        Dim phaseName As String = ""
                        Dim aktLevel As Integer = 0   'speichert den Level direkt nach dem Lesen der Phase
                        Dim cphaseLevel As Integer = 0 'speichert den Level der momentan in cphase gespeicherten Phase
                        Dim lastlevel As Integer = 0  'speichert den Level des vorausgehenden elements
                        Dim breadcrumb As String = ""
                        Dim anfang As Integer, ende As Integer  ', projDauer As Integer

                        Dim farbeAktuell As Object
                        Dim r As Integer, k As Integer


                        .Unprotect(Password:="x")       ' Blattschutz aufheben


                        Dim tmpws As Excel.Range = CType(wsRessourcen.Range("Phasen_des_Projekts"), Excel.Range)

                        rng = .Range("Phasen_des_Projekts")

                        Dim testrange As Excel.Range = CType(.Cells(1, 2000), Excel.Range)

                        Dim gefundenRange As Excel.Range = testrange.Find(What:="Summe")
                        If IsNothing(gefundenRange) Then
                            ' alte Version des Steckbriefes 
                            ressOff = 1
                            ressSumOffset = -1              ' keine Summe vorhanden
                            Call logfileSchreiben("alte Version des ProjektSteckbriefes: ohne 'Summe'", hproj.name, anzFehler)
                        Else

                            ' die beiden ersten Spalten verbinden, sofern nicht schon gemacht und abspeichern
                            Dim verbRange As Excel.Range
                            Dim iv As Integer

                            For iv = 0 To rng.Rows.Count - 1
                                verbRange = .Range(.Cells(rng.Row + iv, rng.Column), .Cells(rng.Row + iv, rng.Column + 1))
                                verbRange.Merge()
                            Next

                            ressOff = gefundenRange.Column - rng.Column - 1
                            ressSumOffset = gefundenRange.Column - rng.Column - 2
                            'Call logfileSchreiben("neue Version des ProjektSteckbriefes: mit 'Summe'", hproj.name, anzFehler)


                            '' die beiden ersten Spalten verbinden, sofern nicht schon gemacht und abspeichern
                            'Dim verbRange As Excel.Range
                            'Dim iv As Integer

                            'For iv = 0 To rng.Rows.Count - 1
                            '    verbRange = .Range(.Cells(rng.Row + iv, rng.Column), .Cells(rng.Row + iv, rng.Column + 1))
                            '    verbRange.Merge()
                            'Next
                        End If

                        Dim hstr As String = CStr(CType(.Range("Phasen_des_Projekts").Cells(1), Excel.Range).Value)
                        hstr = elemNameOfElemID(rootPhaseName)

                        If CStr(CType(.Range("Phasen_des_Projekts").Cells(1), Excel.Range).Value) <> elemNameOfElemID(rootPhaseName) Then


                            ' ProjektPhase wird hinzugefügt, sofern sie nich
                            cphase = New clsPhase(parent:=hproj)
                            fertig = False


                            ' Phasen Dauer wird gleich der Dauer des Projekts gesetzt
                            With cphase
                                .nameID = rootPhaseName
                                Dim startOffset As Integer = 0
                                .changeStartandDauer(startOffset, ProjektdauerIndays)
                                Dim phaseStartdate As Date = .getStartDate
                                Dim phaseEnddate As Date = .getEndDate
                                firsttime = True
                            End With
                            'Call MsgBox("Projektnamen/Phasen Konflikt in awinImportProjekt" & vbLf & "Problem wurde behoben")

                        End If


                        zeile = 0

                        For Each zelle In rng

                            zeile = zeile + 1



                            ' nachsehen, ob Phase angegeben oder Rolle/Kosten
                            hstr = CStr(zelle.Value)
                            Dim x As Integer = CInt(zelle.IndentLevel)
                            If x Mod einrückTiefe <> 0 Then
                                Call logfileSchreiben("Fehler beim Lesen Ressourcen: die Einrückung ist keine durch '" & CStr(einrückTiefe) & "' teilbare Zahl", hproj.name, anzFehler)
                                Throw New ArgumentException("Fehler beim Lesen Ressourcen: die Einrückung ist keine durch '" & CStr(einrückTiefe) & "' teilbare Zahl")
                            End If
                            aktLevel = CInt(x / einrückTiefe)

                            If Len(CType(zelle.Value, String)) > 0 Then
                                phaseName = CType(zelle.Value, String).Trim
                            Else
                                phaseName = ""
                            End If

                            ' hier wird die Rollen bzw Kosten Information ausgelesen
                            Dim hname As String = ""
                            Try

                                If Not IsNothing(CType(zelle.Offset(0, 1), Excel.Range).Value) Then
                                    hname = CType(CType(zelle.Offset(0, 1), Excel.Range).Value, String).Trim
                                End If

                            Catch ex1 As Exception
                                hname = ""
                            End Try

                            If Len(phaseName) > 0 And Len(hname) <= 0 Then
                                chkPhase = True
                                chkRolle = False
                                If Not firsttime Then
                                    firsttime = True
                                End If
                            End If

                            If Len(phaseName) <= 0 And Len(hname) > 0 Then
                                If zeile = 1 Then
                                    Call MsgBox(" es fehlt die ProjektPhase")
                                Else
                                    chkPhase = False
                                    chkRolle = True
                                End If
                            Else
                            End If

                            If Len(phaseName) > 0 And Len(hname) > 0 Then
                                chkPhase = True
                                chkRolle = True
                            End If

                            If Len(phaseName) <= 0 And Len(hname) <= 0 Then
                                chkPhase = False
                                chkRolle = False
                                ' beim 1.mal: abspeichern der letzten Phase mit Ihren Rollen
                                ' beim 2.mal: for - Schleife abbrechen
                            End If

                            Select Case chkPhase
                                Case True

                                    If Not fertig Then

                                        lastelemID = cphase.nameID
                                        lastphase = cphase
                                        lastlevel = cphaseLevel
                                    End If

                                    ' in cphase wird die Phase mit Namen phaseName, bereits über Termine in der Hierarchie des Projekts eingetragen
                                    ' gespeichert
                                    ' das muss später überprüft werden können, um ggf gleichnamige Phasen in einer Breadcrumb Stufe richtig zuordnen zu können

                                    ' wenn in einer und derselben Hierarchy-Stufe mehrere gleichnamige Phasen vorkommen, so muss später anhand der Liste der 
                                    ' Phase-Nummern geprüft werden, welche denn die richtige Phase ist 
                                    Dim phaseIndex() As Integer

                                    If phaseName = hproj.name Or phaseName = elemNameOfElemID(rootPhaseName) Then

                                        cphase = hproj.getPhaseByID(rootPhaseName)
                                        ReDim phaseIndex(0)
                                        phaseIndex(0) = 1
                                        'das ist derselbe Effekt wie der untenstehende Befehl, nur schneller; und das Ergebnis muss ja gleich sein 
                                        ' phaseIndex = hproj.hierarchy.getPhaseIndices(cphase.name, "")

                                    Else

                                        ' erzeugen des breadcrumb, um nachsehen zu können, ob diese Phase in der gleichen Hierarchiestufe
                                        ' bereits über Termine eingelesen wurde
                                        If aktLevel > lastlevel Then

                                            If breadcrumb = "" Then
                                                breadcrumb = "."
                                            Else
                                                breadcrumb = breadcrumb & "#" & lastphase.name
                                            End If

                                        ElseIf aktLevel = lastlevel Then
                                            ' aktlevel = lastlevel: also nicht tun
                                        Else

                                            While aktLevel < lastlevel
                                                Dim hhstr As String = ""
                                                Dim type As Integer = -1
                                                Dim pvName As String = ""
                                                Call splitHryFullnameTo2(breadcrumb, hhstr, breadcrumb, type, pvName)
                                                lastlevel = lastlevel - 1
                                            End While

                                        End If

                                        ' Prüfung, ob die Phase phaseName in der bereits aus Termine bestehenden Hierarchie mit dem gleiche breadcrumb existiert, sonst Fehler



                                        If Not hproj.hierarchy.containsPhase(phaseName, breadcrumb) Then

                                            ReDim phaseIndex(0)
                                            Call logfileSchreiben("Fehler beim Lesen Ressourcen: bei Phase '" & phaseName & "#" & breadcrumb & "'", hproj.name, anzFehler)
                                            Throw New ArgumentException("Fehler beim Lesen Ressourcen: bei Phase '" & phaseName & "#" & breadcrumb & "'")
                                        Else

                                            phaseIndex = hproj.hierarchy.getPhaseIndices(phaseName, breadcrumb)

                                            cphase = hproj.getPhase(phaseIndex(0))
                                            cphaseLevel = hproj.hierarchy.getIndentLevel(cphase.nameID)

                                        End If

                                    End If

                                    fertig = False

                                    ' ur: 12.10.2015: neu:  Bedarfe nur als Summe angegeben

                                    ' Auslesen der Phasen Dauer und anschließend vergleichen, ob die in Termine mit der in Ressource übereinstimmt
                                    ' d.h. rel.Anfang und rel.Ende müssen übereinstimmen, wenn relStart und relEnde nicht übereinstimmen, so werden Sie einfach so gesetzt.

                                    Dim maxcol As Integer = hproj.anzahlRasterElemente
                                    Dim col As Integer


                                    col = 1
                                    While CInt(zelle.Offset(0, ressOff + col).Interior.ColorIndex) = -4142 And
                                             Not (CType(zelle.Offset(0, ressOff + col).Value, String) = "x") And
                                             col <= maxcol

                                        col = col + 1

                                    End While


                                    If col >= maxcol Then

                                        ' Phase und deren Länge wird nicht dargestellt in Tabellenblatt Ressourcen
                                        anfang = cphase.relStart
                                        ende = cphase.relEnde

                                    Else
                                        ' Phasenlänge wird dargestellt in Tabellenblatt Ressourcen, also überprüfen

                                        anfang = col

                                        Try
                                            ende = anfang + 1

                                            If CInt(zelle.Offset(0, ressOff + anfang).Interior.ColorIndex) = -4142 Then
                                                While CType(zelle.Offset(0, ressOff + ende).Value, String) = "x"
                                                    ende = ende + 1
                                                End While
                                                ende = ende - 1
                                            Else
                                                farbeAktuell = zelle.Offset(0, ressOff + anfang).Interior.Color
                                                While CInt(zelle.Offset(0, ressOff + ende).Interior.Color) = CInt(farbeAktuell)

                                                    ende = ende + 1
                                                End While
                                                ende = ende - 1
                                            End If

                                        Catch ex As Exception
                                            Call logfileSchreiben("Fehler beim Lesen Ressourcen: Es wurden keine oder falsche Angaben zur Phasendauer der Phase '" & phaseName & "' gemacht." & vbLf &
                                                                       "Bitte überprüfen Sie dies.", hproj.name, anzFehler)
                                            Throw New ArgumentException("Fehler beim Lesen Ressourcen: Es wurden keine oder falsche Angaben zur Phasendauer der Phase '" & phaseName & "' gemacht." & vbLf &
                                                                       "Bitte überprüfen Sie dies.")
                                        End Try

                                    End If


                                    ' Prüfung, ob die Phase cphase in Termine und Ressourcen übereinstimmt in relStart und relEnde
                                    Dim rightOneFound As Boolean = (anfang = cphase.relStart And ende = cphase.relEnde)
                                    Dim tmpIX As Integer = 1

                                    If phaseIndex.Length > 1 Then
                                        While Not rightOneFound And tmpIX <= phaseIndex.Length - 1
                                            cphase = hproj.getPhase(phaseIndex(tmpIX))
                                            rightOneFound = (anfang = cphase.relStart And ende = cphase.relEnde)
                                            tmpIX = tmpIX + 1
                                        End While
                                    End If


                                    If Not rightOneFound Then

                                        'Call MsgBox("Fehler beim Lesen der Ressourcen: die Dauer der Phase " & cphase.name & "' ist fehlerhaft")
                                        Throw New ArgumentException("Fehler beim Lesen der Ressourcen: die Dauer der Phase '" & cphase.name & "' ist fehlerhaft")

                                    End If


                                    Select Case chkRolle
                                        Case True
                                            Throw New ArgumentException("Rollen/Kosten-Bedarfe zur Phase '" & phaseName & "' bitte in die darauffolgenden Zeilen eintragen")
                                        Case False  ' es wurde nur eine Phase angegeben: korrekt

                                    End Select


                                Case False ' auslesen Rollen- bzw. Kosten-Information


                                    Select Case chkRolle
                                        Case True

                                            ' hier wird die Rollen bzw Kosten Information ausgelesen
                                            '
                                            ' entweder nun Rollen/Kostendefinition oder Ende der Phasen
                                            '
                                            If RoleDefinitions.containsName(hname) Then
                                                Try
                                                    r = CInt(RoleDefinitions.getRoledef(hname).UID)


                                                    ''ur:12.10.2015: Eingabe einer Summe in Ressourcen nun möglich, 
                                                    Try
                                                        summe = CDbl(zelle.Offset(0, 1 + ressSumOffset).Value)
                                                    Catch ex As Exception
                                                        summe = -1
                                                    End Try

                                                    If summe > 0.0 Then    ' Verteilung der Summe auf die Monate über Dauer der Phase

                                                        ReDim oldXwerte(0)
                                                        oldXwerte(0) = summe

                                                        With cphase

                                                            anfang = .relStart
                                                            ende = .relEnde
                                                            ReDim Xwerte(ende - anfang)

                                                            .berechneBedarfe(.getStartDate, .getEndDate, oldXwerte, 1, Xwerte)
                                                        End With

                                                        ''ur:12.10.2015:  eingefügt

                                                    Else

                                                        '  Anfang Check , ob richtige Kästchen Werte enthalten
                                                        Dim msgstr As String = " Fehler bei der Verteilung benötigter Kapazitäten" & vbCrLf & "für Rolle " & hname & " in Spalte "
                                                        Dim checkok As Boolean = True

                                                        Dim i As Integer
                                                        For i = 1 To hproj.anzahlRasterElemente

                                                            Dim wertvorhanden As Boolean = (CDbl(zelle.Offset(0, i + ressOff).Value) <> 0.0)
                                                            If (i < anfang Or i > ende) And wertvorhanden Then
                                                                msgstr = msgstr & " ," & i
                                                                checkok = False
                                                            End If

                                                        Next
                                                        If Not checkok Then
                                                            Call logfileSchreiben(msgstr, hproj.name, anzFehler)
                                                            'Call MsgBox(msgstr)
                                                            'Throw New ArgumentException(msgstr)
                                                        End If
                                                        ' Ende Check

                                                        ReDim Xwerte(ende - anfang)

                                                        Dim m As Integer
                                                        For m = anfang To ende

                                                            Try
                                                                Xwerte(m - anfang) = CDbl(zelle.Offset(0, m + ressOff).Value)
                                                            Catch ex As Exception
                                                                Xwerte(m - anfang) = 0.0
                                                            End Try

                                                        Next m


                                                    End If

                                                    ' das muss doch eigentlich heissen: ende - anfang !? 
                                                    'crole = New clsRolle(ende - anfang + 1)
                                                    crole = New clsRolle(ende - anfang)
                                                    With crole
                                                        .RollenTyp = r
                                                        .Xwerte = Xwerte
                                                    End With

                                                    With cphase
                                                        .addRole(crole)
                                                    End With

                                                Catch ex As Exception
                                                    Throw New Exception(ex.Message)
                                                End Try

                                            ElseIf CostDefinitions.containsName(hname) Then

                                                Try

                                                    k = CInt(CostDefinitions.getCostdef(hname).UID)

                                                    ''ur:12.10.2015: Eingabe einer Summe in Ressourcen nun möglich, 
                                                    Try
                                                        summe = CDbl(zelle.Offset(0, 1 + ressSumOffset).Value)
                                                    Catch ex As Exception
                                                        summe = -1
                                                    End Try

                                                    If summe > 0.0 Then        'Summe wird verteilt auf Dauer der Phase

                                                        ReDim oldXwerte(0)
                                                        oldXwerte(0) = summe

                                                        With cphase

                                                            anfang = .relStart
                                                            ende = .relEnde
                                                            ReDim Xwerte(ende - anfang)

                                                            .berechneBedarfe(.getStartDate, .getEndDate, oldXwerte, 1, Xwerte)
                                                        End With

                                                    Else


                                                        ''ur:12.10.2015: 
                                                        '  Anfang Check , ob richtige Kästchen Werte enthalten
                                                        Dim msgstr As String = " Fehler bei der Verteilung benötigter Kapazitäten:" & vbCrLf & "für Kostenart " & hname & " in Spalte "
                                                        Dim checkok As Boolean = True

                                                        Dim i As Integer
                                                        For i = 1 To hproj.anzahlRasterElemente

                                                            Dim wertvorhanden As Boolean = (CDbl(zelle.Offset(0, i + ressOff).Value) <> 0.0)
                                                            If (i < anfang Or i > ende) And wertvorhanden Then
                                                                msgstr = msgstr & " ," & i
                                                                checkok = False
                                                            End If

                                                        Next
                                                        If Not checkok Then
                                                            Call logfileSchreiben(msgstr, hproj.name, anzFehler)
                                                            'Call MsgBox(msgstr)
                                                            'Throw New ArgumentException(msgstr)
                                                        End If
                                                        ' Ende Check

                                                        ReDim Xwerte(ende - anfang)
                                                        Dim m As Integer
                                                        For m = anfang To ende
                                                            Try
                                                                Xwerte(m - anfang) = CDbl(zelle.Offset(0, m + ressOff).Value)
                                                            Catch ex As Exception
                                                                Xwerte(m - anfang) = 0.0
                                                            End Try

                                                        Next m

                                                    End If

                                                    'ccost = New clsKostenart(ende - anfang + 1)
                                                    ccost = New clsKostenart(ende - anfang)
                                                    With ccost
                                                        .KostenTyp = k
                                                        .Xwerte = Xwerte
                                                    End With


                                                    With cphase
                                                        .AddCost(ccost)
                                                    End With

                                                Catch ex As Exception
                                                    Throw New Exception(ex.Message)
                                                End Try

                                            End If

                                        Case False  ' es wurde weder Phase noch Rolle angegeben. 
                                            If firsttime Then
                                                firsttime = False
                                            Else 'beim 2. mal:  ENDE von For-Schleife for each Zelle

                                                Exit For
                                            End If

                                    End Select

                            End Select

                        Next zelle


                    End With
                Catch ex As Exception
                    Call logfileSchreiben("Fehler in awinImportProjectmitHrchy, Lesen Ressourcen: " & ex.Message, hproj.name, anzFehler)
                    Throw New ArgumentException("Fehler in awinImportProjectmitHrchy, Lesen Ressourcen von '" & hproj.name & "' " & vbLf & ex.Message)
                End Try

            End If

            ' ------------------------------------------------------------------
            '   Ende Einlesen Ressourcen
            ' -------------------------------------------------------------------

        Catch ex As Exception
            Call logfileSchreiben("Fehler in awinImportProjectmitHrchy " & ex.Message, hproj.name, anzFehler)
            Throw New ArgumentException("Fehler in awinImportProjectmitHrchy '" & hproj.name & "' " & vbLf & ex.Message)
        End Try

        ' da Ampelfarbe , Beschreibung jetzt in Phase ist, muss das hier , nach Einlesen der Phasen
        hproj.ampelStatus = projektAmpelFarbe
        hproj.ampelErlaeuterung = projektAmpelText


        If isTemplate Then
            ' hier müssen die Werte für die Vorlage übergeben werden.
            Dim projVorlage As New clsProjektvorlage
            projVorlage.VorlagenName = hproj.name
            projVorlage.Schrift = hproj.Schrift
            projVorlage.Schriftfarbe = hproj.Schriftfarbe
            projVorlage.farbe = hproj.farbe
            projVorlage.earliestStart = -6
            projVorlage.latestStart = 6
            projVorlage.AllPhases = hproj.AllPhases
            projVorlage.hierarchy = hproj.hierarchy
            hprojTemp = projVorlage

        Else
            hprojekt = hproj
        End If

    End Sub

    ' '' ''' <summary>
    ' '' ''' liest einen ProjektSteckbrief mit Hierarchie ein, vor mit PT113 die Spalte Summe hinzugefügt wurde
    ' '' ''' </summary>
    ' '' ''' <param name="hprojekt"></param>
    ' '' ''' <param name="hprojTemp"></param>
    ' '' ''' <param name="isTemplate"></param>
    ' '' ''' <param name="importDatum"></param>
    ' '' ''' <remarks></remarks>
    ' ''Public Sub awinImportProjectmitHrchy_beforePT113(ByRef hprojekt As clsProjekt, ByRef hprojTemp As clsProjektvorlage, ByVal isTemplate As Boolean, ByVal importDatum As Date)

    ' ''    Dim zeile As Integer, spalte As Integer
    ' ''    Dim hproj As New clsProjekt
    ' ''    Dim hwert As Integer
    ' ''    Dim anzFehler As Integer = 0
    ' ''    Dim ProjektdauerIndays As Integer = 0
    ' ''    Dim endedateProjekt As Date


    ' ''    ' Vorbedingung: das Excel File. das importiert werden soll , ist bereits geöffnet 

    ' ''    zeile = 1
    ' ''    spalte = 1
    ' ''    ' ------------------------------------------------------------------------------------------------------
    ' ''    ' Einlesen der Stammdaten
    ' ''    ' ------------------------------------------------------------------------------------------------------

    ' ''    Try
    ' ''        Dim wsGeneralInformation As Excel.Worksheet = CType(appInstance.ActiveWorkbook.Worksheets("Stammdaten"), _
    ' ''            Global.Microsoft.Office.Interop.Excel.Worksheet)
    ' ''        With wsGeneralInformation

    ' ''            .Unprotect(Password:="x")       ' Blattschutz aufheben

    ' ''            ' Projekt-Name auslesen
    ' ''            hproj.name = CType(.Range("Projekt_Name").Value, String)
    ' ''            hproj.farbe = .Range("Projekt_Name").Interior.Color
    ' ''            hproj.Schriftfarbe = .Range("Projekt_Name").Font.Color
    ' ''            hproj.Schrift = CInt(.Range("Projekt_Name").Font.Size)


    ' ''            ' Kurzbeschreibung, kein Problem, wenn nicht da ...
    ' ''            Try
    ' ''                hproj.description = CType(.Range("ProjektBeschreibung").Value, String)
    ' ''            Catch ex As Exception

    ' ''            End Try


    ' ''            ' Verantwortlich - kein Problem wenn nicht da 
    ' ''            Try
    ' ''                hproj.leadPerson = CType(.Range("Projektleiter").Value, String)
    ' ''            Catch ex As Exception

    ' ''            End Try


    ' ''            ' Start
    ' ''            hproj.startDate = CType(.Range("StartDatum").Value, Date)

    ' ''            ' Ende

    ' ''            endedateProjekt = CType(.Range("EndeDatum").Value, Date)  ' Projekt-Ende für spätere Verwendung merken
    ' ''            ProjektdauerIndays = calcDauerIndays(hproj.startDate, endedateProjekt)
    ' ''            Dim startOffset As Long = DateDiff(DateInterval.Day, hproj.startDate, hproj.startDate.AddMonths(0))

    ' ''            ' Budget
    ' ''            Try
    ' ''                hproj.Erloes = CType(.Range("Budget").Value, Double)
    ' ''            Catch ex1 As Exception

    ' ''            End Try


    ' ''            ' Ampel-Farbe
    ' ''            hwert = CType(.Range("Bewertung").Value, Integer)

    ' ''            If hwert >= 0 And hwert <= 3 Then
    ' ''                hproj.ampelStatus = hwert
    ' ''            End If

    ' ''            ' Ampel-Bewertung 
    ' ''            hproj.ampelErlaeuterung = CType(.Range("BewertgErläuterung").Value, String)


    ' ''        End With
    ' ''    Catch ex As Exception
    ' ''        Throw New ArgumentException("Fehler in awinImportProject, Lesen Stammdaten")
    ' ''    End Try

    ' ''    ' ------------------------------------------------------------------------------------------------------
    ' ''    ' Einlesen der Attribute
    ' ''    ' ------------------------------------------------------------------------------------------------------

    ' ''    Try
    ' ''        Dim wsAttribute As Excel.Worksheet
    ' ''        Try
    ' ''            wsAttribute = CType(appInstance.ActiveWorkbook.Worksheets("Attribute"), _
    ' ''               Global.Microsoft.Office.Interop.Excel.Worksheet)
    ' ''        Catch ex As Exception
    ' ''            wsAttribute = Nothing
    ' ''        End Try

    ' ''        If Not IsNothing(wsAttribute) Then

    ' ''            With wsAttribute

    ' ''                .Unprotect(Password:="x")       ' Blattschutz aufheben


    ' ''                '   Varianten-Name
    ' ''                Try
    ' ''                    hproj.variantName = CType(.Range("Variant_Name").Value, String)
    ' ''                    hproj.variantName = hproj.variantName.Trim
    ' ''                    If hproj.variantName.Length = 0 Then
    ' ''                        hproj.variantName = ""
    ' ''                    End If
    ' ''                Catch ex1 As Exception
    ' ''                    hproj.variantName = ""
    ' ''                End Try


    ' ''                ' Business Unit - kein Problem wenn nicht da   
    ' ''                Try
    ' ''                    hproj.businessUnit = CType(.Range("Business_Unit").Value, String)
    ' ''                Catch ex As Exception

    ' ''                End Try

    ' ''                ' Status    ist ein read-only Feld
    ' ''                hproj.Status = ProjektStatus(1)
    ' ''                ' hproj.Status = .Range("Status").Value

    ' ''                ' Risiko
    ' ''                hproj.Risiko = CDbl(.Range("Risiko").Value)


    ' ''                ' Strategic Fit
    ' ''                hproj.StrategicFit = CDbl(.Range("Strategischer_Fit").Value)


    ' ''                '' Komplexitätszahl - kein Problem, wenn nicht da  --- BMW---
    ' ''                'Try
    ' ''                '    hproj.complexity = CType(.Range("Complexity").Value, Double)
    ' ''                'Catch ex As Exception
    ' ''                '    hproj.complexity = 0.5 ' Default
    ' ''                'End Try

    ' ''                '' Volumen - kein Problem, wenn nicht da    --- BMW ---
    ' ''                'Try
    ' ''                '    hproj.volume = CType(.Range("Volume").Value, Double)
    ' ''                'Catch ex As Exception
    ' ''                '    hproj.volume = 10 ' Default
    ' ''                'End Try



    ' ''            End With
    ' ''        End If
    ' ''    Catch ex As Exception
    ' ''        Throw New ArgumentException("Fehler in awinImportProject, Lesen Attribute")
    ' ''    End Try


    ' ''    ' ------------------------------------------------------------------------------------------------------
    ' ''    ' Einlesen der Ressourcen
    ' ''    ' ------------------------------------------------------------------------------------------------------
    ' ''    Dim wsRessourcen As Excel.Worksheet
    ' ''    Try
    ' ''        wsRessourcen = CType(appInstance.ActiveWorkbook.Worksheets("Ressourcen"), _
    ' ''                                                        Global.Microsoft.Office.Interop.Excel.Worksheet)
    ' ''    Catch ex As Exception
    ' ''        wsRessourcen = Nothing
    ' ''        ' ------------------------------------------------------------------------------------------------------
    ' ''        ' Erzeugen und eintragen der Projekt-Phase (= erste Phase mit Dauer des Projekts)
    ' ''        ' ------------------------------------------------------------------------------------------------------
    ' ''        Try
    ' ''            Dim cphase As New clsPhase(hproj)

    ' ''            ' ProjektPhase wird erzeugt
    ' ''            cphase = New clsPhase(parent:=hproj)
    ' ''            cphase.nameID = rootPhaseName

    ' ''            ' Phasen Dauer wird gleich der Dauer des Projekts gesetzt
    ' ''            With cphase
    ' ''                .nameID = rootPhaseName
    ' ''                Dim startOffset As Integer = 0
    ' ''                .changeStartandDauer(startOffset, ProjektdauerIndays)
    ' ''            End With
    ' ''            ' ProjektPhase wird hinzugefügt
    ' ''            hproj.AddPhase(cphase)

    ' ''        Catch ex1 As Exception
    ' ''            Throw New ArgumentException("Fehler in awinImportProject, Erzeugen ProjektPhase")
    ' ''        End Try

    ' ''    End Try

    ' ''    If Not IsNothing(wsRessourcen) Then

    ' ''        Try
    ' ''            With wsRessourcen
    ' ''                Dim rng As Excel.Range
    ' ''                Dim zelle As Excel.Range
    ' ''                Dim chkPhase As Boolean = True
    ' ''                Dim chkRolle As Boolean = True
    ' ''                Dim firsttime As Boolean = False
    ' ''                Dim added As Boolean = True
    ' ''                Dim Xwerte As Double()
    ' ''                Dim crole As clsRolle
    ' ''                Dim cphase As New clsPhase(hproj)
    ' ''                Dim lastphase As clsPhase
    ' ''                Dim lasthrchyNode As clsHierarchyNode
    ' ''                Dim lastelemID As String = ""
    ' ''                Dim ccost As clsKostenart
    ' ''                Dim phaseName As String = ""
    ' ''                Dim aktLevel As Integer = 0   'speichert den Level direkt nach dem Lesen der Phase
    ' ''                Dim cphaseLevel As Integer = 0 'speichert den Level der momentan in cphase gespeicherten Phase
    ' ''                Dim lastlevel As Integer = 0  'speichert den Level des vorausgehenden elements

    ' ''                Dim anfang As Integer, ende As Integer  ', projDauer As Integer

    ' ''                Dim farbeAktuell As Object
    ' ''                Dim r As Integer, k As Integer


    ' ''                .Unprotect(Password:="x")       ' Blattschutz aufheben


    ' ''                Dim tmpws As Excel.Range = CType(wsRessourcen.Range("Phasen_des_Projekts"), Excel.Range)

    ' ''                rng = .Range("Phasen_des_Projekts")

    ' ''                Dim hstr As String = CStr(CType(.Range("Phasen_des_Projekts").Cells(1), Excel.Range).Value)
    ' ''                hstr = elemNameOfElemID(rootPhaseName)

    ' ''                If CStr(CType(.Range("Phasen_des_Projekts").Cells(1), Excel.Range).Value) <> elemNameOfElemID(rootPhaseName) Then


    ' ''                    ' ProjektPhase wird hinzugefügt
    ' ''                    cphase = New clsPhase(parent:=hproj)
    ' ''                    added = False


    ' ''                    ' Phasen Dauer wird gleich der Dauer des Projekts gesetzt
    ' ''                    With cphase
    ' ''                        .nameID = rootPhaseName
    ' ''                        Dim startOffset As Integer = 0
    ' ''                        .changeStartandDauer(startOffset, ProjektdauerIndays)
    ' ''                        Dim phaseStartdate As Date = .getStartDate
    ' ''                        Dim phaseEnddate As Date = .getEndDate
    ' ''                        firsttime = True
    ' ''                    End With
    ' ''                    'Call MsgBox("Projektnamen/Phasen Konflikt in awinImportProjekt" & vbLf & "Problem wurde behoben")

    ' ''                End If

    ' ''                zeile = 0

    ' ''                For Each zelle In rng

    ' ''                    zeile = zeile + 1

    ' ''                    ' nachsehen, ob Phase angegeben oder Rolle/Kosten
    ' ''                    hstr = CStr(zelle.Value)
    ' ''                    Dim x As Integer = CInt(zelle.IndentLevel)
    ' ''                    If x Mod einrückTiefe <> 0 Then
    ' ''                        Throw New ArgumentException("die Einrückung ist keine durch '" & CStr(einrückTiefe) & "' teilbare Zahl")
    ' ''                    End If
    ' ''                    aktLevel = CInt(x / einrückTiefe)

    ' ''                    If Len(CType(zelle.Value, String)) > 0 Then
    ' ''                        phaseName = CType(zelle.Value, String).Trim
    ' ''                    Else
    ' ''                        phaseName = ""
    ' ''                    End If

    ' ''                    ' hier wird die Rollen bzw Kosten Information ausgelesen
    ' ''                    Dim hname As String
    ' ''                    Try
    ' ''                        hname = CType(zelle.Offset(0, 1).Value, String).Trim
    ' ''                    Catch ex1 As Exception
    ' ''                        hname = ""
    ' ''                    End Try

    ' ''                    If Len(phaseName) > 0 And Len(hname) <= 0 Then
    ' ''                        chkPhase = True
    ' ''                        chkRolle = False
    ' ''                        If Not firsttime Then
    ' ''                            firsttime = True
    ' ''                        End If
    ' ''                    End If

    ' ''                    If Len(phaseName) <= 0 And Len(hname) > 0 Then
    ' ''                        If zeile = 1 Then
    ' ''                            Call MsgBox(" es fehlt die ProjektPhase")
    ' ''                        Else
    ' ''                            chkPhase = False
    ' ''                            chkRolle = True
    ' ''                        End If
    ' ''                    Else
    ' ''                    End If

    ' ''                    If Len(phaseName) > 0 And Len(hname) > 0 Then
    ' ''                        chkPhase = True
    ' ''                        chkRolle = True
    ' ''                    End If

    ' ''                    If Len(phaseName) <= 0 And Len(hname) <= 0 Then
    ' ''                        chkPhase = False
    ' ''                        chkRolle = False
    ' ''                        ' beim 1.mal: abspeichern der letzten Phase mit Ihren Rollen
    ' ''                        ' beim 2.mal: for - Schleife abbrechen
    ' ''                    End If

    ' ''                    Select Case chkPhase
    ' ''                        Case True
    ' ''                            If Not added Then
    ' ''                                ' '' ''  hproj.AddPhase(cphase)

    ' ''                                Dim hrchynode As New clsHierarchyNode
    ' ''                                hrchynode.elemName = cphase.name


    ' ''                                If cphaseLevel = 0 Then
    ' ''                                    hrchynode.parentNodeKey = ""

    ' ''                                ElseIf cphaseLevel = 1 Then
    ' ''                                    hrchynode.parentNodeKey = rootPhaseName

    ' ''                                ElseIf cphaseLevel - lastlevel = 1 Then
    ' ''                                    hrchynode.parentNodeKey = lastelemID

    ' ''                                ElseIf cphaseLevel - lastlevel = 0 Then
    ' ''                                    hrchynode.parentNodeKey = hproj.hierarchy.getParentIDOfID(lastelemID)

    ' ''                                ElseIf lastlevel - cphaseLevel >= 1 Then
    ' ''                                    Dim hilfselemID As String = lastelemID
    ' ''                                    For l As Integer = 1 To lastlevel - cphaseLevel
    ' ''                                        hilfselemID = hproj.hierarchy.getParentIDOfID(hilfselemID)
    ' ''                                    Next l
    ' ''                                    hrchynode.parentNodeKey = hproj.hierarchy.getParentIDOfID(hilfselemID)
    ' ''                                Else
    ' ''                                    Throw New ArgumentException("Fehler beim Import! Hierarchie kann nicht richtig aufgebaut werden")
    ' ''                                End If

    ' ''                                hproj.AddPhase(cphase, parentID:=hrchynode.parentNodeKey)

    ' ''                                ' '' ''hproj.hierarchy.addNode(hrchynode, cphase.nameID)
    ' ''                                hrchynode.indexOfElem = hproj.AllPhases.Count
    ' ''                                ' merken von letzem Element (Knoten,Phase,Meilenstein)
    ' ''                                lasthrchyNode = hrchynode
    ' ''                                lastelemID = cphase.nameID
    ' ''                                lastphase = cphase
    ' ''                                lastlevel = cphaseLevel
    ' ''                            End If


    ' ''                            cphase = New clsPhase(parent:=hproj)
    ' ''                            added = False

    ' ''                            ' Auslesen der Phasen Dauer
    ' ''                            anfang = 1  ' anfang enthält den rel.Anfang einer Phase
    ' ''                            Try
    ' ''                                While CInt(zelle.Offset(0, anfang + 1).Interior.ColorIndex) = -4142 And
    ' ''                                    Not (CType(zelle.Offset(0, anfang + 1).Value, String) = "x")
    ' ''                                    anfang = anfang + 1
    ' ''                                End While
    ' ''                            Catch ex As Exception
    ' ''                                Throw New ArgumentException("Es wurden keine oder falsche Angaben zur Phasendauer der Phase '" & phaseName & "' gemacht." & vbLf &
    ' ''                                                            "Bitte überprüfen Sie dies.")
    ' ''                            End Try

    ' ''                            ende = anfang + 1

    ' ''                            If CInt(zelle.Offset(0, anfang + 1).Interior.ColorIndex) = -4142 Then
    ' ''                                While CType(zelle.Offset(0, ende + 1).Value, String) = "x"
    ' ''                                    ende = ende + 1
    ' ''                                End While
    ' ''                                ende = ende - 1
    ' ''                            Else
    ' ''                                farbeAktuell = zelle.Offset(0, anfang + 1).Interior.Color
    ' ''                                While CInt(zelle.Offset(0, ende + 1).Interior.Color) = CInt(farbeAktuell)

    ' ''                                    ende = ende + 1
    ' ''                                End While
    ' ''                                ende = ende - 1
    ' ''                            End If

    ' ''                            With cphase
    ' ''                                If phaseName = hproj.name Or phaseName = elemNameOfElemID(rootPhaseName) Then
    ' ''                                    .nameID = rootPhaseName
    ' ''                                    ' nichts tun, die erste Phase hat dann schon ihren richtigen Namen 
    ' ''                                Else
    ' ''                                    .nameID = hproj.hierarchy.findUniqueElemKey(phaseName, False)
    ' ''                                End If
    ' ''                                cphaseLevel = aktLevel

    ' ''                                ' Änderung 28.11.13: jetzt wird die Phasen Länge exakt bestimmt , über startoffset in Tagen und dauerinDays als Länge
    ' ''                                Dim startOffset As Long
    ' ''                                Dim dauerIndays As Long
    ' ''                                startOffset = DateDiff(DateInterval.Day, hproj.startDate, hproj.startDate.AddMonths(anfang - 1))
    ' ''                                dauerIndays = calcDauerIndays(hproj.startDate.AddDays(startOffset), ende - anfang + 1, True)

    ' ''                                .changeStartandDauer(startOffset, dauerIndays)
    ' ''                                .offset = 0

    ' ''                                ' hier muss eine Routine aufgerufen werden, die die Dauer in Tagen berechnet !!!!!!
    ' ''                                Dim phaseStartdate As Date = .getStartDate
    ' ''                                Dim phaseEnddate As Date = .getEndDate

    ' ''                            End With
    ' ''                            Select Case chkRolle
    ' ''                                Case True
    ' ''                                    Throw New ArgumentException("Rollen/Kosten-Bedarfe zur Phase '" & phaseName & "' bitte in die darauffolgenden Zeilen eintragen")
    ' ''                                Case False  ' es wurde nur eine Phase angegeben: korrekt

    ' ''                            End Select

    ' ''                        Case False ' auslesen Rollen- bzw. Kosten-Information

    ' ''                            Select Case chkRolle
    ' ''                                Case True
    ' ''                                    ' hier wird die Rollen bzw Kosten Information ausgelesen
    ' ''                                    '
    ' ''                                    ' entweder nun Rollen/Kostendefinition oder Ende der Phasen
    ' ''                                    '
    ' ''                                    If RoleDefinitions.Contains(hname) Then
    ' ''                                        Try
    ' ''                                            r = CInt(RoleDefinitions.getRoledef(hname).UID)

    ' ''                                            ReDim Xwerte(ende - anfang)


    ' ''                                            Dim m As Integer
    ' ''                                            For m = anfang To ende

    ' ''                                                Try
    ' ''                                                    Xwerte(m - anfang) = CDbl(zelle.Offset(0, m + 1).Value)
    ' ''                                                Catch ex As Exception
    ' ''                                                    Xwerte(m - anfang) = 0.0
    ' ''                                                End Try

    ' ''                                            Next m

    ' ''                                            crole = New clsRolle(ende - anfang + 1)
    ' ''                                            With crole
    ' ''                                                .RollenTyp = r
    ' ''                                                .Xwerte = Xwerte
    ' ''                                            End With

    ' ''                                            With cphase
    ' ''                                                .addRole(crole)
    ' ''                                            End With
    ' ''                                        Catch ex As Exception
    ' ''                                            '
    ' ''                                            ' handelt es sich um die Kostenart Definition?
    ' ''                                            '
    ' ''                                        End Try

    ' ''                                    ElseIf CostDefinitions.Contains(hname) Then

    ' ''                                        Try

    ' ''                                            k = CInt(CostDefinitions.getCostdef(hname).UID)

    ' ''                                            ReDim Xwerte(ende - anfang)

    ' ''                                            Dim m As Integer
    ' ''                                            For m = anfang To ende
    ' ''                                                Try
    ' ''                                                    Xwerte(m - anfang) = CDbl(zelle.Offset(0, m + 1).Value)
    ' ''                                                Catch ex As Exception
    ' ''                                                    Xwerte(m - anfang) = 0.0
    ' ''                                                End Try

    ' ''                                            Next m

    ' ''                                            ccost = New clsKostenart(ende - anfang + 1)
    ' ''                                            With ccost
    ' ''                                                .KostenTyp = k
    ' ''                                                .Xwerte = Xwerte
    ' ''                                            End With


    ' ''                                            With cphase
    ' ''                                                .AddCost(ccost)
    ' ''                                            End With

    ' ''                                        Catch ex As Exception

    ' ''                                        End Try

    ' ''                                    End If

    ' ''                                Case False  ' es wurde weder Phase noch Rolle angegeben. 
    ' ''                                    If firsttime Then
    ' ''                                        firsttime = False
    ' ''                                    Else 'beim 2. mal: letzte Phase hinzufügen; ENDE von For-Schleife for each Zelle
    ' ''                                        '''''hproj.AddPhase(cphase)

    ' ''                                        Dim hrchynode As New clsHierarchyNode
    ' ''                                        hrchynode.elemName = cphase.name


    ' ''                                        If cphaseLevel = 0 Then
    ' ''                                            hrchynode.parentNodeKey = ""

    ' ''                                        ElseIf cphaseLevel = 1 Then
    ' ''                                            hrchynode.parentNodeKey = rootPhaseName

    ' ''                                        ElseIf cphaseLevel - lastlevel = 1 Then
    ' ''                                            hrchynode.parentNodeKey = lastelemID

    ' ''                                        ElseIf cphaseLevel - lastlevel = 0 Then
    ' ''                                            hrchynode.parentNodeKey = hproj.hierarchy.getParentIDOfID(lastelemID)

    ' ''                                        ElseIf lastlevel - cphaseLevel >= 1 Then
    ' ''                                            Dim hilfselemID As String = lastelemID
    ' ''                                            For l As Integer = 1 To lastlevel - cphaseLevel
    ' ''                                                hilfselemID = hproj.hierarchy.getParentIDOfID(hilfselemID)
    ' ''                                            Next l
    ' ''                                            hrchynode.parentNodeKey = hproj.hierarchy.getParentIDOfID(hilfselemID)
    ' ''                                        Else
    ' ''                                            Throw New ArgumentException("Fehler beim Import! Hierarchie kann nicht richtig aufgebaut werden")
    ' ''                                        End If

    ' ''                                        hproj.AddPhase(cphase, parentID:=hrchynode.parentNodeKey)

    ' ''                                        Exit For
    ' ''                                    End If

    ' ''                            End Select

    ' ''                    End Select

    ' ''                Next zelle


    ' ''            End With
    ' ''        Catch ex As Exception
    ' ''            Throw New ArgumentException("Fehler in awinImportProject, Lesen Ressourcen von '" & hproj.name & "' " & vbLf & ex.Message)
    ' ''        End Try

    ' ''    End If

    ' ''    '' hier wurde jetzt die Reihenfolge geändert - erst werden die Phasen Definitionen eingelesen ..

    ' ''    '' jetzt werden die Daten für die Phasen sowie die Termine/Deliverables eingelesen 

    ' ''    Try
    ' ''        Dim wsTermine As Excel.Worksheet
    ' ''        Try
    ' ''            wsTermine = CType(appInstance.ActiveWorkbook.Worksheets("Termine"), _
    ' ''                                                         Global.Microsoft.Office.Interop.Excel.Worksheet)
    ' ''        Catch ex As Exception
    ' ''            wsTermine = Nothing
    ' ''        End Try

    ' ''        If Not IsNothing(wsTermine) Then
    ' ''            Try
    ' ''                With wsTermine
    ' ''                    Dim lastrow As Integer
    ' ''                    Dim phaseNameID As String
    ' ''                    Dim milestoneName As String
    ' ''                    Dim milestoneDate As Date
    ' ''                    Dim resultVerantwortlich As String = ""
    ' ''                    Dim bewertungsAmpel As Integer
    ' ''                    Dim explanation As String
    ' ''                    Dim deliverables As String
    ' ''                    Dim bewertungsdatum As Date = importDatum
    ' ''                    Dim Nummer As String
    ' ''                    Dim tbl As Excel.Range
    ' ''                    Dim rowOffset As Integer
    ' ''                    Dim columnOffset As Integer


    ' ''                    .Unprotect(Password:="x")       ' Blattschutz aufheben

    ' ''                    tbl = .Range("ErgebnTabelle")
    ' ''                    rowOffset = tbl.Row
    ' ''                    columnOffset = tbl.Column

    ' ''                    lastrow = CInt(CType(.Cells(2000, columnOffset), Excel.Range).End(XlDirection.xlUp).Row)

    ' ''                    ' ur: 12.05.2015: hier wurde die Sortierung der ErgebnTabelle entfernt

    ' ''                    Dim cphase As clsPhase = Nothing
    ' ''                    Dim breadCrumb As String = ""
    ' ''                    Dim lastLevel As Integer = 0

    ' ''                    For zeile = rowOffset To lastrow


    ' ''                        Dim cMilestone As clsMeilenstein
    ' ''                        Dim cBewertung As clsBewertung

    ' ''                        Dim objectName As String
    ' ''                        Dim startDate As Date, endeDate As Date
    ' ''                        ' 
    ' ''                        Dim errMessage As String = ""
    ' ''                        Dim aktLevel As Integer = 0

    ' ''                        Dim isPhase As Boolean = False
    ' ''                        Dim isMeilenstein As Boolean = False
    ' ''                        Dim cphaseExisted As Boolean = True

    ' ''                        '' ''If zeile = 68 Then
    ' ''                        '' ''    zeile = 68
    ' ''                        '' ''End If
    ' ''                        Try
    ' ''                            ' Wenn es keine Phasen gibt in diesem Projekt, so wird trotzdem die Phase1, die ProjektPhase erzeugt.

    ' ''                            If hproj.AllPhases.Count = 0 Then
    ' ''                                Dim duration As Integer
    ' ''                                Dim offset As Integer

    ' ''                                ' Erzeuge ProjektPhase mit Länge des Projekts
    ' ''                                cphase = New clsPhase(parent:=hproj)
    ' ''                                cphase.nameID = rootPhaseName
    ' ''                                'cphaseExisted = False       ' Phase existiert noch nicht

    ' ''                                offset = 0

    ' ''                                If ProjektdauerIndays < 1 Or offset < 0 Then
    ' ''                                    Throw New Exception("unzulässige Angaben für Offset und Dauer: " & _
    ' ''                                                        offset.ToString & ", " & duration.ToString)
    ' ''                                End If

    ' ''                                cphase.changeStartandDauer(offset, ProjektdauerIndays)
    ' ''                                hproj.AddPhase(cphase)

    ' ''                            End If                            'Phase 1 ist nun angelegt


    ' ''                            Try
    ' ''                                ' String aus erster Spalte der Tabelle lesen

    ' ''                                objectName = CType(CType(.Cells(zeile, columnOffset), Excel.Range).Value, String).Trim

    ' ''                                ' Level abfragen

    ' ''                                Dim x As Integer = CInt(CType(.Cells(zeile, columnOffset), Excel.Range).IndentLevel)
    ' ''                                If x Mod einrückTiefe <> 0 Then
    ' ''                                    Throw New ArgumentException("die Einrückung ist keine durch '" & CStr(einrückTiefe) & "' teilbare Zahl")
    ' ''                                End If
    ' ''                                aktLevel = CInt(x / einrückTiefe)


    ' ''                            Catch ex As Exception
    ' ''                                objectName = Nothing
    ' ''                                Throw New Exception("In Tabelle 'Termine' ist der PhasenName nicht angegeben ")
    ' ''                                Exit For ' Ende der For-Schleife, wenn keine laufende Nummer mehr existiert
    ' ''                            End Try


    ' ''                            Try
    ' ''                                startDate = CDate(CType(.Cells(zeile, columnOffset + 2), Excel.Range).Value)
    ' ''                            Catch ex As Exception
    ' ''                                startDate = Date.MinValue
    ' ''                            End Try


    ' ''                            If objectName = elemNameOfElemID(rootPhaseName) Or PhaseDefinitions.Contains(objectName) Then

    ' ''                                isPhase = True
    ' ''                                isMeilenstein = False


    ' ''                            ElseIf startDate <> Date.MinValue Then
    ' ''                                Throw New ArgumentException("'" & objectName & "' ist eine Phase, die nicht im CustomizationFile definiert ist. Bitte korrigieren Sie dies!")
    ' ''                            Else

    ' ''                                isPhase = False
    ' ''                                isMeilenstein = True

    ' ''                            End If


    ' ''                            '  ur: 12.05.2015: Änderung, damit Meilensteine, die den gleichen Namen haben wie Phasen, trotzdem als Meilensteine erkannt werden.
    ' ''                            '                 gilt aktuell aber nur für den BMW-Import
    ' ''                            If awinSettings.importTyp = 2 Then
    ' ''                                If PhaseDefinitions.Contains(objectName) _
    ' ''                                    And startDate = Date.MinValue Then

    ' ''                                    isPhase = False
    ' ''                                    isMeilenstein = True
    ' ''                                End If
    ' ''                            End If

    ' ''                            Try
    ' ''                                endeDate = CDate(CType(.Cells(zeile, columnOffset + 3), Excel.Range).Value)
    ' ''                            Catch ex As Exception
    ' ''                                endeDate = Date.MinValue
    ' ''                            End Try


    ' ''                            If DateDiff(DateInterval.Day, hproj.startDate, startDate) < 0 Then
    ' ''                                ' kein gültiges Startdatum angegeben

    ' ''                                If startDate <> Date.MinValue Then
    ' ''                                    cphase = Nothing
    ' ''                                    Throw New Exception("Die Phase '" & objectName & "' beginnt vor dem Projekt !" & vbLf &
    ' ''                                                 "Bitte korrigieren Sie dies in der Datei'" & hproj.name & ".xlsx'")
    ' ''                                Else
    ' ''                                    ' objectName ist ein Meilenstein

    ' ''                                    'ur: 1.6.2015   Meilenstein hat den Namen einer Phase
    ' ''                                    If PhaseDefinitions.Contains(objectName) _
    ' ''                                        And startDate = Date.MinValue Then

    ' ''                                        isPhase = False
    ' ''                                        isMeilenstein = True
    ' ''                                    End If

    ' ''                                    'ur:12.05.2015:
    ' ''                                    ' '' '' ''If IsNothing(cphase) Then
    ' ''                                    ' '' '' ''    If hproj.AllPhases.Count > 0 Then
    ' ''                                    ' '' '' ''        cphase = hproj.getPhase(1)
    ' ''                                    ' '' '' ''    Else
    ' ''                                    ' '' '' ''        ' Erzeuge ProjektPhase mit Länge des Projekts

    ' ''                                    ' '' '' ''    End If

    ' ''                                    ' '' '' ''End If
    ' ''                                End If


    ' ''                                'isPhase = False

    ' ''                            Else
    ' ''                                'objectName ist eine Phase
    ' ''                                'isPhase = True

    ' ''                                ' ist der Phasen Name in der Liste der definitionen überhaupt bekannt ? 
    ' ''                                If Not PhaseDefinitions.Contains(objectName) Then

    ' ''                                    ' jetzt noch prüfen, ob es sich um die Phase (1) handelt, dann kann sie ja nicht in der PhaseDefinitions enthalten sein  ..
    ' ''                                    If elemNameOfElemID(rootPhaseName) = objectName Or hproj.name = objectName Then
    ' ''                                        ' alles ok
    ' ''                                    Else
    ' ''                                        Throw New Exception("Phase '" & objectName & "' ist nicht definiert!" & vbLf &
    ' ''                                                       "Bitte löschen Sie diese Phase aus '" & hproj.name & "'.xlsx, Tabellenblatt 'Termine'")

    ' ''                                    End If

    ' ''                                End If

    ' ''                                ' an dieser stelle ist sichergestellt, daß der Phasen Name bekannt ist
    ' ''                                ' Prüfen, ob diese Phase bereits in hproj über das ressourcen Sheet angelegt wurde 
    ' ''                                ' tk: dieser Befehl holt jetzt die erste Phase mit deisem NAmen, berücksichtigt aber noch nicht die Position ind er Hierarchie; 
    ' ''                                ' das muss noch ergänzt werden 
    ' ''                                If hproj.name = objectName Or elemNameOfElemID(rootPhaseName) = objectName Then
    ' ''                                    cphase = hproj.getPhaseByID(rootPhaseName)
    ' ''                                    breadCrumb = ""
    ' ''                                Else

    ' ''                                    If aktLevel > lastLevel Then

    ' ''                                        If breadCrumb = "" Then
    ' ''                                            breadCrumb = "."
    ' ''                                        Else
    ' ''                                            breadCrumb = breadCrumb & "#" & cphase.name
    ' ''                                        End If

    ' ''                                    ElseIf aktLevel = lastLevel Then
    ' ''                                        ' aktlevel = lastlevel: also nicht tun
    ' ''                                    Else

    ' ''                                        While aktLevel < lastLevel
    ' ''                                            Dim hstr As String = ""
    ' ''                                            Call splitHryFullnameTo2(breadCrumb, hstr, breadCrumb)
    ' ''                                            lastLevel = lastLevel - 1
    ' ''                                        End While

    ' ''                                    End If
    ' ''                                    cphase = hproj.getPhase(objectName, breadCrumb)

    ' ''                                    If IsNothing(cphase) Then
    ' ''                                        If aktLevel <> hproj.hierarchy.getIndentLevel(cphase.nameID) Then

    ' ''                                            ' ur: 11.05.2015: fehler, wenn die Phase nicht exisitiert, 
    ' ''                                            '               nicht erzeugen
    ' ''                                            ' Phase existiert nicht mit dem gleichen Breadcrumb
    ' ''                                            Throw New ArgumentException("Die Phase '" & objectName & "' existiert nicht in dieser angegebenen Stufe" & vbLf & _
    ' ''                                                                        "Bitte korrigieren Sie die Importdatei!" & "BreadCrumb = " & breadCrumb)

    ' ''                                        End If



    ' ''                                    End If

    ' ''                                End If


    ' ''                            End If

    ' ''                            If isPhase Then  'xxxx Phase
    ' ''                                Try

    ' ''                                    Dim duration As Long
    ' ''                                    Dim offset As Long



    ' ''                                    duration = calcDauerIndays(startDate, endeDate)
    ' ''                                    offset = DateDiff(DateInterval.Day, hproj.startDate, startDate)


    ' ''                                    If duration < 1 Or offset < 0 Then
    ' ''                                        If startDate = Date.MinValue And endeDate = Date.MinValue Then
    ' ''                                            Throw New Exception(" zu '" & objectName & "' wurde kein Datum eingetragen!")
    ' ''                                        Else
    ' ''                                            Throw New Exception("unzulässige Angaben für Offset und Dauer: " & _
    ' ''                                                                offset.ToString & ", " & duration.ToString)
    ' ''                                        End If
    ' ''                                    End If

    ' ''                                    cphase.changeStartandDauer(offset, duration)

    ' ''                                    ' jetzt wird auf Inkonsistenz geprüft 
    ' ''                                    Dim inkonsistent As Boolean = False

    ' ''                                    If cphase.countRoles > 0 Or cphase.countCosts > 0 Then
    ' ''                                        ' prüfen , ob es Inkonsistenzen gibt ? 
    ' ''                                        Dim r As Integer
    ' ''                                        For r = 1 To cphase.countRoles
    ' ''                                            If cphase.getRole(r).Xwerte.Length <> cphase.relEnde - cphase.relStart + 1 Then
    ' ''                                                inkonsistent = True
    ' ''                                            End If
    ' ''                                        Next

    ' ''                                        Dim k As Integer
    ' ''                                        For k = 1 To cphase.countCosts
    ' ''                                            If cphase.getCost(k).Xwerte.Length <> cphase.relEnde - cphase.relStart + 1 Then
    ' ''                                                inkonsistent = True
    ' ''                                            End If
    ' ''                                        Next
    ' ''                                    End If

    ' ''                                    If inkonsistent Then
    ' ''                                        anzFehler = anzFehler + 1
    ' ''                                        Throw New Exception("Der Import konnte nicht fertiggestellt werden. " & vbLf & "Die Dauer der Phase '" & cphase.name & "'  in 'Termine' ist ungleich der in 'Ressourcen' " & vbLf &
    ' ''                                                             "Korrigieren Sie bitte gegebenenfalls diese Inkonsistenz in der Datei '" & vbLf & hproj.name & ".xlsx'")
    ' ''                                    End If
    ' ''                                    ' '' '' ''If Not cphaseExisted Then
    ' ''                                    ' '' '' ''    ' ur: 11.05.2015: parentID bestimmen fehlt hier noch
    ' ''                                    ' '' '' ''    hproj.AddPhase(cphase, parentID:=rootPhaseName)
    ' ''                                    ' '' '' ''End If


    ' ''                                Catch ex As Exception
    ' ''                                    Throw New Exception(ex.Message)
    ' ''                                End Try

    ' ''                            Else

    ' ''                                If aktLevel > lastLevel Then

    ' ''                                    If breadCrumb = "" Then
    ' ''                                        breadCrumb = "."
    ' ''                                    Else
    ' ''                                        breadCrumb = breadCrumb & "#" & cphase.name
    ' ''                                    End If

    ' ''                                ElseIf aktLevel = lastLevel Then
    ' ''                                    ' aktlevel = lastlevel: also nicht tun
    ' ''                                Else

    ' ''                                    While aktLevel < lastLevel
    ' ''                                        Dim hstr As String = ""
    ' ''                                        Call splitHryFullnameTo2(breadCrumb, hstr, breadCrumb)
    ' ''                                        lastLevel = lastLevel - 1
    ' ''                                    End While

    ' ''                                End If

    ' ''                                phaseNameID = cphase.nameID
    ' ''                                cMilestone = New clsMeilenstein(parent:=cphase)
    ' ''                                cBewertung = New clsBewertung

    ' ''                                milestoneName = objectName.Trim
    ' ''                                milestoneDate = endeDate

    ' ''                                ' wenn der freefloat nicht zugelassen ist und der Meilenstein ausserhalb der Phasen-Grenzen liegt 
    ' ''                                ' muss abgebrochen werden 

    ' ''                                If Not awinSettings.milestoneFreeFloat And _
    ' ''                                    (DateDiff(DateInterval.Day, cphase.getStartDate, milestoneDate) < 0 Or _
    ' ''                                     DateDiff(DateInterval.Day, cphase.getEndDate, milestoneDate) > 0) Then
    ' ''                                    Throw New Exception("Der Meilenstein liegt ausserhalb seiner Phase" & vbLf & _
    ' ''                                                        milestoneName & " nicht innerhalb " & cphase.name & vbLf & _
    ' ''                                                             "Korrigieren Sie bitte diese Inkonsistenz in der Datei '" & vbLf & hproj.name & ".xlsx'")
    ' ''                                End If


    ' ''                                ' wenn kein Datum angegeben wurde, soll das Ende der Phase als Datum angenommen werden 
    ' ''                                If DateDiff(DateInterval.Month, hproj.startDate, milestoneDate) < -1 Then
    ' ''                                    milestoneDate = hproj.startDate.AddDays(cphase.startOffsetinDays + cphase.dauerInDays)
    ' ''                                Else
    ' ''                                    If DateDiff(DateInterval.Day, endedateProjekt, endeDate) > 0 Then
    ' ''                                        Call MsgBox("der Meilenstein '" & milestoneName & "' liegt später als das Ende des gesamten Projekts" & vbLf &
    ' ''                                                    "Bitte korrigieren Sie dies im Tabellenblatt Ressourcen der Datei '" & hproj.name & ".xlsx")
    ' ''                                    End If

    ' ''                                End If

    ' ''                                ' resultVerantwortlich = CType(.Cells(zeile, 5).value, String)
    ' ''                                Try
    ' ''                                    bewertungsAmpel = CType(CType(.Cells(zeile, columnOffset + 4), Excel.Range).Value, Integer)
    ' ''                                Catch ex As Exception
    ' ''                                    bewertungsAmpel = 0
    ' ''                                End Try

    ' ''                                explanation = CType(CType(.Cells(zeile, columnOffset + 5), Excel.Range).Value, String)

    ' ''                                ' Ergänzung tk 2.11 deliverables ergänzt 
    ' ''                                deliverables = CType(CType(.Cells(zeile, columnOffset + 6), Excel.Range).Value, String)


    ' ''                                If bewertungsAmpel < 0 Or bewertungsAmpel > 3 Then
    ' ''                                    ' es gibt keine Bewertung
    ' ''                                    bewertungsAmpel = 0
    ' ''                                End If
    ' ''                                ' damit Kriterien auch eingelesen werden, wenn noch keine Bewertung existiert ...
    ' ''                                With cBewertung
    ' ''                                    '.bewerterName = resultVerantwortlich
    ' ''                                    .colorIndex = bewertungsAmpel
    ' ''                                    .datum = importDatum
    ' ''                                    .description = explanation
    ' ''                                    .deliverables = deliverables
    ' ''                                End With



    ' ''                                With cMilestone
    ' ''                                    .setDate = milestoneDate
    ' ''                                    '.verantwortlich = resultVerantwortlich
    ' ''                                    .nameID = hproj.hierarchy.findUniqueElemKey(milestoneName, True)
    ' ''                                    If Not cBewertung Is Nothing Then
    ' ''                                        .addBewertung(cBewertung)
    ' ''                                    End If
    ' ''                                End With


    ' ''                                Try
    ' ''                                    With hproj.getPhaseByID(phaseNameID)
    ' ''                                        .addMilestone(cMilestone)
    ' ''                                    End With
    ' ''                                Catch ex1 As Exception
    ' ''                                    Throw New Exception(ex1.Message)
    ' ''                                End Try



    ' ''                            End If

    ' ''                        Catch ex As Exception
    ' ''                            If zeile <> lastrow Then
    ' ''                                ' beim lesen des ImportFiles ist ein Fehler aufgetreten
    ' ''                                Throw New Exception(ex.Message)
    ' ''                            End If
    ' ''                            ' letzte belegte Zeile wurde bereits bearbeitet.
    ' ''                            zeile = lastrow + 1 ' erzwingt das Ende der For - Schleife
    ' ''                            Nummer = Nothing


    ' ''                        End Try

    ' ''                        lastLevel = aktLevel                ' indentlevel merken
    ' ''                    Next

    ' ''                End With
    ' ''            Catch ex As Exception
    ' ''                Throw New Exception(ex.Message)
    ' ''            End Try

    ' ''        End If
    ' ''        If anzFehler > 0 Then
    ' ''            Call MsgBox("Anzahl Fehler bei Import der Termine von " & hproj.name & " : " & anzFehler)
    ' ''        End If

    ' ''    Catch ex As Exception
    ' ''        Throw New Exception(ex.Message)
    ' ''    End Try

    ' ''    If isTemplate Then
    ' ''        ' hier müssen die Werte für die Vorlage übergeben werden.
    ' ''        Dim projVorlage As New clsProjektvorlage
    ' ''        projVorlage.VorlagenName = hproj.name
    ' ''        projVorlage.Schrift = hproj.Schrift
    ' ''        projVorlage.Schriftfarbe = hproj.Schriftfarbe
    ' ''        projVorlage.farbe = hproj.farbe
    ' ''        projVorlage.earliestStart = -6
    ' ''        projVorlage.latestStart = 6
    ' ''        projVorlage.AllPhases = hproj.AllPhases
    ' ''        projVorlage.hierarchy = hproj.hierarchy
    ' ''        hprojTemp = projVorlage

    ' ''    Else
    ' ''        hprojekt = hproj
    ' ''    End If

    ' ''End Sub

    ''' <summary>
    ''' lädt die jeweils letzten PName/Variante Projekte aus MongoDB in alleProjekte
    ''' lädt ausserdem alle definierten Konstellationen
    ''' zeigt dann die letzte (last) an 
    ''' </summary>
    ''' <remarks></remarks>
    Sub awinletzteKonstellationLaden(ByVal databaseName As String)

        'Dim allProjectsList As SortedList(Of String, clsProjekt)
        Dim zeitraumVon As Date = StartofCalendar.AddMonths(showRangeLeft - 1)
        Dim zeitraumbis As Date = StartofCalendar.AddMonths(showRangeRight - 1)
        Dim storedHeute As Date = Now
        Dim storedGestern As Date = storedHeute.AddDays(-1)
        Dim pname As String = ""
        Dim variantName As String = ""
        Dim request As New Request(awinSettings.databaseURL, databaseName, dbUsername, dbPasswort)
        Dim lastConstellation As New clsConstellation
        Dim hproj As clsProjekt

        If request.pingMongoDb() Then

            projectConstellations = request.retrieveConstellationsFromDB()

            ' Showprojekte leer machen 
            Try
                'NoShowProjekte.Clear()
                ShowProjekte.Clear()
                lastConstellation = projectConstellations.getConstellation(calcLastSessionScenarioName)
            Catch ex As Exception
                'Call MsgBox("in awinProjekteInitialLaden Fehler ...")
            End Try

            ' jetzt Showprojekte aufbauen - und zwar so, dass Konstellation <Last> wiederhergestellt wird
            For Each kvp As KeyValuePair(Of String, clsConstellationItem) In lastConstellation.Liste

                Try
                    hproj = AlleProjekte.getProject(kvp.Key)
                    hproj.startDate = kvp.Value.start
                    hproj.tfZeile = kvp.Value.zeile
                    If kvp.Value.show Then
                        ' nur dann 
                        ShowProjekte.Add(hproj)
                    End If

                Catch ex As Exception
                    Call MsgBox("in ProjekteInitialLaden: " & ex.Message)
                End Try
            Next

        Else
            Call MsgBox("Datenbank-Verbindung ist unterbrochen !")
        End If

    End Sub

    ''' <summary>
    ''' lädt die Projekte im definierten Zeitraum (nach)
    ''' </summary>
    ''' <param name="databaseName"></param>
    ''' <remarks></remarks>
    Sub awinProjekteImZeitraumLaden(ByVal databaseName As String, ByVal filter As clsFilter)

        Dim zeitraumVon As Date = StartofCalendar.AddMonths(showRangeLeft - 1)
        Dim zeitraumbis As Date = StartofCalendar.AddMonths(showRangeRight - 1)
        Dim storedHeute As Date = Now
        Dim storedGestern As Date = storedHeute.AddDays(-1)
        Dim pname As String = ""
        Dim variantName As String = ""
        Dim request As New Request(awinSettings.databaseURL, databaseName, dbUsername, dbPasswort)
        Dim lastConstellation As New clsConstellation
        Dim projekteImZeitraum As New SortedList(Of String, clsProjekt)
        Dim projektHistorie As New clsProjektHistorie


        Dim ok As Boolean = True
        Dim filterIsActive As Boolean
        Dim toShowListe As New SortedList(Of Double, String)


        ' wurde ein definierter Filter mit übergeben ?
        If IsNothing(filter) Then
            filterIsActive = False
        Else
            If filter.isEmpty Then
                filterIsActive = False
            Else
                filterIsActive = True
            End If
        End If

        If request.pingMongoDb() Then

            projekteImZeitraum = request.retrieveProjectsFromDB(pname, variantName, zeitraumVon, zeitraumbis, storedGestern, storedHeute, True)
        Else
            Call MsgBox("Datenbank-Verbindung ist unterbrochen")
        End If

        If AlleProjekte.Count > 0 Then
            ' es sind bereits Projekte geladen 
            Dim atleastOne As Boolean = False

            For Each kvp As KeyValuePair(Of String, clsProjekt) In projekteImZeitraum

                If filterIsActive Then
                    ok = filter.doesNotBlock(kvp.Value)
                Else
                    ok = True
                End If

                If ok Then
                    ' Ist das Projekt bereits in AlleProjekte ? 
                    If AlleProjekte.Containskey(kvp.Key) Then
                        ' das Projekt soll nicht überschrieben werden ...
                        ' also nichts tun 
                    Else
                        ' Workaround: 
                        Dim tmpValue As Integer = kvp.Value.dauerInDays
                        ' tk, Änderung 19.1.17 nicht mehr notwendig ..
                        ' Call awinCreateBudgetWerte(kvp.Value)

                        AlleProjekte.Add(kvp.Value)
                        If ShowProjekte.contains(kvp.Value.name) Then
                            ' auch hier ist nichts zu tun, dann ist bereits eine andere Variante aktiv ...
                        Else
                            ShowProjekte.Add(kvp.Value)
                            atleastOne = True
                        End If
                    End If

                End If

            Next

            ' jetzt ist Showprojekte und AlleProjekte aufgebaut ... 
            ' jetzt muss ClearPlanTafel kommen 
            If atleastOne Then
                Call awinClearPlanTafel()
                Call awinZeichnePlanTafel(True)
            End If

        Else

            ShowProjekte.Clear()
            ' ShowProjekte aufbauen

            For Each kvp As KeyValuePair(Of String, clsProjekt) In projekteImZeitraum

                If filterIsActive Then
                    ok = filter.doesNotBlock(kvp.Value)
                Else
                    ok = True
                End If

                If ok Then

                    Dim tmpValue As Integer = kvp.Value.dauerInDays
                    ' tk, Änderung 19.1.17 nicht mehr notwendig ..
                    ' Call awinCreateBudgetWerte(kvp.Value)
                    AlleProjekte.Add(kvp.Value)

                    Try
                        ' bei Vorhandensein von mehreren Varianten, immer die Standard Variante laden
                        If ShowProjekte.contains(kvp.Value.name) Then
                            If kvp.Value.variantName = "" Then
                                ShowProjekte.Remove(kvp.Value.name)
                                ShowProjekte.Add(kvp.Value)
                            End If
                        Else
                            ShowProjekte.Add(kvp.Value)
                        End If

                    Catch ex As Exception
                        Call MsgBox(ex.Message)
                    End Try
                End If

            Next

            Call awinZeichnePlanTafel(True)

        End If


    End Sub
    ' wurde ersetzt durch addConstellation
    '' ''' <summary>
    '' ''' lädt ein bestimmtes Portfolio von der Datenbank und zeigt es  
    '' ''' in der Projekttafel an.
    '' ''' 
    '' ''' </summary>
    '' ''' <param name="activeConstellation">
    '' ''' Konstellation, die geladen werden soll  
    '' ''' </param>
    '' ''' <remarks></remarks>
    '' ''' 
    ''Public Sub loadConstellation(ByVal activeConstellation As clsConstellation, ByVal storedAtOrBefore As Date)

    ''    Dim hproj As New clsProjekt
    ''    Dim nvErrorMessage As String = ""
    ''    Dim neErrorMessage As String = " (Datum kann nicht angepasst werden)"
    ''    Dim outPutCollection = New Collection
    ''    Dim outputLine As String = ""


    ''    ' prüfen, ob diese Constellation bereits existiert ..
    ''    If IsNothing(activeConstellation) Then
    ''        Call MsgBox(" das Szenario darf nicht NULL sein ... ")
    ''        Exit Sub
    ''    End If

    ''    ShowProjekte.Clear()

    ''    ' jetzt werden die Start-Values entsprechend gesetzt ..

    ''    For Each kvp As KeyValuePair(Of String, clsConstellationItem) In activeConstellation.Liste

    ''        If AlleProjekte.Containskey(kvp.Key) Then
    ''            ' Projekt ist bereits im Hauptspeicher geladen
    ''            hproj = AlleProjekte.getProject(kvp.Key)

    ''        ElseIf Not noDB Then

    ''            Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
    ''            If request.pingMongoDb() Then

    ''                If request.projectNameAlreadyExists(kvp.Value.projectName, kvp.Value.variantName, storedAtOrBefore) Then

    ''                    ' Projekt ist noch nicht im Hauptspeicher geladen, es muss aus der Datenbank geholt werden.
    ''                    hproj = request.retrieveOneProjectfromDB(kvp.Value.projectName, kvp.Value.variantName, storedAtOrBefore)
    ''                    If Not IsNothing(hproj) Then
    ''                        ' Projekt muss nun in die Liste der geladenen Projekte eingetragen werden
    ''                        AlleProjekte.Add(hproj)
    ''                    Else
    ''                        outputLine = kvp.Value.projectName & "(" & kvp.Value.variantName & ") Code: 098 " & nvErrorMessage
    ''                        outPutCollection.Add(outputLine)
    ''                    End If

    ''                Else

    ''                    hproj = Nothing
    ''                    outputLine = kvp.Value.projectName & "(" & kvp.Value.variantName & ")" & nvErrorMessage
    ''                    outPutCollection.Add(outputLine)
    ''                End If
    ''            Else
    ''                Throw New ArgumentException("Datenbank-Verbindung ist unterbrochen!" & vbLf & "Projekt '" & kvp.Value.projectName & "'konnte nicht geladen werden")
    ''            End If

    ''            ''Else      ' not noDB
    ''            ''    Throw New ArgumentException("Projekt '" & kvp.Value.projectName & "'konnte nicht geladen werden")

    ''        End If

    ''        If Not IsNothing(hproj) Then
    ''            If hproj.name = kvp.Value.projectName Then

    ''                With hproj

    ''                    .tfZeile = kvp.Value.zeile

    ''                End With

    ''                If kvp.Value.show Then

    ''                    Try

    ''                        ShowProjekte.Add(hproj)

    ''                    Catch ex1 As Exception
    ''                        outputLine = hproj.name & "(" & hproj.variantName & ")" & " (konnte der Session nicht hinzugefügt werden)"
    ''                        outPutCollection.Add(outputLine)
    ''                    End Try

    ''                    ' jetzt zeichnen des Projektes 
    ''                    ' neu zeichnen des Projekts 
    ''                    Dim tmpCollection As New Collection
    ''                    Call ZeichneProjektinPlanTafel(tmpCollection, hproj.name, hproj.tfZeile, tmpCollection, tmpCollection)

    ''                End If


    ''            End If
    ''        End If


    ''    Next

    ''    ' die aktuelle Konstellation in "Last" speichern 
    ''    'Call storeSessionConstellation("Last")

    ''    If outPutCollection.Count > 0 Then
    ''        Call showOutPut(outPutCollection, _
    ''                        "Meldungen", _
    ''                        "zum Zeitpunkt " & storedAtOrBefore.ToString & " nicht in DB vorhanden:")
    ''    End If

    ''End Sub

    ''' <summary>
    ''' fügt die in der Konstellation aufgeführten Projekte hinzu; 
    ''' wenn Sie bereits geladen sind, wird nachgesehen, ob die richtige Variante aktiviert ist 
    ''' ggf. wird diese Variante dann aktiviert 
    ''' </summary>
    ''' <param name="activeConstellation"></param>
    ''' <remarks></remarks>
    Public Sub addConstellation(ByVal activeConstellation As clsConstellation, ByVal storedAtOrBefore As Date)

        Dim hproj As New clsProjekt
        Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
        Dim nvErrorMessage As String = ""
        Dim neErrorMessage As String = " (Datum kann nicht angepasst werden)"
        Dim outPutCollection = New Collection
        Dim outputLine As String = ""
        Dim tryZeile As Integer

        Dim boardwasEmpty As Boolean = (ShowProjekte.Count = 0)
        ' ab diesem Wert soll neu gezeichnet werden 
        Dim startOfFreeRows As Integer = projectboardShapes.getMaxZeile
        Dim zeilenOffset As Integer = 0

        ' prüfen, ob diese Constellation auch existiert ..
        If IsNothing(activeConstellation) Then
            Call MsgBox(" das Portfolio darf nicht NULL sein ... ")
            Exit Sub
        End If

        ' jetzt muss das Sort-Kriterium übernommen werden 
        If boardwasEmpty And activeConstellation.sortCriteria >= 0 Then
            currentSessionConstellation.sortCriteria = activeConstellation.sortCriteria
        End If


        ' jetzt werden die einzelnen Projekte dazugeholt 

        For Each kvp As KeyValuePair(Of String, clsConstellationItem) In activeConstellation.Liste

           

            Dim showIT As Boolean = kvp.Value.show

            If AlleProjekte.Containskey(kvp.Key) Then
                ' Projekt ist bereits im Hauptspeicher geladen
                hproj = AlleProjekte.getProject(kvp.Key)

                ' ist es aber auch der richtige TimeStamp ? 
                If hproj.timeStamp > storedAtOrBefore Then
                    ' in einer Session dürfen keine TimeStamps aktuellen bzw. früheren TimeStamps gemischt werden ... 
                    ' Meldung in der , und der Nutzer muss alles neu laden 

                    outputLine = "es gibt Projekte mit jüngerem TimeStamp in der Session ... "
                    outPutCollection.Add(outputLine)
                    outputLine = "die Aktion wurde abgebrochen ... "
                    outPutCollection.Add(outputLine)
                    outputLine = "bitte löschen Sie die Session und laden Sie dann die Szenarien mit dem gewünschten Versions-Datum"
                    outPutCollection.Add(outputLine)

                    Exit For
                Else
                    If showIT Then

                        If ShowProjekte.contains(hproj.name) Then
                            ' dann soll das Projekt da bleiben, wo es ist 
                            Dim shownProject As clsProjekt = ShowProjekte.getProject(hproj.name)
                            If shownProject.variantName = hproj.variantName Then
                                ' es wird bereits gezeigt, nichts machen ...
                            Else
                                tryZeile = shownProject.tfZeile
                                ' jetzt die Variante aktivieren 
                                Call replaceProjectVariant(hproj.name, hproj.variantName, False, True, tryZeile)
                            End If

                        ElseIf boardwasEmpty Then
                            'tryZeile = kvp.Value.zeile
                            tryZeile = activeConstellation.getBoardZeile(hproj.name)
                            Call replaceProjectVariant(hproj.name, hproj.variantName, False, True, tryZeile)
                        Else

                            'tryZeile = kvp.Value.zeile + startOfFreeRows - 1
                            'tryZeile = startOfFreeRows + zeilenOffset
                            tryZeile = startOfFreeRows + activeConstellation.getBoardZeile(hproj.name) - 2
                            Call replaceProjectVariant(hproj.name, hproj.variantName, False, True, tryZeile)
                            'zeilenOffset = zeilenOffset + 1
                        End If


                    Else
                        ' gar nichts machen
                    End If

                End If




            Else
                If request.pingMongoDb() Then

                    If request.projectNameAlreadyExists(kvp.Value.projectName, kvp.Value.variantName, storedAtOrBefore) Then

                        ' Projekt ist noch nicht im Hauptspeicher geladen, es muss aus der Datenbank geholt werden.
                        hproj = request.retrieveOneProjectfromDB(kvp.Value.projectName, kvp.Value.variantName, storedAtOrBefore)

                        If Not IsNothing(hproj) Then
                            ' Projekt muss nun in die Liste der geladenen Projekte eingetragen werden
                            Dim newPosition As Integer = -1
                            If currentSessionConstellation.sortCriteria = ptSortCriteria.customTF Then
                                If boardwasEmpty Then
                                    ' den gleichen key verwenden wie in der activeConstellation
                                    newPosition = activeConstellation.getBoardZeile(hproj.name)
                                Else
                                    newPosition = activeConstellation.getBoardZeile(hproj.name) + startOfFreeRows
                                End If
                            End If
                            AlleProjekte.Add(hproj, True, newPosition)
                            ' jetzt die Variante aktivieren 
                            ' aber nur wenn es auch das Flag show hat 
                            If showIT Then

                                If boardwasEmpty Then
                                    'tryZeile = kvp.Value.zeile
                                    tryZeile = activeConstellation.getBoardZeile(hproj.name)
                                    Call replaceProjectVariant(hproj.name, hproj.variantName, False, True, tryZeile)
                                Else
                                    'tryZeile = startOfFreeRows + zeilenOffset
                                    tryZeile = startOfFreeRows + activeConstellation.getBoardZeile(hproj.name) - 2
                                    Call replaceProjectVariant(hproj.name, hproj.variantName, False, False, tryZeile)
                                    'zeilenOffset = zeilenOffset + 1
                                End If

                            End If
                        Else
                            outputLine = kvp.Value.projectName & "(" & kvp.Value.variantName & ") Code: 098 " & nvErrorMessage
                            outPutCollection.Add(outputLine)
                        End If

                    Else
                        hproj = Nothing

                        outputLine = kvp.Value.projectName & "(" & kvp.Value.variantName & ")" & nvErrorMessage
                        outPutCollection.Add(outputLine)

                        'Call MsgBox("Projekt '" & kvp.Value.projectName & "'konnte nicht geladen werden")
                        'Throw New ArgumentException("Projekt '" & kvp.Value.projectName & "'konnte nicht geladen werden")
                    End If
                Else
                    Throw New ArgumentException("Datenbank-Verbindung ist unterbrochen!" & vbLf & "Projekt '" & kvp.Value.projectName & "'konnte nicht geladen werden")
                End If
            End If

        Next


        If outPutCollection.Count > 0 Then

            If outPutCollection.Count > 0 Then
                Call showOutPut(outPutCollection, _
                                "Meldungen", _
                                "zum Zeitpunkt " & storedAtOrBefore.ToString & " nicht in DB vorhanden:")
            End If

        End If


    End Sub

    ''' <summary>
    ''' zeigt die Konstellation bzw Konstellationen auf der Projekt-Tafel an 
    ''' addToSession gibt an, ob AlleProjekte und ggf ShowProjekte ergänzt wird 
    ''' </summary>
    ''' <param name="constellationsToShow"></param>
    ''' <param name="clearBoard">setzt ShowProjekte zurück, löscht das Zeichenbrett; lässt AlleProjekte unverändert </param>
    ''' <param name="clearSession">setzt alles zurück></param>
    ''' <param name="storedAtOrBefore"></param>
    ''' <remarks></remarks>
    Public Sub showConstellations(ByVal constellationsToShow As clsConstellations, ByVal clearBoard As Boolean, ByVal clearSession As Boolean, ByVal storedAtOrBefore As Date)

        Try
            Dim boardWasEmpty As Boolean = (ShowProjekte.Count = 0)
            Dim sessionWasEmpty As Boolean = (AlleProjekte.Count = 0)


            If clearSession And Not sessionWasEmpty Then
                Call clearCompleteSession()

            ElseIf clearBoard And Not boardWasEmpty Then
                Call clearProjectBoard()

            End If

            Dim i As Integer = 0
            For Each kvp As KeyValuePair(Of String, clsConstellation) In constellationsToShow.Liste

                Dim activeConstellation As clsConstellation = kvp.Value

                ' jetzt den Sortier-Modus anpassen 
                If activeConstellation.sortCriteria <> currentSessionConstellation.sortCriteria Then

                    If activeConstellation.sortCriteria >= 0 Then
                        currentSessionConstellation.sortCriteria = activeConstellation.sortCriteria
                        '' ''Else
                        '' ''    currentSessionConstellation.sortCriteria = ptSortCriteria.customTF
                        '' ''    activeConstellation.sortCriteria = ptSortCriteria.customTF
                    End If
                    '' ''Else
                    '' ''    If activeConstellation.sortCriteria < 0 Then
                    '' ''        currentSessionConstellation.sortCriteria = ptSortCriteria.customTF
                    '' ''        activeConstellation.sortCriteria = ptSortCriteria.customTF
                    '' ''    End If

                End If

                Call addConstellation(activeConstellation, storedAtOrBefore)
                ' das Folgende ist unnötig, ggf wuden ja bereits die nötigen Resets gemacht ... 
                ''If i = 0 And (boardWasEmpty Or Not addToSession) Then
                ''    Call loadConstellation(activeConstellation, storedAtOrBefore)
                ''Else
                ''    Call addConstellation(activeConstellation, storedAtOrBefore)
                ''End If

                i = i + 1

            Next

            If constellationsToShow.Count = 1 Then
                If clearSession Or sessionWasEmpty Or _
                    clearBoard Or boardWasEmpty Then
                    currentConstellationName = constellationsToShow.Liste.ElementAt(0).Value.constellationName
                Else
                    currentConstellationName = calcLastSessionScenarioName()
                    ' hier muss jetzt der sortType auf CustomTF gesetzt werden 

                    If Not IsNothing(currentSessionConstellation) Then
                        currentSessionConstellation.sortCriteria = ptSortCriteria.customTF
                    End If

                End If
            Else
                currentConstellationName = calcLastSessionScenarioName()
                ' hier muss jetzt der sortType auf CustomTF gesetzt werden  
                If Not IsNothing(currentSessionConstellation) Then
                    currentSessionConstellation.sortCriteria = ptSortCriteria.customTF
                End If
            End If

            Call awinNeuZeichnenDiagramme(2)

            ' die aktuelle Konstellation in "Last" speichern 
            'Call storeSessionConstellation("Last")

        Catch ex As Exception
            Call MsgBox("Fehler bei Laden : " & vbLf & ex.Message)
        End Try


    End Sub

    ''' <summary>
    ''' speichert eine einzelne Konstellation in die Datenbank
    ''' dabei werden alle Projekte und Projekt-Varianten, die noch nicht oder in anderer Form in der Datenbank gespeichert sind, abgespeichert 
    ''' </summary>
    ''' <param name="currentConstellation"></param>
    ''' <remarks></remarks>
    Public Sub storeSingleConstellationToDB(ByRef outPutCollection As Collection, _
                                            ByVal currentConstellation As clsConstellation)
        Dim anzahlNeue As Integer = 0
        Dim anzahlChanged As Integer = 0
        Dim DBtimeStamp As Date = Date.Now
        Dim outputLine As String = ""


        Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)

        ' jetzt müssen auch alle Projekte, die in der Constellation referenziert werden, aber noch nicht 
        ' in der Datenbank gespeichert sind, abgespeichert werden ... 
        For Each kvp As KeyValuePair(Of String, clsConstellationItem) In currentConstellation.Liste

            Dim hproj As clsProjekt = AlleProjekte.getProject(kvp.Key)

            If Not IsNothing(hproj) Then
                If Not request.projectNameAlreadyExists(hproj.name, hproj.variantName, Date.Now) Then
                    ' speichern des Projektes 
                    hproj.timeStamp = DBtimeStamp
                    If request.storeProjectToDB(hproj, dbUsername) Then

                        If awinSettings.englishLanguage Then
                            outputLine = "stored: " & hproj.name & ", " & hproj.variantName
                            outPutCollection.Add(outputLine)
                        Else
                            outputLine = "gespeichert: " & hproj.name & ", " & hproj.variantName
                            outPutCollection.Add(outputLine)
                        End If

                        anzahlNeue = anzahlNeue + 1

                        Dim wpItem As clsWriteProtectionItem = request.getWriteProtection(hproj.name, hproj.variantName)
                        writeProtections.upsert(wpItem)

                    Else
                        ' kann eigentlich gar nicht sein ... wäre nur dann der Fall, wenn ein Projekt komplett gelöscht wurde , aber der Schreibschutz nicht gelöscht wurde 
                        If awinSettings.englishLanguage Then
                            outputLine = "protected project: " & hproj.name & ", " & hproj.variantName
                        Else
                            outputLine = "geschütztes Projekt: " & hproj.name & ", " & hproj.variantName
                        End If
                        outPutCollection.Add(outputLine)

                        Dim wpItem As clsWriteProtectionItem = request.getWriteProtection(hproj.name, hproj.variantName)
                        writeProtections.upsert(wpItem)

                    End If
                Else
                    ' ein in dem Szenario enthaltenes Projekt wird gespeichert , wenn es Unterschiede gibt 
                    Dim oldProj As clsProjekt = request.retrieveOneProjectfromDB(hproj.name, hproj.variantName, Date.Now)
                    ' Type = 0: Projekt wird mit Variante bzw. anderem zeitlichen Stand verglichen ...
                    If Not hproj.isIdenticalTo(oldProj) Then
                        hproj.timeStamp = DBtimeStamp
                        If request.storeProjectToDB(hproj, dbUsername) Then

                            If awinSettings.englishLanguage Then
                                outputLine = "stored: " & hproj.name & ", " & hproj.variantName
                                outPutCollection.Add(outputLine)
                            Else
                                outputLine = "gespeichert: " & hproj.name & ", " & hproj.variantName
                                outPutCollection.Add(outputLine)
                            End If

                            ' alles ok
                            anzahlChanged = anzahlChanged + 1

                            Dim wpItem As clsWriteProtectionItem = request.getWriteProtection(hproj.name, hproj.variantName)
                            writeProtections.upsert(wpItem)
                        Else
                            If awinSettings.englishLanguage Then
                                outputLine = "protected project: " & hproj.name & ", " & hproj.variantName
                            Else
                                outputLine = "geschütztes Projekt: " & hproj.name & ", " & hproj.variantName
                            End If
                            outPutCollection.Add(outputLine)

                            Dim wpItem As clsWriteProtectionItem = request.getWriteProtection(hproj.name, hproj.variantName)
                            writeProtections.upsert(wpItem)

                        End If
                    End If
                End If
            End If


        Next

        ' jetzt wird die 
        Try
            If request.storeConstellationToDB(currentConstellation) Then
                
            Else
                If awinSettings.englishLanguage Then
                    outputLine = "Error when writing scenario: " & currentConstellation.constellationName
                Else
                    outputLine = "Fehler beim Schreiben Szenario: " & currentConstellation.constellationName
                End If
                outPutCollection.Add(outputLine)

            End If
        Catch ex As Exception
            If awinSettings.englishLanguage Then
                outputLine = "Error when writing scenario - Database active?"
            Else
                outputLine = "Fehler beim Schreiben Szenario - Datenbank läuft?"
            End If
            Throw New ArgumentException(outputLine)
        End Try


        Dim tsMessage As String = ""
        If anzahlNeue + anzahlChanged > 0 Then
            If awinSettings.englishLanguage Then
                tsMessage = "Zeitstempel: " & DBtimeStamp.ToShortDateString & ", " & DBtimeStamp.ToShortTimeString
            Else
                tsMessage = "Timestamp: " & DBtimeStamp.ToShortDateString & ", " & DBtimeStamp.ToShortTimeString
            End If

        End If

        If awinSettings.englishLanguage Then
            outputLine = "Stored ... " & vbLf & _
                "Portfolio: " & currentConstellation.constellationName & vbLf & vbLf & _
                "Number new projects/project-variants: " & anzahlNeue.ToString & vbLf & _
                "Number changed projects/project-variants: " & anzahlChanged.ToString & vbLf & _
                tsMessage
        Else
            outputLine = "Gespeichert ... " & vbLf & _
                "Portfolio: " & currentConstellation.constellationName & vbLf & vbLf & _
                "Anzahl neue Projekte und Projekt-Varianten: " & anzahlNeue.ToString & vbLf & _
                "Anzahl geänderte Projekte / Projekt-Varianten: " & anzahlChanged.ToString & vbLf & _
                tsMessage
        End If
        outPutCollection.Add(outputLine)



    End Sub

    ''' <summary>
    ''' löscht ein bestimmtes Portfolio aus der Datenbank und der Liste der Portfolios im Hauptspeicher
    ''' 
    ''' </summary>
    ''' <param name="constellationName">
    ''' Name, unter dem das Portfolio in der Datenbank gespeichert wurde 
    ''' </param>
    ''' <remarks></remarks>
    ''' 
    Public Sub awinRemoveConstellation(ByVal constellationName As String, ByVal deleteDB As Boolean)

        Dim returnValue As Boolean = True
        Dim activeConstellation As New clsConstellation


        ' prüfen, ob diese Constellation überhaupt existiert ..
        Try
            activeConstellation = projectConstellations.getConstellation(constellationName)
        Catch ex As Exception
            Call MsgBox(" Projekt-Konstellation " & constellationName & " existiert nicht ")
            Exit Sub
        End Try

        If deleteDB Then
            Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
            If Request.pingMongoDb() Then

                ' Konstellation muss aus der Datenbank gelöscht werden.
                returnValue = request.removeConstellationFromDB(activeConstellation)
                If returnValue = False Then
                    Call MsgBox("Fehler bei Löschen Portfolio : " & activeConstellation.constellationName)
                End If
            Else
                Throw New ArgumentException("Datenbank-Verbindung ist unterbrochen!" & vbLf & "Projekt '" & activeConstellation.constellationName & "'konnte nicht gelöscht werden")
                returnValue = False
            End If
        End If

        If returnValue Then
            Try
                ' Konstellation muss aus der Liste aller Portfolios entfernt werden.
                projectConstellations.Remove(activeConstellation.constellationName)
            Catch ex1 As Exception
                Call MsgBox("Fehler in awinRemoveConstellation aufgetreten: " & ex1.Message)
            End Try
        Else
            Call MsgBox("Es ist ein Fehler beim Löschen es Portfolios aus der Datenbank aufgetreten ")
        End If

    End Sub

    ''' <summary>
    ''' lädt die über pName#vName angegebene Variante aus der Datenbank;
    ''' show = true: es wird in Showprojekte eingetragen; sonst nur in AlleProjekte 
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <param name="vName"></param>
    ''' <remarks></remarks>
    Public Sub loadProjectfromDB(ByRef outputCollection As Collection, _
                                 ByVal pName As String, vName As String, ByVal show As Boolean, _
                                 ByVal storedAtORBefore As Date)

        Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
        Dim hproj As clsProjekt
        Dim key As String = calcProjektKey(pName, vName)

        ' ab diesem Wert soll neu gezeichnet werden 
        Dim freieZeile As Integer = projectboardShapes.getMaxZeile

        hproj = request.retrieveOneProjectfromDB(pName, vName, storedAtORBefore)


        If Not IsNothing(hproj) Then
            ' prüfen, ob AlleProjekte das Projekt bereits enthält 
            ' danach ist sichergestellt, daß AlleProjekte das Projekt bereit enthält 
            If AlleProjekte.Containskey(key) Then
                AlleProjekte.Remove(key)
            End If

            AlleProjekte.Add(hproj)

            ' jetzt die writeProtections aktualisieren 
            Dim wpItem As clsWriteProtectionItem = request.getWriteProtection(hproj.name, hproj.variantName)
            writeProtections.upsert(wpItem)

            If show Then
                ' prüfen, ob es bereits in der Showprojekt enthalten ist
                ' diese Prüfung und die entsprechenden Aktionen erfolgen im 
                ' replaceProjectVariant

                Call replaceProjectVariant(pName, vName, False, True, freieZeile)

            End If
        Else
            Dim outputLine As String = "existiert nicht: " & pName & ", " & vName & " @ " & storedAtORBefore.ToString
            outputCollection.Add(outputLine)
        End If


    End Sub

    ''' <summary>
    ''' löscht in der Datenbank alle Timestamps der Projekt-Variante pname, variantname
    ''' die Timestamps werden zudem alle im Papierkorb gesichert 
    ''' </summary>
    ''' <param name="pname">Projektname</param>
    ''' <param name="variantName">Variantenname</param>
    ''' <remarks></remarks>
    Public Sub deleteCompleteProjectVariant(ByRef outputCollection As Collection, _
                                            ByVal pname As String, ByVal variantName As String, ByVal kennung As Integer, _
                                            Optional ByVal keepAnzVersions As Integer = 100)

        Dim outputLine As String = ""

        Dim anzTests As Integer = 0
        Dim anzDeleted As Integer = 0
        If kennung = PTTvActions.delFromDB Or _
            kennung = PTTvActions.delAllExceptFromDB Then

            Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
            'Dim requestTrash As New Request(awinSettings.databaseURL, awinSettings.databaseName & "Trash", dbUsername, dbPasswort)

            If kennung = PTTvActions.delAllExceptFromDB Then

                ' an dieser Stelle wird gecheckt
                ' 1. ist es eine echte Variante und hat sie keine customFields? 
                ' 2. wenn ja, dann hole die Basis-Variante , hat sie CustomFields
                ' 3. wenn ja, dann kopiere die Custom-Fields und speichere die Variante 
                ' mach dann den den Rest 
                ' Start Sonderbehandlung 
                If variantName <> "" Then
                    Dim anzCorrected As Integer = 0
                    Dim variantProject As clsProjekt
                    Dim baseProject As clsProjekt
                    Dim vExisted As Boolean = False
                    Dim bExisted As Boolean = False
                    Dim oCollection As New Collection
                    Dim keyV As String = calcProjektKey(pname, variantName)
                    Dim keyB As String = calcProjektKey(pname, "")

                    If Not AlleProjekte.Containskey(keyV) Then
                        Call loadProjectfromDB(oCollection, pname, variantName, False, Date.Now)
                    Else
                        vExisted = True
                    End If

                    If Not AlleProjekte.Containskey(keyB) Then
                        Call loadProjectfromDB(oCollection, pname, "", False, Date.Now)
                    Else
                        bExisted = True
                    End If

                    variantProject = AlleProjekte.getProject(keyV)
                    baseProject = AlleProjekte.getProject(keyB)
                    '
                    ' Sonderbehandlung alter Fehler bei Variantenbildung: Custom-Fields wurde nicht aus Base-Variant übernommen 
                    If Not IsNothing(variantProject) And Not IsNothing(baseProject) Then

                        ' Sonderbehandlung wegen ehemaligem Fehler, wo bei Varianten-Bildung die Custom-fields aus der Base-Variant nicht übernommen wurden 
                        If variantProject.getCustomFieldsCount = 0 And baseProject.getCustomFieldsCount > 0 Then
                            variantProject.copyCustomFieldsFrom(baseProject)
                            Dim zeitStempel As Date = variantProject.timeStamp

                            ' jetzt löschen, dann speichern ; wenn das löschen schiefgeht aufgrund Schreibschutz, dann geht auch das Speichern schief ... 
                            If writeProtections.isProtected(keyV, dbUsername) Then
                                ' kann nichts machen ...
                            Else
                                If request.deleteProjectTimestampFromDB(pname, variantName, zeitStempel, dbUsername) Then
                                    ' all ok 
                                    If request.storeProjectToDB(variantProject, dbUsername) Then
                                        ' alles ok; jetzt  
                                    Else

                                    End If
                                End If
                            End If


                        End If

                    End If

                    If Not vExisted Then
                        If AlleProjekte.Containskey(keyV) Then
                            AlleProjekte.Remove(keyV)
                        End If
                    End If
                    If Not bExisted Then
                        If AlleProjekte.Containskey(keyB) Then
                            AlleProjekte.Remove(keyB)
                        End If
                    End If
                End If

                ' Ende Sonderbehandlung  
                ' 
                '

                Dim timeStampsToDelete As Collection = identifyTimeStampsToDelete(pname, variantName, keepAnzVersions)


                If timeStampsToDelete.Count >= 1 Then

                    For Each singleTimeStamp As Date In timeStampsToDelete


                        If request.deleteProjectTimestampFromDB(pname, variantName, singleTimeStamp, dbUsername) Then
                            ' all ok 
                            anzDeleted = anzDeleted + 1
                        Else
                            If awinSettings.englishLanguage Then
                                outputLine = "-->Error deleting (protected?): " & pname & ", " & variantName & ", " & singleTimeStamp.ToShortDateString
                            Else
                                outputLine = "-->Fehler beim Löschen (geschützt?): " & pname & ", " & variantName & ", " & singleTimeStamp.ToShortDateString
                            End If

                            outputCollection.Add(outputLine)
                        End If

                    Next

                    If awinSettings.englishLanguage Then
                        outputLine = pname & " (" & variantName & "): " & anzDeleted & " timestamps deleted"
                    Else
                        outputLine = pname & " (" & variantName & "): " & anzDeleted & " TimeStamps gelöscht"
                    End If

                    outputCollection.Add(outputLine)

                Else
                    If awinSettings.englishLanguage Then
                        outputLine = outputLine = pname & " (" & variantName & "): 0 timestamps deleted"
                    Else
                        outputLine = pname & " (" & variantName & "): 0 TimeStamps gelöscht"
                    End If

                    outputCollection.Add(outputLine)
                End If


            Else
                ' jetzt alle Timestamps in der Datenbank löschen 

                ' das darf aber nur passieren, wenn das Projekt, die Variante in keinem Szenario mehr referenziert wird ... 
                ' das hier ist eine doppelte Schranke sozusagen - in der Aufruf Schnittstelle wird das auch schon überprüft  
                If notReferencedByAnyPortfolio(pname, variantName) Then
                    Try

                        If Not IsNothing(projekthistorie) Then
                            projekthistorie.clear() ' alte Historie löschen
                        End If

                        projekthistorie.liste = request.retrieveProjectHistoryFromDB _
                                                (projectname:=pname, variantName:=variantName, _
                                                 storedEarliest:=Date.MinValue, storedLatest:=Date.Now.AddDays(1))


                        ' jetzt über alle Elemente der Projekthistorie ..
                        For Each kvp As KeyValuePair(Of Date, clsProjekt) In projekthistorie.liste

                            If request.deleteProjectTimestampFromDB(pname, variantName, kvp.Key, dbUsername) Then
                                ' all ok 
                                anzDeleted = anzDeleted + 1
                            Else
                                If awinSettings.englishLanguage Then
                                    outputLine = "-->Error deleting (protected?): " & pname & ", " & variantName & ", " & kvp.Key.ToShortDateString
                                Else
                                    outputLine = "-->Fehler beim Löschen (geschützt?): " & pname & ", " & variantName & ", " & kvp.Key.ToShortDateString
                                End If

                                outputCollection.Add(outputLine)

                            End If

                        Next

                    Catch ex As Exception

                    End Try

                Else
                    If variantName = "" Then
                        If awinSettings.englishLanguage Then
                            outputLine = "delete denied: " & pname & " - Scenarios: "
                        Else
                            outputLine = "Löschen verweigert:  " & pname & " - Szenarien: "
                        End If

                    Else
                        If awinSettings.englishLanguage Then
                            outputLine = "delete denied: " & pname & " (" & variantName & ") " & " - Scenarios: "
                        Else
                            outputLine = "Löschen verweigert:  " & pname & " (" & variantName & ") " & " - Szenarien: "
                        End If

                    End If
                    outputLine = outputLine & projectConstellations.getSzenarioNamesWith(pname, variantName)
                    outputCollection.Add(outputLine)
                End If


            End If



        ElseIf kennung = PTTvActions.delFromSession Or _
            kennung = PTTvActions.deleteV Then

            ' eine einzelne Variante kann nur gelöscht werden, wenn 
            ' es sich weder um die variantName = "" noch um die aktuell gezeigte Variante handelt 

            Dim hproj As clsProjekt
            Try
                hproj = ShowProjekte.getProject(pname)
            Catch ex As Exception
                hproj = Nothing
            End Try

            If IsNothing(hproj) Then
                Dim key As String = calcProjektKey(pname, variantName)
                AlleProjekte.Remove(key)

            ElseIf hproj.variantName <> variantName Then
                Dim key As String = calcProjektKey(pname, variantName)
                AlleProjekte.Remove(key)

            Else
                ' es wird in Showprojekte und in AlleProjekte gelöscht, ausserdem auch auf der Projekt-Tafel 

                Dim key As String = calcProjektKey(pname, variantName)

                Try

                    ' jetzt muss die bisherige Variante aus Showprojekte rausgenommen werden ..
                    ShowProjekte.Remove(hproj.name)

                    ' die gewählte Variante wird rausgenommen
                    AlleProjekte.Remove(key)

                    Call clearProjektinPlantafel(pname)

                Catch ex As Exception
                    If awinSettings.englishLanguage Then
                        outputLine = "delete denied: " & pname & " (" & variantName & ") " & " - Scenarios: ""Error when deleting: " & pname & " (" & variantName & ")"
                    Else
                        outputLine = "delete denied: " & pname & " (" & variantName & ") " & " - Scenarios: ""Fehler beim Löschen: " & pname & " (" & variantName & ")"
                    End If
                    outputCollection.Add(outputLine)
                End Try


            End If



        End If


    End Sub

    ''' <summary>
    ''' bestimmt die Time-Stamps, die gelöscht werden sollen 
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <param name="vName"></param>
    ''' <param name="keepAnzVersions"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function identifyTimeStampsToDelete(ByVal pName As String, vName As String, Optional ByVal keepAnzVersions As Integer = -1) As Collection

        Dim tsToDelete As New Collection

        Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
        'Dim requestTrash As New Request(awinSettings.databaseURL, awinSettings.databaseName & "Trash", dbUsername, dbPasswort)

        If Not IsNothing(projekthistorie) Then
            projekthistorie.clear() ' alte Historie löschen
        End If

        projekthistorie.liste = request.retrieveProjectHistoryFromDB _
                                (projectname:=pName, variantName:=vName, _
                                 storedEarliest:=Date.MinValue, storedLatest:=Date.Now.AddDays(1))

        If projekthistorie.Count <= keepAnzVersions Or projekthistorie.Count < 2 Then
            ' es muss nix gelöscht werden ... 

        Else
            Dim listToKeep As New SortedList(Of Date, String)
            Dim anzahlTS As Integer = projekthistorie.Count


            ' das erste Projekt merken 
            If Not listToKeep.ContainsKey(projekthistorie.ElementAt(0).timeStamp) Then
                listToKeep.Add(projekthistorie.ElementAt(0).timeStamp, "")
            End If

            ' das letzte Projekt merken 
            If Not listToKeep.ContainsKey(projekthistorie.ElementAt(anzahlTS - 1).timeStamp) Then
                listToKeep.Add(projekthistorie.ElementAt(anzahlTS - 1).timeStamp, "")
            End If



            ' das letzte Projekt merken, das im Vergleich zum ersten verändert ist ... 
            Dim cIX As Integer = anzahlTS - 1
            Dim lastKeptProjekt As clsProjekt = projekthistorie.ElementAt(cIX)



            Dim vIX As Integer = cIX - 1
            Dim vglProjekt As clsProjekt = projekthistorie.ElementAt(cIX)

            If vIX >= 0 Then
                vglProjekt = projekthistorie.ElementAt(vIX)
            End If


            Dim finished As Boolean = (vIX <= 0)
            Dim anzKept As Integer = listToKeep.Count

            Do While Not finished And anzKept < keepAnzVersions

                Do While vglProjekt.isIdenticalTo(lastKeptProjekt) And vIX >= 1
                    vIX = vIX - 1
                    vglProjekt = projekthistorie.ElementAt(vIX)
                Loop

                ' jetzt ist das vglProjekt ungleich dem lastkeptProjekt oder das Ende ist erreicht 
                If vIX <= 0 Then
                    ' end of operation 
                    finished = True
                Else
                    ' falls es Duplikate gibt: das früheste Projekt finden, das identisch zu vglProjekt ist  
                    vIX = vIX - 1
                    Dim memorizeProjekt As clsProjekt = projekthistorie.ElementAt(vIX)

                    Do Until Not memorizeProjekt.isIdenticalTo(vglProjekt) Or vIX = 0
                        vIX = vIX - 1
                        memorizeProjekt = projekthistorie.ElementAt(vIX)
                    Loop

                    If Not memorizeProjekt.isIdenticalTo(vglProjekt) Then
                        ' es wurde ein Unterschied festgestellt 
                        vIX = vIX + 1
                        lastKeptProjekt = projekthistorie.ElementAt(vIX)

                        If Not listToKeep.ContainsKey(projekthistorie.ElementAt(vIX).timeStamp) Then
                            listToKeep.Add(projekthistorie.ElementAt(vIX).timeStamp, "")
                        End If

                        vIX = vIX - 1
                        If vIX = 0 Then
                            finished = True
                        End If

                        vglProjekt = projekthistorie.ElementAt(vIX)
                    Else
                        finished = True
                    End If


                End If

                anzKept = listToKeep.Count
            Loop


            ' jetzt wird die ProjektHistorie um die toKeepVersions erleichtert ...
            Dim errorOccurred As Boolean = False
            For Each kvp As KeyValuePair(Of Date, String) In listToKeep

                Try
                    If projekthistorie.contains(kvp.Key) Then
                        projekthistorie.remove(kvp.Key)
                    Else
                        errorOccurred = True
                    End If
                Catch ex As Exception
                    errorOccurred = True
                End Try


            Next

            If listToKeep.Count < 1 Then
                ' nichts tun ... 
            Else
                ' jetzt wird gelöscht ... 
                ' hier nur bei diesem Projekt weitermachen, wenn kein Fehler aufgetreten ist; das ist sonst zu kritisch 

                If Not errorOccurred Then

                    For Each kvp As KeyValuePair(Of Date, clsProjekt) In projekthistorie.liste

                        If Not tsToDelete.Contains(kvp.Key) Then
                            tsToDelete.Add(kvp.Key, kvp.Key)
                        End If

                    Next

                End If
            End If

        End If

        identifyTimeStampsToDelete = tsToDelete

    End Function

    ''' <summary>
    ''' gibt true zurück, wenn diese Projekt-Variante in keinem Portfolio enthalten ist ... 
    ''' </summary>
    ''' <param name="pname"></param>
    ''' <param name="variantName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function notReferencedByAnyPortfolio(ByVal pname As String, ByVal variantName As String) As Boolean

        Dim atleastOneReference As Boolean = False

        For Each kvp As KeyValuePair(Of String, clsConstellation) In projectConstellations.Liste

            'If kvp.Key = calcLastSessionScenarioName() Or kvp.Key = calcLastEditorScenarioName() Then
            If kvp.Key = calcLastSessionScenarioName() Then
                ' nichts tun , die zählen nicht 
            Else
                Dim pvName As String = calcProjektKey(pname, variantName)
                atleastOneReference = atleastOneReference Or kvp.Value.contains(pvName, False)
            End If


        Next

        notReferencedByAnyPortfolio = Not atleastOneReference

    End Function



    ''' <summary>
    ''' löscht den angegebenen timestamp von pname#variantname aus der Datenbank
    ''' speichert den timestamp im Papierkorb
    ''' </summary>
    ''' <param name="pname"></param>
    ''' <param name="variantName"></param>
    ''' <param name="timeStamp"></param>
    ''' <param name="first"></param>
    ''' <remarks></remarks>
    Public Sub deleteProjectVariantTimeStamp(ByRef outputCollection As Collection, _
                                             ByVal pname As String, ByVal variantName As String, _
                                                  ByVal timeStamp As Date, ByRef first As Boolean)

        Dim outputLine As String = ""
        Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
        Dim requestTrash As New Request(awinSettings.databaseURL, awinSettings.databaseName & "Trash", dbUsername, dbPasswort)
        Dim hproj As clsProjekt

        If first Then
            projekthistorie.clear() ' alte Historie löschen
            projekthistorie.liste = request.retrieveProjectHistoryFromDB _
                                   (projectname:=pname, variantName:=variantName, _
                                    storedEarliest:=Date.MinValue, storedLatest:=Date.Now)
            first = False
        End If



        hproj = projekthistorie.ElementAtorBefore(timeStamp)

        If DateDiff(DateInterval.Second, timeStamp, hproj.timeStamp) <> 0 Then
            outputLine = "Fehler:" & timeStamp.ToShortDateString & vbLf & _
            hproj.timeStamp.ToShortDateString
            outputCollection.Add(outputLine)
            'Call MsgBox("hier ist was faul" & timeStamp.ToShortDateString & vbLf & _
            '             hproj.timeStamp.ToShortDateString)
        End If
        timeStamp = hproj.timeStamp

        If IsNothing(hproj) Then
            outputLine = "Timestamp " & timeStamp.ToShortDateString & vbLf & _
                        "zu Projekt " & projekthistorie.First.getShapeText & " nicht gefunden"
            outputCollection.Add(outputLine)
            'Call MsgBox("Timestamp " & timeStamp.ToShortDateString & vbLf & _
            '            "zu Projekt " & projekthistorie.First.getShapeText & " nicht gefunden")

        Else
            ' Speichern im Papierkorb, dann löschen
            'If requestTrash.storeProjectToDB(hproj) Then
            If request.deleteProjectTimestampFromDB(projectname:=pname, variantName:=variantName, _
                                  stored:=timeStamp, userName:=dbUsername) Then
                'Call MsgBox("ok, gelöscht")
            Else
                outputLine = "Fehler beim Löschen von " & pname & ", " & variantName & ", " & _
                              timeStamp.ToShortDateString
                outputCollection.Add(outputLine)
                'Call MsgBox("Fehler beim Löschen von " & pname & ", " & variantName & ", " & _
                '            timeStamp.ToShortDateString)
            End If
            '    Else
            '    ' es ging etwas schief


            '    Call MsgBox("Fehler beim Speichern im Papierkorb:" & vbLf & _
            '                hproj.name & ", " & hproj.timeStamp.ToShortDateString)
            'End If

        End If

    End Sub




    ''' <summary>
    ''' erzeugt die Excel Datei mit den Projekt-Ressourcen Zuordnungen 
    ''' Vorbedingung Ressourcen Datei ist bereits geöffnet
    ''' 
    ''' </summary>
    ''' <param name="typus">
    ''' 0: alle Ressourcen in einer Datei ; 1: pro Rolle eine Datei ; 2: pro Kostenart eine Datei 
    ''' </param>
    ''' <param name="qualifier">
    ''' gibt den Bezeichner der Rolle / Kostenart an 
    ''' </param>
    ''' <remarks></remarks>
    Public Sub awinExportRessZuordnung(ByVal typus As Integer, ByVal qualifier As String)

        Dim anzRollen As Integer
        Dim i As Integer, m As Integer
        Dim heute As Date = Date.Now
        Dim heuteColumn As Integer
        Dim currentRole As String = " "
        Dim kapaValues() As Double
        Dim currentColor As Long
        Dim zeile As Integer = 1
        Dim zeitSpanne As Integer = 6
        Dim rng As Excel.Range, destinationRange As Excel.Range
        Dim bedarfsWerte() As Double
        Dim projWerte() As Double
        Dim mycollection As New Collection
        Dim statusColor As Long = awinSettings.AmpelNichtBewertet
        Dim statusValue As Double = 0.0
        Dim xlsBlattname(2) As String
        Dim colPointer As Integer = 2
        Dim loopi As Integer = 1
        'Dim currentColumn As Integer = 1
        Dim vorausschau As Integer = 3
        Dim cellFormula As String
        Dim personalrange As Excel.Range
        Dim rngSource As Excel.Range
        Dim rngTarget As Excel.Range
        Dim rcol As Integer
        Dim anzPeople As Integer


        Dim startZeile As Integer, endZeile As Integer
        Dim tmpDate As Date


        xlsBlattname(0) = "Summary"
        xlsBlattname(1) = "Zuordnung"
        xlsBlattname(2) = "Kapazität"

        ReDim bedarfsWerte(zeitSpanne - 1)
        ReDim kapaValues(zeitSpanne - 1)

        If typus = 0 Then
            anzRollen = RoleDefinitions.Count
        Else
            anzRollen = 1
        End If

        heuteColumn = getColumnOfDate(heute) + 1


        ' ----------------------------------------------------------
        ' Schreiben Summary 
        '-----------------------------------------------------------
        Try

            With CType(appInstance.Worksheets(xlsBlattname(0)), Global.Microsoft.Office.Interop.Excel.Worksheet)


                ' Löschen der alten Werte 
                rng = .Range(.Cells(2, 1), .Cells(2002, 21))
                rng.Clear()

                ' Schreiben der betrachteten Monate in zeile 1
                If typus = 0 Then
                    CType(.Cells(zeile, 1), Global.Microsoft.Office.Interop.Excel.Range).Value = "Rolle"
                    CType(.Cells(zeile, 2), Global.Microsoft.Office.Interop.Excel.Range).Value = "Projekt"
                Else
                    CType(.Cells(zeile, 1), Global.Microsoft.Office.Interop.Excel.Range).Value = " "
                    CType(.Cells(zeile, 2), Global.Microsoft.Office.Interop.Excel.Range).Value = "Projekt"
                End If


                If awinSettings.zeitEinheit = "PM" Then
                    m = 1
                    CType(.Cells(zeile, 5), Global.Microsoft.Office.Interop.Excel.Range).Value = heute.AddMonths(m)
                    CType(.Cells(zeile, 6), Global.Microsoft.Office.Interop.Excel.Range).Value = heute.AddMonths(m + 1)
                    rng = .Range(.Cells(zeile, 5), .Cells(zeile, 6))


                    With rng
                        .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                        .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                        .NumberFormat = "mmm-yy"
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = Excel.Constants.xlContext
                        .MergeCells = False
                        .ColumnWidth = 10
                    End With


                    destinationRange = .Range(.Cells(zeile, 5), .Cells(zeile, 5 + zeitSpanne - 1))

                    With destinationRange
                        .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                        .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                        .NumberFormat = "mmm-yy"
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = Excel.Constants.xlContext
                        .MergeCells = False
                        .ColumnWidth = 10
                    End With

                    rng.AutoFill(Destination:=destinationRange, Type:=Excel.XlAutoFillType.xlFillMonths)



                ElseIf awinSettings.zeitEinheit = "PW" Then
                ElseIf awinSettings.zeitEinheit = "PT" Then

                End If
                zeile = 2

                Dim sumrangeAnfang As Integer, sumrangeEnde As Integer

                For i = 1 To anzRollen

                    zeile = zeile + 1
                    currentRole = " "
                    If typus = 0 Then
                        With RoleDefinitions.getRoledef(i)
                            currentRole = .name
                            For m = 0 To zeitSpanne - 1
                                kapaValues(m) = .kapazitaet(m + heuteColumn) + _
                                                .externeKapazitaet(m + heuteColumn)
                            Next
                            currentColor = CLng(.farbe)
                        End With
                    ElseIf typus = 1 Then
                        Try
                            With RoleDefinitions.getRoledef(qualifier)
                                currentRole = .name
                                For m = 0 To zeitSpanne - 1
                                    kapaValues(m) = .kapazitaet(m + heuteColumn) + _
                                                    .externeKapazitaet(m + heuteColumn)
                                Next
                                currentColor = CLng(.farbe)
                            End With
                        Catch ex As Exception

                        End Try
                    Else
                        Call MsgBox("Kostenarten noch nicht definiert ...")
                        Exit Sub
                    End If


                    CType(.Cells(zeile, 1), Global.Microsoft.Office.Interop.Excel.Range).Value = currentRole
                    ' jetzt wird der Bereich mit hellgrauer Farbe abgesetzt 
                    rng = .Range(.Cells(zeile - 1, 1), .Cells(zeile, 1 + 5 + zeitSpanne - 2))
                    rng.Interior.Color = awinSettings.AmpelNichtBewertet

                    zeile = zeile + 1

                    mycollection.Add(currentRole)
                    sumrangeAnfang = zeile

                    For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste
                        ' benötigt dieses Projekt in den nächsten Monaten die Rolle <currentRole>


                        If kvp.Value.Start <= getColumnOfDate(Date.Now) + zeitSpanne And _
                            kvp.Value.Start + kvp.Value.anzahlRasterElemente - 1 >= getColumnOfDate(Date.Now) + 1 And _
                            kvp.Value.Status <> ProjektStatus(3) And _
                            kvp.Value.Status <> ProjektStatus(4) Then

                            With kvp.Value
                                'statusValue = 
                                ReDim bedarfsWerte(zeitSpanne - 1)
                                ReDim projWerte(.anzahlRasterElemente - 1)
                                projWerte = .getBedarfeInMonths(mycollection, DiagrammTypen(1))

                                Dim aix As Integer
                                aix = heuteColumn - .Start

                                If aix >= 0 Then
                                    For m = 0 To zeitSpanne - 1
                                        If m + aix <= .anzahlRasterElemente - 1 Then
                                            bedarfsWerte(m) = projWerte(m + aix)
                                        End If
                                    Next
                                Else
                                    For m = 0 To zeitSpanne - 1
                                        If m + aix >= 0 Then
                                            bedarfsWerte(m) = projWerte(m + aix)
                                        End If
                                    Next
                                End If


                            End With

                            ' wenn die Summe größer Null ist , wird eine Zeile in das Excel File eingetragen
                            If bedarfsWerte.Sum > 0 Then
                                ' jetzt werden der Status und die Ampelbewertung errechnet ...
                                Call getStatusColorProject(kvp.Value, 1, 1, " ", statusValue, statusColor)

                                CType(.Cells(zeile, 2), Global.Microsoft.Office.Interop.Excel.Range).Value = kvp.Value.name
                                CType(.Cells(zeile, 3), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color = statusColor
                                CType(.Cells(zeile, 4), Global.Microsoft.Office.Interop.Excel.Range).Value = statusValue
                                CType(.Cells(zeile, 4), Global.Microsoft.Office.Interop.Excel.Range).NumberFormat = "0.00"
                                rng = .Range(.Cells(zeile, 5), .Cells(zeile, 5 + zeitSpanne - 1))
                                rng.Value = bedarfsWerte
                                zeile = zeile + 1
                            End If


                        End If

                    Next

                    sumrangeEnde = zeile - 1

                    ' jetzt wird der Bereich der Projekte mit der entsprechenden Farbe gekennzeichnet 
                    rng = .Range(.Cells(sumrangeAnfang, 1), .Cells(sumrangeEnde, 1))
                    rng.Interior.Color = currentColor

                    ' jetzt wird die Summenformel eingesetzt 
                    CType(.Cells(zeile, 1), Global.Microsoft.Office.Interop.Excel.Range).Value = "Summe"
                    cellFormula = "=SUM(R[" & sumrangeAnfang - zeile & "]C:R[" & sumrangeEnde - zeile & "]C)"
                    For m = 0 To 5
                        CType(.Cells(zeile, 5 + m), Global.Microsoft.Office.Interop.Excel.Range).FormulaR1C1 = cellFormula
                    Next

                    zeile = zeile + 1

                    ' jetzt wird die Kapa eingetragen 
                    CType(.Cells(zeile, 1), Global.Microsoft.Office.Interop.Excel.Range).Value = "Kapazität"
                    For m = 0 To 5
                        CType(.Cells(zeile, 5 + m), Global.Microsoft.Office.Interop.Excel.Range).Value = kapaValues(m)
                    Next

                    ' jetzt wird der Bereich mit hellgrauer Farbe abgesetzt 
                    rng = .Range(.Cells(zeile - 1, 1), .Cells(zeile, 1 + 5 + zeitSpanne - 2))
                    rng.Interior.Color = awinSettings.AmpelNichtBewertet

                    zeile = zeile + 3

                    Try
                        mycollection.Clear()
                    Catch ex As Exception
                        mycollection = New Collection
                    End Try

                Next

            End With

        Catch ex As Exception
            Call MsgBox("Register " & xlsBlattname(0) & " existiert nicht ")
        End Try

        '
        ' wenn nur die Zusammenfassung gefragt war: dann wird jetzt die Routine verlassen 
        '
        If typus = 0 Then
            Exit Sub
        End If

        ' ----------------------------------------------------------
        ' Schreiben der Feinplanungs-Sheet - für jeden Monat der Vorausschau eines 
        '-----------------------------------------------------------

        Dim projekttitelZeile As Integer = 3
        Dim bewertungszeile As Integer = 4
        mycollection.Add(currentRole)

        Dim currentWS As Excel.Worksheet

        For loopi = 1 To vorausschau
            Dim blattName As String = xlsBlattname(1) & " " & Date.Now.AddMonths(loopi).ToString("MMM yy")

            Try


                ' suchen nach dem Register Blattname, wenn es bereits existiert wird es überschrieben
                ' wenn es noch nicht existiert, wird es angelegt
                currentWS = CType(appInstance.Worksheets(blattName), Global.Microsoft.Office.Interop.Excel.Worksheet)


                ' wenn das schon existiert, wird es einfach überschrieben 
                'Try
                '    currentWS.Name = blattName & " " & Date.Now.ToLongDateString
                '    'currentWS.Delete()

                'Catch ex1 As Exception

                '    Call MsgBox("Tabelle " & blattName & " kann nicht umbenannt werden ")
                '    Exit Sub

                'End Try



            Catch ex As Exception

                Try

                    With CType(appInstance.Worksheets.Add(Before:=appInstance.Worksheets(xlsBlattname(0))), _
                                    Global.Microsoft.Office.Interop.Excel.Worksheet)
                        .Name = blattName

                    End With

                Catch ex2 As Exception

                    Call MsgBox("Tabelle Summary nicht vorhanden ... " & vbLf & ex2.Message)
                    Exit Sub

                End Try


            End Try

            ' jetzt muss es auf alle Fälle existieren 
            currentWS = CType(appInstance.Worksheets(blattName), Global.Microsoft.Office.Interop.Excel.Worksheet)



            Try
                ' jetzt zurücksetzen der Planungs-Unterstützung, Register Zuordnung 

                With currentWS

                    .Unprotect()

                    ' -------------------------------------------------------
                    ' Inhalt leeren ...
                    ' -------------------------------------------------------
                    With .Range(.Cells(1, 1), .Cells(500, 500))
                        .ClearContents()
                        .Interior.Color = awinSettings.AmpelNichtBewertet
                    End With

                    ' -------------------------------------------------------
                    ' Überschrift schreiben 
                    ' -------------------------------------------------------
                    With CType(.Cells(1, 1), Global.Microsoft.Office.Interop.Excel.Range)
                        .Value = xlsBlattname(1) & " " & qualifier & " " & Date.Now.AddMonths(loopi).ToString("MMM yy")
                        .Font.Size = 20
                        .Font.Bold = True
                    End With

                    CType(.Rows(1), Global.Microsoft.Office.Interop.Excel.Range).RowHeight = 48


                    CType(.Rows("2:100"), Global.Microsoft.Office.Interop.Excel.Range).RowHeight = 16
                    CType(.Rows(5), Global.Microsoft.Office.Interop.Excel.Range).RowHeight = 1
                    CType(.Columns(1), Global.Microsoft.Office.Interop.Excel.Range).ColumnWidth = 25
                    '.columns(2).columnwidth = 12
                    CType(.Columns("B:CV"), Global.Microsoft.Office.Interop.Excel.Range).ColumnWidth = 8

                    ' -------------------------------------------------------
                    ' 3. Zeile schreiben: Kapazität und Projekt-Titel    
                    ' -------------------------------------------------------
                    With CType(.Rows(projekttitelZeile), Global.Microsoft.Office.Interop.Excel.Range)
                        .RowHeight = 96
                        .Font.Size = 12
                        .Font.Bold = True
                    End With

                    With CType(.Cells(projekttitelZeile, 2), Global.Microsoft.Office.Interop.Excel.Range)
                        .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                        .VerticalAlignment = Excel.XlVAlign.xlVAlignBottom
                        .Value = "Kapazität"
                        .Orientation = 90
                        '.AddIndent = True
                        '.IndentLevel = 1
                    End With

                    With .Range(.Cells(projekttitelZeile, 4), .Cells(projekttitelZeile, 100))
                        .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                        .VerticalAlignment = Excel.XlVAlign.xlVAlignBottom
                        .WrapText = True
                        .Orientation = 90
                        '.AddIndent = True
                        '.IndentLevel = 1
                    End With


                    ' -------------------------------------------------------
                    ' 4. Zeile schreiben: Projekt-Bewertung  
                    ' -------------------------------------------------------

                    With .Range(.Cells(bewertungszeile, 3), .Cells(bewertungszeile, 100))
                        .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                        .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                        .Font.Size() = 8
                        .Font.Bold = False
                        .NumberFormat = "0.0#"
                    End With





                End With
            Catch ex As Exception
                Call MsgBox("Fehler mit Tabellenblatt " & blattName)
                Exit Sub
            End Try


            ' jetzt werden die Zuordnungs-Werte ausgelesen 
            ' erweitert 26.7.13 - Planungshilfe für den Ressourcen Manager erstellen
            Try
                ' jetzt werden erst mal die Personen ausgelesen und deren Kapa im besagten Monat 
                With CType(appInstance.Worksheets(xlsBlattname(2)), Global.Microsoft.Office.Interop.Excel.Worksheet)

                    ' wenn noch keine Personen angelegt sind -> Exit 
                    personalrange = .Range("Personenliste")
                    startZeile = personalrange.Row
                    rcol = personalrange.Column
                    ' die letzte Zeile ist letzte Person   
                    endZeile = startZeile + personalrange.Rows.Count - 1
                    anzPeople = endZeile - startZeile + 1
                    rngSource = .Range(.Cells(startZeile, rcol), .Cells(endZeile, rcol))
                    rngTarget = CType(CType(appInstance.Worksheets(blattName), Global.Microsoft.Office.Interop.Excel.Worksheet).Cells(6, 1), _
                                                Global.Microsoft.Office.Interop.Excel.Range)

                    rngSource.Copy(rngTarget)


                    ' jetzt wird die Spalte gesucht, wo die Werte für den nächsten Monat stehen 
                    rcol = 2
                    Dim found As Boolean
                    tmpDate = CDate(CType(.Cells(1, rcol), Global.Microsoft.Office.Interop.Excel.Range).Value)

                    If DateDiff(DateInterval.Month, heute, tmpDate) > 0 Then
                        found = True
                    Else
                        found = False
                    End If

                    Do While Not found And rcol < 240
                        rcol = rcol + 1
                        tmpDate = CDate(CType(.Cells(1, rcol), Global.Microsoft.Office.Interop.Excel.Range).Value)
                        If DateDiff(DateInterval.Month, heute, tmpDate) > loopi - 1 Then
                            found = True
                        Else
                            found = False
                        End If
                    Loop

                    If found Then
                        ' jetzt müssen die Werte referenziert werden 


                        Dim k As Integer
                        For k = startZeile To endZeile
                            cellFormula = "=" & xlsBlattname(2).Trim & "!R[-4]C[" & rcol - 2 & "]"
                            CType(currentWS.Cells(k - startZeile + 6, 2), _
                                    Global.Microsoft.Office.Interop.Excel.Range).FormulaR1C1 = cellFormula
                        Next

                        CType(currentWS.Cells(endZeile - startZeile + 7, 1), _
                                Global.Microsoft.Office.Interop.Excel.Range).Value = "Extern"

                    Else
                        Call MsgBox("keine Werte für Folge-Monate von " & heute.ToShortDateString & " gefunden ...")
                        Exit Sub
                    End If


                End With
            Catch ex As Exception
                Call MsgBox("es sind keine Mitarbeiter im Register " & xlsBlattname(2) & " angelegt")
                Exit Sub
            End Try


            'currentColumn = currentColumn + 3
            ' jetzt wird der Rest der Zuordnungs-Datei geschrieben : die Projekt-Daten

            With currentWS

                Dim anzProjekte As Integer = 0
                For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

                    Dim tmpWert As Double = 0.0
                    If kvp.Value.Start <= getColumnOfDate(Date.Now) + loopi And _
                              kvp.Value.Status <> ProjektStatus(3) And _
                              kvp.Value.Status <> ProjektStatus(4) Then

                        With kvp.Value

                            tmpWert = .getBedarfeInMonth(mycollection, DiagrammTypen(1), heuteColumn + loopi - 1)

                        End With

                        ' wenn die Summe größer Null ist , wird eine Zeile in das Excel File eingetragen
                        If tmpWert > 0 Then
                            ' jetzt werden der Status und die Ampelbewertung errechnet ...
                            Call getStatusColorProject(kvp.Value, 1, 1, " ", statusValue, statusColor)

                            CType(.Cells(projekttitelZeile, 4 + anzProjekte), Global.Microsoft.Office.Interop.Excel.Range).Value = kvp.Value.name

                            ' Schreiben der Summen-Formel 

                            cellFormula = "=SUM(R[-" & anzPeople + 1 & "]C:R[-1]C)"
                            CType(.Cells(6 + anzPeople + 1, 4 + anzProjekte), Global.Microsoft.Office.Interop.Excel.Range).FormulaR1C1 = cellFormula

                            ' Schreiben des Bedarfs
                            CType(.Cells(6 + anzPeople + 2, 4 + anzProjekte), Global.Microsoft.Office.Interop.Excel.Range).Value = tmpWert

                            ' Schreiben der Farbe
                            CType(.Cells(bewertungszeile, 4 + anzProjekte), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color = statusColor

                            ' Schreiben des Wertes 
                            CType(.Cells(bewertungszeile, 4 + anzProjekte), Global.Microsoft.Office.Interop.Excel.Range).Value = statusValue

                            'currentColumn = currentColumn + 1
                            anzProjekte = anzProjekte + 1


                        End If


                    End If

                Next

                If anzProjekte > 0 Then

                    ' Schreiben der Zeilen-Summen: Summe Zuordnung pro MA
                    cellFormula = "=SUM(RC[1]:RC[" & anzProjekte & "])"

                    Dim k As Integer
                    For k = 6 To 6 + anzPeople
                        CType(.Cells(k, 3), Global.Microsoft.Office.Interop.Excel.Range).FormulaR1C1 = cellFormula
                    Next

                    CType(.Cells(7 + anzPeople, 5 + anzProjekte), Global.Microsoft.Office.Interop.Excel.Range).Value = ""
                    CType(.Cells(8 + anzPeople, 4 + anzProjekte), Global.Microsoft.Office.Interop.Excel.Range).Value = "Projekt-Bedarf"
                    CType(.Columns(3 + anzProjekte + 1), Global.Microsoft.Office.Interop.Excel.Range).ColumnWidth = 25

                    ' Schreiben des Prozentsatzes wieviel des Projektbedarfes wird durch interne abgedeckt 
                    cellFormula = "=SUM(R[-" & anzPeople + 2 & "]C:R[-3]C)/SUM(RC[2]:RC[" & anzProjekte + 1 & "])"

                    With CType(.Cells(8 + anzPeople, 2), Global.Microsoft.Office.Interop.Excel.Range)
                        .FormulaR1C1 = cellFormula
                        .NumberFormat = "0%"
                        .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                    End With

                    ' die Eingabe Felder farblos machen  
                    .Range(.Cells(6, 4), .Cells(6 + anzPeople, 4 + anzProjekte - 1)).Interior.ColorIndex = Excel.Constants.xlNone




                    With .Range(.Cells(6, 1), .Cells(6 + anzPeople, 1))
                        .Interior.Color = awinSettings.AmpelNichtBewertet
                    End With

                    With .Range(.Cells(6, 2), .Cells(6 + anzPeople, 2))
                        .Font.Size = 12
                        .Font.Bold = True
                        .HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
                        .AddIndent = True
                        .IndentLevel = 2
                    End With

                    ' die vertikalen Summen-Felder etwas einrücken ..
                    With .Range(.Cells(6, 3), .Cells(6 + anzPeople, 3))
                        .HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
                        .AddIndent = True
                        .IndentLevel = 2
                    End With


                    With .Range(.Cells(8 + anzPeople, 4), .Cells(8 + anzPeople, 4 + anzProjekte))
                        .Font.Size = 12
                        .Font.Bold = True
                    End With

                Else
                    'CType(.Cells(6, 6), Global.Microsoft.Office.Interop.Excel.Range).value = "kein Projekt-Bedarf für diese Rolle"
                    CType(.Cells(8 + anzPeople, 4 + 1), Global.Microsoft.Office.Interop.Excel.Range).Value = "kein Projekt-Bedarf für diese Rolle"
                End If


                ' jetzt wird das Zuordnungs-Blatt geschützt 
                .Range(.Cells(1, 1), .Cells(1 + 8, 1 + 4 + anzProjekte)).Locked = True
                ' Freigeben des Eingabe Bereiches 
                .Range(.Cells(6, 4), .Cells(6 + anzPeople, 4 + anzProjekte - 1)).Locked = False
                ' Auskommentiert, weil es zu Fehlern führt; ausserdem ist nicht mehr klar, wozu das überhaupt benötigt wird 
                'CType(.Cells(6, 4), Global.Microsoft.Office.Interop.Excel.Range).Activate()

                .Protect()

            End With

        Next








    End Sub



    ''' <summary>
    ''' liest die im Diretory ../ressource manager liegenden detaillierten Kapa files zu den Rollen aus
    ''' und hinterlegt es an entsprechender Stelle im hrole.kapazitaet
    ''' wenn die Details als Rollen angelegt sind, dann werden diese Rollen gleich mitausgelesen 
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
        Dim lastSpalte As Integer


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
                    lastSpalte = CType(currentWS.Cells(1, 2000), Global.Microsoft.Office.Interop.Excel.Range).End(Excel.XlDirection.xlToLeft).Column

                    ' bevor jetzt die eigentliche Kapa dieser Rolle aus intern_sum ausgelesen wird, wird geschaut, ob 
                    ' es eine zusammengesetzte Rolle ist
                    ' das wird dadurch entschieden, ob bis zur summenzeile bekannte Rollen auftauchen. Das sind dann die Sub-Roles 


                    Dim atleastOneSubRole As Boolean = False
                    Dim aktzeile As Integer = 2
                    Do While aktzeile < summenZeile

                        Dim subRoleName As String = CStr(CType(currentWS.Cells(aktzeile, spalte - 1), Excel.Range).Value)

                        If Not IsNothing(subRoleName) Then
                            subRoleName = subRoleName.Trim
                            If subRoleName.Length > 0 And RoleDefinitions.containsName(subRoleName) Then

                                Dim subRole As clsRollenDefinition = RoleDefinitions.getRoledef(subRoleName)

                                Try
                                    atleastOneSubRole = True
                                    ' es ist eine Sub-Rolle

                                    hrole.addSubRole(subRole.UID, subRoleName, RoleDefinitions.Count)

                                    spalte = 2
                                    tmpDate = CDate(CType(currentWS.Cells(1, spalte), Excel.Range).Value)

                                    ' erstmal dahin positionieren, wo das Datum auch mit StartOfCalendar harmoniert 
                                    Do While DateDiff(DateInterval.Month, StartofCalendar, tmpDate) < 0 And spalte <= lastSpalte
                                        spalte = spalte + 1
                                        tmpDate = CDate(CType(currentWS.Cells(1, spalte), Excel.Range).Value)
                                    Loop

                                    Do While spalte < 241 And spalte <= lastSpalte

                                        index = getColumnOfDate(tmpDate)
                                        If index >= 1 Then
                                            tmpKapa = CDbl(CType(currentWS.Cells(aktzeile, spalte), Excel.Range).Value)

                                            If index <= 240 And index > 0 And tmpKapa >= 0 Then
                                                subRole.kapazitaet(index) = tmpKapa
                                            End If
                                        End If

                                        spalte = spalte + 1
                                        tmpDate = CDate(CType(currentWS.Cells(1, spalte), Excel.Range).Value)
                                    Loop

                                Catch ex As Exception

                                End Try


                            End If

                        End If

                        aktzeile = aktzeile + 1
                        ' jetzt spalte wieder auf 2 setzen 
                        spalte = 2
                    Loop

                    ' die internen Kapas einer Sammelrolle sind NULL 
                    If atleastOneSubRole Then
                        For i As Integer = 1 To 240
                            hrole.kapazitaet(i) = 0
                        Next
                    End If



                    Try
                        extSummenZeile = currentWS.Range("extern_sum").Row
                    Catch ex As Exception
                        extSummenZeile = 0
                    End Try

                    tmpDate = CDate(CType(currentWS.Cells(1, spalte), Excel.Range).Value)

                    Do While DateDiff(DateInterval.Month, StartofCalendar, tmpDate) > 0 And _
                            spalte < 241 And spalte <= lastSpalte
                        index = getColumnOfDate(tmpDate)
                        tmpKapa = CDbl(CType(currentWS.Cells(summenZeile, spalte), Excel.Range).Value)


                        If extSummenZeile > 0 Then
                            extTmpKapa = CDbl(CType(currentWS.Cells(extSummenZeile, spalte), Excel.Range).Value)
                        Else
                            extTmpKapa = 0.0
                        End If

                        If index <= 240 And index > 0 Then

                            If atleastOneSubRole Then
                                ' alles ist Null , wird erst später aufgrund der Sub-Rollen berechnet 
                            Else
                                If tmpKapa >= 0 Then
                                    hrole.kapazitaet(index) = tmpKapa
                                End If
                            End If

                            If extTmpKapa >= 0 Then
                                hrole.externeKapazitaet(index) = extTmpKapa
                            End If


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
    ''' liest das im Diretory ../ressource manager evt. liegende File 'Urlaubsplaner*.xlsx' File  aus
    ''' und hinterlegt an entsprechender Stelle im hrole.kapazitaet die verfügbaren Tage der entsprechenden Rolle
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub readUrlOfRole(ByVal kapaFileName As String)

        Dim ok As Boolean = True
        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim formerSU As Boolean = appInstance.ScreenUpdating
        Dim msgtxt As String = ""
        Dim fehler As Boolean = False
        Dim oPCollection As New Collection

        Dim kapaWB As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim spalte As Integer = 2
        Dim firstUrlspalte As Integer = 5
        Dim noColor As Integer = -4142
        Dim whiteColor As Integer = 2
        Dim currentWS As Excel.Worksheet
        Dim index As Integer
        Dim tmpDate As Date

        Dim year As Integer = DatePart(DateInterval.Year, Date.Now)
        Dim anzMonthDays As Integer = 0
        Dim colDate As Integer = 0
        Dim anzDays As Integer = 0

        Dim lastZeile As Integer
        Dim lastSpalte As Integer
        Dim monthDays As New SortedList(Of Integer, Integer)

        Dim hrole As New clsRollenDefinition
        Dim rolename As String = ""

        Dim outPutCollection As New Collection

        If formerEE Then
            appInstance.EnableEvents = False
        End If

        If formerSU Then
            appInstance.ScreenUpdating = False
        End If

        enableOnUpdate = False

        ' öffnen des Files 
        If My.Computer.FileSystem.FileExists(kapaFileName) Then

            Try
                kapaWB = appInstance.Workbooks.Open(kapaFileName)

                Try
                    For index = 1 To appInstance.Worksheets.Count

                        'If Not ok Then
                        '    Exit For
                        'End If
                 

                        currentWS = CType(appInstance.Worksheets(index), Global.Microsoft.Office.Interop.Excel.Worksheet)
                        Dim hstr() As String = Split(currentWS.Name, "Halbjahr", , )
                        If hstr.Length > 1 Then

                            ok = True
                            ' Auslesen der Jahreszahl, falls vorhanden
                            If Not IsNothing(CType(currentWS.Cells(1, 2), Global.Microsoft.Office.Interop.Excel.Range).Value) Then
                                year = CType(currentWS.Cells(1, 2), Global.Microsoft.Office.Interop.Excel.Range).Value
                            End If

                            monthDays.Clear()
                            anzDays = 0

                            lastZeile = CType(currentWS.Cells(2000, 1), Global.Microsoft.Office.Interop.Excel.Range).End(Excel.XlDirection.xlUp).Row
                            lastSpalte = CType(currentWS.Cells(4, 2000), Global.Microsoft.Office.Interop.Excel.Range).End(Excel.XlDirection.xlToLeft).Column

                            Dim vglColor As Integer = noColor         ' keine Farbe
                            Dim i As Integer = firstUrlspalte

                            While ok And i <= lastSpalte

                                If vglColor <> CType(currentWS.Cells(1, i), Global.Microsoft.Office.Interop.Excel.Range).Interior.ColorIndex Then
                                    ok = (anzDays = anzMonthDays) Or (anzDays = 0)
                                    vglColor = CType(currentWS.Cells(1, i), Global.Microsoft.Office.Interop.Excel.Range).Interior.ColorIndex
                                    anzDays = 1
                                Else
                                    If CType(currentWS.Cells(1, i), Global.Microsoft.Office.Interop.Excel.Range).Text <> "" Then
                                        Dim monthName As String = CType(currentWS.Cells(1, i), Global.Microsoft.Office.Interop.Excel.Range).Text
                                        ' ''Dim strDate As String = "01." & monthName & " " & year
                                        ' ''Dim hdate As DateTime = DateValue(strDate)

                                        Dim isdate As Boolean = DateTime.TryParse(monthName & " " & year.ToString, tmpDate)
                                        If isdate Then
                                            colDate = getColumnOfDate(tmpDate)
                                            anzMonthDays = DateTime.DaysInMonth(year, Month(tmpDate))
                                            monthDays.Add(colDate, anzMonthDays)
                                        End If
                                    End If

                                    anzDays = anzDays + 1
                                End If

                                i = i + 1
                            End While


                            If Not ok Then

                                fehler = True

                                If awinSettings.englishLanguage Then
                                    msgtxt = "Error reading planning holidays: Please check die calendar in this file ..."
                                Else
                                    msgtxt = "Fehler beim Lesen der Urlaubsplanung: Bitte prüfen Sie die Korrektheit des Kalenders ..."
                                End If
                                If Not oPCollection.Contains(msgtxt) Then
                                    oPCollection.Add(msgtxt, msgtxt)
                                End If
                                'Call MsgBox(msgtxt)

                                If formerEE Then
                                    appInstance.EnableEvents = True
                                End If

                                If formerSU Then
                                    appInstance.ScreenUpdating = True
                                End If

                                enableOnUpdate = True
                                If awinSettings.englishLanguage Then
                                    msgtxt = "Your planning holidays couldn't be read, because of problems"
                                Else
                                    msgtxt = "Ihre Urlaubsplanung konnte nicht berücksichtigt werden"
                                End If
                                If Not oPCollection.Contains(msgtxt) Then
                                    oPCollection.Add(msgtxt, msgtxt)
                                End If
                                'Call showOutPut(oPCollection, "Lesen Urlaubsplanung wurde mit Fehler abgeschlossen", "Meldungen zu Lesen Urlaubsplanung")
                                Throw New ArgumentException(msgtxt)
                            Else

                                For iZ = 5 To lastZeile

                                    rolename = CType(currentWS.Cells(iZ, 2), Global.Microsoft.Office.Interop.Excel.Range).Text
                                    If rolename <> "" Then
                                        hrole = RoleDefinitions.getRoledef(rolename)
                                        If Not IsNothing(hrole) Then

                                            Dim iSp As Integer = firstUrlspalte
                                            Dim anzArbTage As Double = 0
                                            Dim anzArbStd As Double = 0

                                            For Each kvp As KeyValuePair(Of Integer, Integer) In monthDays

                                                Dim colOfDate As Integer = kvp.Key
                                                anzDays = kvp.Value
                                                For sp = iSp + 0 To iSp + anzDays - 1

                                                    If iSp <= lastSpalte Then
                                                        Dim hint As Integer = CInt(CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Interior.ColorIndex)

                                                        If CInt(CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Interior.ColorIndex) = noColor _
                                                            Or CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Interior.ColorIndex = whiteColor Then

                                                            If Not IsNothing(CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Value) Then

                                                                If CDbl(CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Value) >= 0 And _
                                                                       CDbl(CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Value) <= 24 Then
                                                                    anzArbStd = anzArbStd + CDbl(CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                                                Else
                                                                    If awinSettings.englishLanguage Then
                                                                        msgtxt = "Error reading the  amount of working hours of " & hrole.name & " ..."
                                                                    Else
                                                                        msgtxt = "Fehler beim Lesen der Anzahl zu leistenden Arbeitsstunden " & hrole.name & " ..."
                                                                    End If
                                                                    If Not oPCollection.Contains(msgtxt) Then
                                                                        oPCollection.Add(msgtxt, msgtxt)
                                                                    End If
                                                                    'Call MsgBox(msgtxt)
                                                                    fehler = True
                                                                    Throw New ArgumentException(msgtxt)
                                                                End If


                                                            Else
                                                                ' Dim colorInddown As Integer = CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Borders(XlBordersIndex.xlDiagonalDown).ColorIndex
                                                                Dim colorIndup As Integer = CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Borders(XlBordersIndex.xlDiagonalUp).ColorIndex

                                                                ' Wenn das Feld nicht durch einen Diagonalen Strich gekennzeichnet ist
                                                                If CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Borders(XlBordersIndex.xlDiagonalUp).ColorIndex = noColor Then
                                                                    anzArbStd = anzArbStd + 8
                                                                Else
                                                                    ' freier Tag für Teilzeitbeschäftigte
                                                                End If

                                                            End If
                                                        End If
                                                    Else
                                                        If awinSettings.englishLanguage Then
                                                            msgtxt = "Error reading the amount of working days of " & hrole.name & " ..."
                                                        Else
                                                            msgtxt = "Fehler beim Lesen der verfügbaren Arbeitstage von " & hrole.name & " ..."
                                                        End If
                                                        fehler = True
                                                        If Not oPCollection.Contains(msgtxt) Then
                                                            oPCollection.Add(msgtxt, msgtxt)
                                                        End If
                                                        Throw New ArgumentException(msgtxt)
                                                    End If

                                                Next

                                                anzArbTage = anzArbStd / 8
                                                hrole.kapazitaet(colOfDate) = anzArbTage
                                                iSp = iSp + anzDays
                                                anzArbTage = 0              ' Anzahl Arbeitstage wieder zurücksetzen für den nächsten Monat
                                                anzArbStd = 0               ' Anzahl zu leistender Arbeitsstunden wieder zurücksetzen für den nächsten Monat

                                            Next

                                        Else

                                            If awinSettings.englishLanguage Then
                                                msgtxt = "Role " & rolename & " not defined ..."
                                            Else
                                                msgtxt = "Rolle " & rolename & " nicht definiert ..."
                                            End If
                                            If Not oPCollection.Contains(msgtxt) Then
                                                oPCollection.Add(msgtxt, msgtxt)
                                            End If
                                            'Call MsgBox(msgtxt)
                                            fehler = True
                                        End If
                                    Else

                                        If awinSettings.englishLanguage Then
                                            msgtxt = "No Name of role given ..."
                                        Else
                                            msgtxt = "kein Rollenname angegeben ..."
                                        End If
                                        If Not oPCollection.Contains(msgtxt) Then
                                            oPCollection.Add(msgtxt, msgtxt)
                                        End If
                                        'Call MsgBox(msgtxt)
                                    End If

                                Next iZ

                            End If   ' ende von if not OK
                        Else
                            If awinSettings.visboDebug Then

                                If awinSettings.englishLanguage Then
                                    msgtxt = "Worksheet " & hstr(0) & "doesn't belongs to planning holidays ..."
                                Else
                                    msgtxt = "Worksheet" & hstr(0) & " gehört nicht zum Urlaubsplaner ..."
                                End If
                                If Not oPCollection.Contains(msgtxt) Then
                                    oPCollection.Add(msgtxt, msgtxt)
                                End If
                                Call MsgBox(msgtxt)
                            End If

                        End If

                    Next index


                Catch ex2 As Exception
                    If fehler Then
                        'Call MsgBox(msgtxt)
                        Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
                        RoleDefinitions = request.retrieveRolesFromDB(DateTime.Now)

                        msgtxt = "Es wurden nun die Kapazitäten aus der Datenbank gelesen ..."
                        If awinSettings.englishLanguage Then
                            msgtxt = "Therefore read the capacity of every Role from the DB  ..."
                        End If
                        If Not oPCollection.Contains(msgtxt) Then
                            oPCollection.Add(msgtxt, msgtxt)
                        End If
                        Call MsgBox(msgtxt)
                    End If
                End Try

                'kapaWB.Close(SaveChanges:=False)
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
        kapaWB.Close(SaveChanges:=False)

        Call showOutPut(oPCollection, "Meldungen zu Lesen Urlaubsplanung", "Folgende Problem sind beim Lesen der Urlaubsplanung aufgetreten")

        ' ''If outPutCollection.Count > 0 Then
        ' ''    Call showOutPut(outPutCollection, _
        ' ''                    "Meldungen Einlesevorgang Urlaubsdatei", _
        ' ''                    "zum Zeitpunkt " & storedAtOrBefore.ToString & " aufgeführte Rolle nicht definiert")
        ' ''End If


    End Sub


    ''' <summary>
    ''' liefert die Namen aller Projekte im Show, die nicht zum angegebenen Filter passen ...
    ''' </summary>
    ''' <param name="filterName"></param>
    ''' <remarks></remarks>
    Friend Function getProjectNamesNotFittingToFilter(ByVal filterName As String) As Collection

        Dim nameCollection As New Collection
        Dim filter As New clsFilter
        Dim ok As Boolean = False

        Dim todoListe As New Collection


        filter = filterDefinitions.retrieveFilter("Last")

        If IsNothing(filter) Then

            ' nichts tun und Showprojekte bleibt unverändert ... 
        Else

            For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

                If Not filter.isEmpty Then
                    ok = filter.doesNotBlock(kvp.Value)
                Else
                    ok = True
                End If

                If Not ok Then
                    ' aus Showprojekte rausnehmen und Projekt-Tafel aktualisieren 
                    Try
                        nameCollection.Add(kvp.Value.name)
                    Catch ex As Exception

                    End Try
                Else

                End If

            Next

            ' Liste gefüllt mit Projekte, die auf den aktuellen Filter passen

        End If

        getProjectNamesNotFittingToFilter = nameCollection

    End Function

    ''' <summary>
    ''' baut aus der Datenbank die Projekt-Varianten Liste auf, die zu dem gegeb. Zeitpunkt bereits in der Datenbank existiert haben 
    ''' </summary>
    ''' <remarks></remarks>
    Friend Function buildPvNamesList(ByVal storedAtOrBefore As Date) As SortedList(Of String, String)

        Dim zeitraumVon As Date = StartofCalendar
        Dim zeitraumbis As Date = StartofCalendar.AddYears(50)

        If showRangeLeft > 0 And showRangeRight > showRangeLeft Then
            ' es ist ein Zeitraum definiert 
            zeitraumVon = getDateofColumn(showRangeLeft, False)
            zeitraumbis = getDateofColumn(showRangeRight, True)
        End If


        Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
        buildPvNamesList = request.retrieveProjectVariantNamesFromDB(zeitraumVon, zeitraumbis, storedAtOrBefore)

    End Function

    ' wird nicht mehr verwendet - jetzt mit upDateTreeView gelöst 
    '' ''' <summary>
    '' ''' baut den aktuell gültigen Treeview auf  
    '' ''' </summary>
    '' ''' <remarks></remarks>
    ''Friend Sub buildTreeview(ByRef projektHistorien As clsProjektDBInfos, _
    ''                          ByRef TreeviewProjekte As TreeView, _
    ''                          ByRef pvNamesList As SortedList(Of String, String), _
    ''                          ByVal constellation As clsConstellation, _
    ''                          ByVal aKtionskennung As Integer, _
    ''                          ByVal quickList As Boolean, _
    ''                          ByVal storedAtOrBefore As Date)

    ''    Dim nodeLevel0 As TreeNode
    ''    Dim zeitraumVon As Date = StartofCalendar
    ''    Dim zeitraumbis As Date = StartofCalendar.AddYears(20)
    ''    'Dim storedHeute As Date = Now
    ''    Dim storedGestern As Date = StartofCalendar
    ''    Dim pname As String = ""
    ''    Dim variantName As String = ""
    ''    Dim loadErrorMsg As String = ""

    ''    If showRangeLeft > 0 And showRangeRight > showRangeLeft Then
    ''        ' es ist ein Zeitraum definiert 
    ''        zeitraumVon = getDateofColumn(showRangeLeft, False)
    ''        zeitraumbis = getDateofColumn(showRangeRight, True)
    ''    End If

    ''    ' steuert, ob erstmal nur Projekt-Namen, Varianten-Namen gelesen werden 
    ''    ' geht wesentlich schneller, wenn es sich um eine Datenbank mit sehr vielen Projekten handelt ... 


    ''    Dim deletedProj As Integer = 0



    ''    ' alles zurücksetzen 
    ''    projektHistorien.clear()

    ''    With TreeviewProjekte
    ''        .Nodes.Clear()
    ''    End With


    ''    ' Alle Projekte aus DB
    ''    ' projekteInDB = request.retrieveProjectsFromDB(pname, variantName, zeitraumVon, zeitraumbis, storedGestern, storedHeute, True)

    ''    Select Case aKtionskennung

    ''        Case PTTvActions.delFromDB
    ''            Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)

    ''            pname = ""
    ''            variantName = ""

    ''            pvNamesList = request.retrieveProjectVariantNamesFromDB(zeitraumVon, zeitraumbis, storedAtOrBefore)
    ''            quickList = True
    ''            loadErrorMsg = "es gibt keine Projekte in der Datenbank"

    ''        Case PTTvActions.delAllExceptFromDB
    ''            Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)

    ''            pname = ""
    ''            variantName = ""

    ''            pvNamesList = request.retrieveProjectVariantNamesFromDB(zeitraumVon, zeitraumbis, storedAtOrBefore)
    ''            quickList = True
    ''            loadErrorMsg = "es gibt keine Projekte in der Datenbank"

    ''        Case PTTvActions.delFromSession

    ''            loadErrorMsg = "es sind keine Projekte geladen"

    ''        Case PTTvActions.chgInSession

    ''            loadErrorMsg = "es sind keine Projekte geladen"

    ''        Case PTTvActions.loadPVS    ' ur: 30.01.2015: aktuell nicht benutzt!!!
    ''            Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
    ''            'Dim requestTrash As New Request(awinSettings.databaseURL, awinSettings.databaseName & "Trash", dbUsername, dbPasswort)

    ''            pname = ""
    ''            variantName = ""

    ''            'ur: 25.01.2015 hier muss die "aktuelleGesamtListe.liste reduziert werden, da evt. ein Filter gesetzt wurde!!!!
    ''            ' tk das applyFilter wird nachher gemacht , ausnahmslos für alle 
    ''            'aktuelleGesamtListe.liste = request.retrieveProjectsFromDB(pname, variantName, zeitraumVon, zeitraumbis, storedGestern, storedAtOrBefore, True)
    ''            loadErrorMsg = "es gibt keine Projekte in der Datenbank"

    ''        Case PTTvActions.loadPV

    ''            Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)

    ''            pname = ""
    ''            variantName = ""
    ''            pvNamesList = request.retrieveProjectVariantNamesFromDB(zeitraumVon, zeitraumbis, storedAtOrBefore)
    ''            quickList = True
    ''            loadErrorMsg = "es gibt keine Projekte in der Datenbank"


    ''        Case PTTvActions.activateV
    ''            loadErrorMsg = "es sind keine Projekte geladen"

    ''        Case PTTvActions.deleteV
    ''            loadErrorMsg = "es sind keine Projekte geladen"


    ''    End Select

    ''    ' '' jetzt wird der Filter angewendet, wenn er angewendet werden soll 
    ''    ' '' das wird jetzt in der Routine mitgegeben 
    ''    ''If applyFilter And aktuelleGesamtListe.Count > 0 Then
    ''    ''    aktuelleGesamtListe = reduzierenWgFilter(aktuelleGesamtListe)
    ''    ''End If

    ''    'Dim aktuelleGesamtListe As clsProjekteAlle = AlleProjekte.createCopy(filteredBy:=constellation)
    ''    If Not IsNothing(constellation) Or pvNamesList.Count >= 1 Then

    ''        With TreeviewProjekte

    ''            .CheckBoxes = True

    ''            Dim projektliste As Collection

    ''            If quickList Then
    ''                projektliste = New Collection
    ''                For Each kvp As KeyValuePair(Of String, String) In pvNamesList
    ''                    Dim tmpName As String = kvp.Key
    ''                    If tmpName.Contains("#") Then
    ''                        Dim tmpStr() As String = tmpName.Split(New Char() {CChar("#")})
    ''                        If Not projektliste.Contains(tmpStr(0)) Then
    ''                            projektliste.Add(tmpStr(0), tmpStr(0))
    ''                        End If
    ''                    Else
    ''                        If Not projektliste.Contains(tmpName) Then
    ''                            projektliste.Add(tmpName, tmpName)
    ''                        End If
    ''                    End If
    ''                Next

    ''            Else
    ''                'projektliste = aktuelleGesamtListe.getProjectNames
    ''                projektliste = constellation.getProjectNames()
    ''            End If

    ''            Dim showPname As Boolean



    ''            For Each pname In projektliste

    ''                showPname = True

    ''                ' im Falle activate Variante / Portfolio definieren: nur die Projekte anzeigen, die auch tatsächlich mehrere Varianten haben 
    ''                If aKtionskennung = PTTvActions.activateV Or aKtionskennung = PTTvActions.deleteV Then
    ''                    If constellation.getVariantZahl(pname) = 0 Then
    ''                        showPname = False
    ''                    End If
    ''                End If

    ''                If showPname Then

    ''                    Dim variantNames As Collection

    ''                    If quickList Then
    ''                        variantNames = getVariantListeFromPVNames(pvNamesList, pname)

    ''                    Else
    ''                        'variantNames = aktuelleGesamtListe.getVariantNames(pname, True)
    ''                        variantNames = constellation.getVariantNames(pname, True)
    ''                    End If

    ''                    nodeLevel0 = .Nodes.Add(pname)

    ''                    ' damit kann evtl direkt auf den Node zugegriffen werden ...
    ''                    nodeLevel0.Name = pname

    ''                    ' Berücksichtigung der Abhängigkeiten im TreeView ...
    ''                    If allDependencies.projectCount > 0 Then
    ''                        ' es gibt irgendwelche Dependencies, die Lead-Projekte, abhängigen Projekte 
    ''                        ' und sowohl-als-auch-Projekte werden farblich markiert  

    ''                        ' die Projekte suchen, von denen dieses Projekt abhängt 
    ''                        Dim passivListe As Collection = allDependencies.passiveListe(pname, PTdpndncyType.inhalt)
    ''                        Dim aktivListe As Collection = allDependencies.activeListe(pname, PTdpndncyType.inhalt)

    ''                        If passivListe.Count > 0 And aktivListe.Count = 0 Then
    ''                            ' ist nur abhängiges Projekt ...
    ''                            nodeLevel0.ForeColor = Color.Gray


    ''                        ElseIf passivListe.Count = 0 And aktivListe.Count > 0 Then
    ''                            ' hat abhängige Projekte  
    ''                            nodeLevel0.ForeColor = Color.OrangeRed

    ''                        ElseIf passivListe.Count > 0 And aktivListe.Count > 0 Then
    ''                            ' hängt ab und hat abhängige Projekte 
    ''                            nodeLevel0.ForeColor = Color.Orange
    ''                        End If

    ''                    End If


    ''                    ' Platzhalter einfügen; wird für alle Aktionskennungen benötigt

    ''                    If variantNames.Count > 1 Or _
    ''                        aKtionskennung = PTTvActions.delFromDB Then

    ''                        nodeLevel0.Tag = "X"
    ''                        For iv As Integer = 1 To variantNames.Count
    ''                            Dim tmpNodeLevel1 As TreeNode = nodeLevel0.Nodes.Add(CStr(variantNames.Item(iv)))
    ''                            If aKtionskennung = PTTvActions.delFromDB Then
    ''                                tmpNodeLevel1.Tag = "P"
    ''                                Dim tmpNodeLevel2 As TreeNode = tmpNodeLevel1.Nodes.Add("Platzhalter-Datum")
    ''                            Else
    ''                                tmpNodeLevel1.Tag = "X"
    ''                            End If

    ''                        Next

    ''                    Else
    ''                        nodeLevel0.Tag = "X"
    ''                    End If

    ''                    If aKtionskennung = PTTvActions.chgInSession Then
    ''                        If ShowProjekte.contains(pname) Then
    ''                            nodeLevel0.Checked = True

    ''                            ' jetzt die betreffende Variante setzen
    ''                            Dim hproj As clsProjekt = ShowProjekte.getProject(pname)
    ''                            Dim vName As String = "(" & hproj.variantName & ")"

    ''                            For Each tmpNode As TreeNode In nodeLevel0.Nodes
    ''                                If tmpNode.Text = vName Then
    ''                                    tmpNode.Checked = True
    ''                                Else
    ''                                    tmpNode.Checked = False
    ''                                End If
    ''                            Next


    ''                        End If
    ''                    ElseIf aKtionskennung = PTTvActions.delFromDB Then

    ''                        Dim vName As String
    ''                        If variantNames.Count > 1 Then

    ''                            Dim allWereReferenced As Boolean = True
    ''                            For Each tmpNode As TreeNode In nodeLevel0.Nodes

    ''                                Dim tmpStr() As String = tmpNode.Text.Split(New Char() {CChar("("), CChar(")")})
    ''                                vName = tmpStr(1)
    ''                                If notReferencedByAnyPortfolio(pname, vName) Then
    ''                                    ' alles ok 
    ''                                    allWereReferenced = False
    ''                                Else
    ''                                    tmpNode.ForeColor = Color.DimGray
    ''                                End If

    ''                            Next

    ''                            If allWereReferenced Then
    ''                                nodeLevel0.ForeColor = Color.DimGray
    ''                            End If

    ''                        Else
    ''                            Dim tmpStr() As String = CStr(variantNames.Item(1)).Split(New Char() {CChar("("), CChar(")")})
    ''                            vName = tmpStr(1)
    ''                            If notReferencedByAnyPortfolio(pname, vName) Then
    ''                                ' alles ok , kann gelöscht werden 
    ''                            Else
    ''                                nodeLevel0.ForeColor = Color.DimGray
    ''                            End If
    ''                        End If

    ''                    End If

    ''                End If

    ''            Next

    ''        End With
    ''    Else
    ''        Call MsgBox(loadErrorMsg)
    ''    End If


    ''End Sub

    

    

    ''' <summary>
    ''' wird hauptsächlich benötigt in Verbindung mit updateTreeView und frmProjPortfolioAdmin 
    ''' liefert eine Liste von Varianten-Namen, eingeschlossen in Klammern, die es zu Projekt pName gibt 
    ''' (), (v1), etc..
    ''' </summary>
    ''' <param name="pvNames"></param>
    ''' <param name="pName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getVariantListeFromPVNames(ByVal pvNames As SortedList(Of String, String), ByVal pName As String) As Collection
        Dim tmpResult As New Collection
        Dim vglName As String
        Dim variantName As String

        For Each kvp As KeyValuePair(Of String, String) In pvNames
            Dim tmpStr() As String = kvp.Key.Split(New Char() {CChar("#")})
            vglName = tmpStr(0)
            variantName = "()"

            If vglName = pName Then
                If tmpStr.Length = 1 Then
                    variantName = "()"
                ElseIf tmpStr.Length > 1 Then
                    variantName = "(" & tmpStr(1) & ")"
                End If

                If Not tmpResult.Contains(variantName) Then
                    tmpResult.Add(variantName, variantName)
                End If
            End If


        Next
        getVariantListeFromPVNames = tmpResult

    End Function

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
        Dim firstMilestone As Boolean = True
        Dim firstPhase As Boolean = True

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

                            If .isMilestone And firstMilestone Then
                                awinSettings.defaultMilestoneClass = .name
                                firstMilestone = False

                            ElseIf Not .isMilestone And firstPhase Then
                                awinSettings.defaultPhaseClass = .name
                                firstPhase = False
                            End If
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

    ''' <summary>
    ''' Prozedur um Username und Passwort für die Datenbank-Benutzung abzufragen und auch zu testen.
    ''' </summary>
    ''' <remarks></remarks>
    Function loginProzedur() As Boolean


        ' tk, 17.11.16 das wird nicht benötigt, rausgenommen, damit die 
        ' Login Prozedur auch von Powerpoint aus aufgerufen werden kann 
        ' appInstance.EnableEvents = False
        ' enableOnUpdate = False

        Dim loginDialog As New frmAuthentication
        Dim returnValue As DialogResult
        Dim i As Integer = 0

        returnValue = DialogResult.Retry

        ' ur: 30.6.2016: Login-Versuche auf fünf limitiert
        While returnValue = DialogResult.Retry And i < 5

            returnValue = loginDialog.ShowDialog
            i = i + 1

        End While

        If returnValue = DialogResult.Abort Or i >= 5 Then
            'Call MsgBox("Customization-File schließen")
            ' appInstance.EnableEvents = True
            ' enableOnUpdate = True
            Return False
        Else
            ' appInstance.EnableEvents = True
            ' enableOnUpdate = True
            Return True
        End If

    End Function

    ''' <summary>
    ''' übergebenene ProjektListe wird um die Projekte reduziert, die nicht zu dem Filter passen
    ''' das wird nur aufgerufen, wenn der Filter angewendet werden soll 
    ''' </summary>
    ''' <param name="projektListe"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function reduzierenWgFilter(ByVal projektListe As clsProjekteAlle) As clsProjekteAlle
        Dim filter As New clsFilter
        Dim ok As Boolean = False
        Dim newProjektliste As New clsProjekteAlle



        ' wenn applyFilter = true, dann soll  unter Anwendung 
        ' des Filters "Last" nachgeladen werden

        filter = filterDefinitions.retrieveFilter("Last")

        If IsNothing(filter) Then

            ' Liste unverändert zurückgeben
            reduzierenWgFilter = projektListe
        Else

            For Each kvp As KeyValuePair(Of String, clsProjekt) In projektListe.liste

                If Not filter.isEmpty Then
                    ok = filter.doesNotBlock(kvp.Value)
                Else
                    ok = True
                End If

                If ok Then
                    Try
                        newProjektliste.Add(kvp.Value, False)
                    Catch ex As Exception
                        Call MsgBox("Fehler in reduzierenWgFilter" & kvp.Key)
                    End Try
                Else

                End If

            Next

            ' Liste gefüllt mit Projekte, die auf den aktuellen Filter passen
            reduzierenWgFilter = newProjektliste
        End If

    End Function

    ''' <summary>
    ''' erstellt das Vorlagen File aus der Liste der Phasen 
    ''' aktuell wird nur die Übergabe von Phasen unterstützt
    ''' </summary>
    ''' <param name="phaseList"></param>
    ''' <param name="milestoneList"></param>
    ''' <remarks></remarks>
    Public Sub createVorlageFromSelection(ByVal phaseList As SortedList(Of String, String), _
                                              ByVal milestoneList As SortedList(Of String, String))

        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim elemName As String = ""
        Dim breadcrumb As String = ""
        Dim lfdNr As Integer = 1
        Dim fullName As String
        Dim ext As String = ""

        appInstance.EnableEvents = False
        enableOnUpdate = False


        ' hier muss jetzt das entsprechende File aufgemacht werden ...
        ' das File 
        Try
            appInstance.Workbooks.Add()


        Catch ex As Exception
            appInstance.EnableEvents = formerEE
            enableOnUpdate = True
            Throw New ArgumentException("Excel Export nicht gefunden - Abbruch")
        End Try

        'appInstance.Workbooks(myCustomizationFile).Activate()
        Dim wsName As Excel.Worksheet = CType(appInstance.ActiveSheet, _
                                                Global.Microsoft.Office.Interop.Excel.Worksheet)


        Dim zeile As Integer = 1
        Dim spalte As Integer = 1

        Dim startDate As Date, endDate As Date
        Dim tmpRange As Excel.Range
        Dim anzahlProjekte As Integer = ShowProjekte.Count

        With wsName
            ' jetzt werden alle Spalten auf Breite 25 gesetzt 
            tmpRange = CType(.Range(.Cells(zeile, spalte), .Cells(zeile, spalte).offset(0, 500)), Excel.Range)
            tmpRange.ColumnWidth = 25

            ' jetzt wird der Header geschrieben 
            CType(.Cells(zeile, spalte), Excel.Range).Value = "Produktlinie"
            CType(.Cells(zeile, spalte + 1), Excel.Range).Value = "Name"
            CType(.Cells(zeile, spalte + 2), Excel.Range).Value = "Projekt-Typ"
            CType(.Cells(zeile, spalte + 3), Excel.Range).Value = "ist abhängig von"
            CType(.Cells(zeile, spalte + 4), Excel.Range).Value = "strat. Bedeutung"
            CType(.Cells(zeile, spalte + 5), Excel.Range).Value = "Risiko der Umsetzung"
            CType(.Cells(zeile, spalte + 6), Excel.Range).Value = "Produktions-Volumen"
            CType(.Cells(zeile, spalte + 7), Excel.Range).Value = "Budget"


            spalte = spalte + 8


            ' hier muss noch korrigiert werden: wenn es bei einem oder mehreren Projekten mehrere Elemente dieses Namens und Breadcrumbs gibt, so 
            ' muss das in dieser Liste auch vorgesehen werden 
            For ix As Integer = 1 To phaseList.Count

                Dim phaseName As String = ""
                CType(.Cells(zeile, spalte), Excel.Range).Value = "Phasen-Name"
                CType(.Cells(zeile, spalte + 1), Excel.Range).Value = "Start-Datum"
                CType(.Cells(zeile, spalte + 2), Excel.Range).Value = "Ende-Datum"
                CType(.Cells(zeile, spalte + 3), Excel.Range).Value = "Skalierungs-Regel"
                CType(.Cells(zeile, spalte + 4), Excel.Range).Value = "Modul-Namen(n"

                tmpRange = CType(.Range(.Cells(zeile, spalte + 1), .Cells(zeile, spalte + 1).offset(anzahlProjekte, 1)), Excel.Range)
                tmpRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                tmpRange.NumberFormat = "dd/mm/yy;@"

                spalte = spalte + 5

            Next


        End With


        'es geht von vorne los 
        spalte = 1
        zeile = 2

        For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

            With wsName
                ' Produktlinie schreiben 

                Try
                    If kvp.Value.businessUnit.Length > 0 Then
                        CType(.Cells(zeile, spalte), Excel.Range).Value = kvp.Value.businessUnit
                    Else
                        CType(.Cells(zeile, spalte), Excel.Range).Value = "-"
                    End If
                Catch ex As Exception
                    CType(.Cells(zeile, spalte), Excel.Range).Value = "-"
                End Try


                ' Name schreiben 
                CType(.Cells(zeile, spalte + 1), Excel.Range).Value = kvp.Value.name

                ' Projekt-Typ schreiben 
                Try
                    If kvp.Value.VorlagenName.Length > 0 Then
                        CType(.Cells(zeile, spalte + 2), Excel.Range).Value = kvp.Value.VorlagenName
                    Else
                        CType(.Cells(zeile, spalte + 2), Excel.Range).Value = "-"
                    End If
                Catch ex As Exception
                    CType(.Cells(zeile, spalte + 2), Excel.Range).Value = "-"
                End Try

                ' ist abhängig von schreiben
                CType(.Cells(zeile, spalte + 3), Excel.Range).Value = ""


                ' strategische Bedeutung schreiben 
                CType(.Cells(zeile, spalte + 4), Excel.Range).Value = kvp.Value.StrategicFit

                ' risiko Kennzahl schreiben 
                CType(.Cells(zeile, spalte + 5), Excel.Range).Value = kvp.Value.Risiko


                ' Produktions-Volumen schreiben 
                CType(.Cells(zeile, spalte + 6), Excel.Range).Value = kvp.Value.volume

                ' Budget schreiben 
                CType(.Cells(zeile, spalte + 7), Excel.Range).Value = ""

                ' Phasen Information schreiben

                spalte = spalte + 8


                ' hier muss noch korrigiert werden: wenn es bei einem oder mehreren Projekten mehrere Elemente dieses Namens und Breadcrumbs gibt, so 
                ' muss das in dieser Liste auch vorgesehen werden 

                Dim cphase As clsPhase
                For ix As Integer = phaseList.Count To 1 Step -1

                    fullName = CStr(phaseList.ElementAt(ix - 1).Value)
                    elemName = ""
                    breadcrumb = ""
                    lfdNr = 0
                    Dim type As Integer = -1
                    Dim pvName As String = ""
                    Call splitBreadCrumbFullnameTo3(fullName, elemName, breadcrumb, lfdNr, type, pvName)


                    cphase = kvp.Value.getPhase(elemName, breadcrumb, lfdNr)
                    Dim phaseName As String

                    If Not IsNothing(cphase) Then
                        Try

                            phaseName = kvp.Value.getBestNameOfID(cphase.nameID, True, False)
                            startDate = cphase.getStartDate
                            endDate = cphase.getEndDate

                            CType(.Cells(zeile, spalte), Excel.Range).Value = phaseName.Replace("#", "-")
                            CType(.Cells(zeile, spalte + 1), Excel.Range).Value = startDate
                            CType(.Cells(zeile, spalte + 2), Excel.Range).Value = endDate
                            CType(.Cells(zeile, spalte + 3), Excel.Range).Value = "1"
                            CType(.Cells(zeile, spalte + 4), Excel.Range).Value = ""

                        Catch ex As Exception


                        End Try
                    Else

                        CType(.Cells(zeile, spalte), Excel.Range).Value = "-"


                    End If

                    spalte = spalte + 5

                Next


            End With

            zeile = zeile + 1
            spalte = 1

        Next

        'Dim expFName As String = awinPath & exportFilesOrdner & _
        '    "\Vorlage_" & Date.Now.ToString.Replace(":", ".") & ".xlsx"

        Dim expFName As String = exportOrdnerNames(PTImpExp.modulScen) & _
            "\Vorlage_" & Date.Now.ToString.Replace(":", ".") & ".xlsx"

        Try
            appInstance.ActiveWorkbook.SaveAs(Filename:=expFName, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges)
        Catch ex As Exception

        End Try

        Try
            appInstance.ActiveWorkbook.Close(SaveChanges:=False)
        Catch ex As Exception

        End Try

        appInstance.EnableEvents = True



    End Sub


    ''' <summary>
    ''' schreibt die übergebenen Phasen und Meilensteine in eine Excel Datei 
    ''' </summary>
    ''' <param name="phaseList"></param>
    ''' <param name="milestoneList"></param>
    ''' <remarks></remarks>
    Public Sub exportSelectionToExcel(ByVal phaseList As SortedList(Of String, String), _
                                            ByVal milestoneList As SortedList(Of String, String))

        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim elemName As String = ""
        Dim breadcrumb As String = ""
        Dim lfdNr As Integer = 1
        Dim fullName As String
        Dim ext As String = ""

        appInstance.EnableEvents = False
        enableOnUpdate = False


        ' hier muss jetzt das entsprechende File aufgemacht werden ...
        ' das File 
        Try
            'appInstance.Workbooks.Open(awinPath & requirementsOrdner & excelExportVorlage)
            appInstance.Workbooks.Add()


        Catch ex As Exception
            appInstance.EnableEvents = formerEE
            enableOnUpdate = True
            Throw New ArgumentException("Excel Export nicht gefunden - Abbruch")
        End Try

        'appInstance.Workbooks(myCustomizationFile).Activate()
        Dim wsName As Excel.Worksheet = CType(appInstance.ActiveSheet, _
                                                Global.Microsoft.Office.Interop.Excel.Worksheet)


        Dim zeile As Integer = 1
        Dim spalte As Integer = 1

        Dim startDate As Date, endDate As Date
        Dim earliestDate As Date, latestDate As Date
        Dim tmpRange As Excel.Range
        Dim anzahlProjekte As Integer = ShowProjekte.Count

        With wsName
            ' jetzt werden alle Spalten auf Breite 25 gesetzt 
            tmpRange = CType(.Range(.Cells(zeile, spalte), .Cells(zeile, spalte).offset(0, 200)), Excel.Range)
            tmpRange.ColumnWidth = 25

            ' jetzt wird der Header geschrieben 
            CType(.Cells(zeile, spalte), Excel.Range).Value = "Produktlinie"
            CType(.Cells(zeile, spalte + 1), Excel.Range).Value = "Name"
            CType(.Cells(zeile, spalte + 2), Excel.Range).Value = "Projekt-Typ"

            spalte = spalte + 2


            ' hier muss noch orrigiert werden: wenn es bei einem oder mehreren Projekten mehrere Elemente dieses Namens und Breadcrumbs gibt, so 
            ' muss das in dieser Liste auch vorgesehen werden 
            For ix As Integer = 1 To phaseList.Count

                Try
                    fullName = CStr(phaseList.ElementAt(ix - 1).Value)
                Catch ex As Exception
                    fullName = ""
                End Try

                Dim type As Integer = -1
                Dim pvName As String = ""
                Call splitBreadCrumbFullnameTo3(fullName, elemName, breadcrumb, lfdNr, type, pvName)

                If lfdNr > 1 Then
                    ext = " " & lfdNr.ToString
                Else
                    ext = ""
                End If
                If breadcrumb = "" Then
                    CType(.Cells(zeile, spalte + ix), Excel.Range).Value = elemName & ext
                Else
                    CType(.Cells(zeile, spalte + ix), Excel.Range).Value = breadcrumb.Replace("#", "-") & "-" & elemName & ext
                End If

            Next

            spalte = spalte + phaseList.Count

            ' hier muss noch orrigiert werden: wenn es bei einem oder mehreren Projekten mehrere Elemente dieses NAmens gibt, so 
            ' muss das in dieser Liste auch vorgesehen werden 

            For ix As Integer = 1 To milestoneList.Count

                Try
                    fullName = CStr(milestoneList.ElementAt(ix - 1).Value)
                Catch ex As Exception
                    fullName = ""
                End Try

                Dim type As Integer = -1
                Dim pvName As String = ""
                Call splitBreadCrumbFullnameTo3(fullName, elemName, breadcrumb, lfdNr, type, pvName)

                If lfdNr > 1 Then
                    ext = " " & lfdNr.ToString
                Else
                    ext = ""
                End If
                If breadcrumb = "" Then
                    CType(.Cells(zeile, spalte + ix), Excel.Range).Value = elemName & ext
                Else
                    CType(.Cells(zeile, spalte + ix), Excel.Range).Value = breadcrumb.Replace("#", "-") & "-" & elemName & ext
                End If


            Next


            ' Datumsformat einstellen 
            Dim s1 As Integer = 4 + phaseList.Count
            Dim o1 As Integer = milestoneList.Count - 1

            ' mittig darstellen 
            tmpRange = CType(.Range(.Cells(zeile, 4), .Cells(zeile, 4).offset(anzahlProjekte, s1 + o1 - 4)), Excel.Range)
            tmpRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

            tmpRange = CType(.Range(.Cells(zeile + 1, s1), .Cells(zeile + 1, s1).offset(anzahlProjekte - 1, o1)), Excel.Range)
            tmpRange.NumberFormat = "dd/mm/yy;@"

            spalte = spalte + milestoneList.Count

            CType(.Cells(zeile, spalte + 1), Excel.Range).Value = "Dauer (T)"
            CType(.Range(.Cells(zeile + 1, spalte + 1), .Cells(zeile + 1, spalte + 1).offset(anzahlProjekte - 1, 0)), Excel.Range).NumberFormat = "0"

            CType(.Cells(zeile, spalte + 2), Excel.Range).Value = "Dauer (M)"
            CType(.Range(.Cells(zeile + 1, spalte + 2), .Cells(zeile + 1, spalte + 2).offset(anzahlProjekte - 1, 0)), Excel.Range).NumberFormat = "0.0"

        End With


        zeile = 2
        spalte = 1
        Dim minCol As Integer
        Dim maxCol As Integer

        For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

            earliestDate = kvp.Value.endeDate
            latestDate = kvp.Value.startDate

            ' wird benötigt, um festzustellen, ob überhaupt eines der Elemente im aktuell 
            ' betrachteten Projekt vorkommt 
            Dim atleastOne As Boolean = False

            With wsName
                ' Produktlinie schreiben 

                Try
                    If kvp.Value.businessUnit.Length > 0 Then
                        CType(.Cells(zeile, spalte), Excel.Range).Value = kvp.Value.businessUnit
                    Else
                        CType(.Cells(zeile, spalte), Excel.Range).Value = "-"
                    End If
                Catch ex As Exception
                    CType(.Cells(zeile, spalte), Excel.Range).Value = "-"
                End Try


                ' Name schreiben 
                CType(.Cells(zeile, spalte + 1), Excel.Range).Value = kvp.Value.name

                ' Projekt-Typ schreiben 
                Try
                    If kvp.Value.VorlagenName.Length > 0 Then
                        CType(.Cells(zeile, spalte + 2), Excel.Range).Value = kvp.Value.VorlagenName
                    Else
                        CType(.Cells(zeile, spalte + 2), Excel.Range).Value = "-"
                    End If
                Catch ex As Exception
                    CType(.Cells(zeile, spalte + 2), Excel.Range).Value = "-"
                End Try


                ' Phasen Information schreiben

                Dim cphase As clsPhase
                spalte = spalte + 3

                For ix As Integer = 1 To phaseList.Count

                    fullName = CStr(phaseList.ElementAt(ix - 1).Value)
                    elemName = ""
                    breadcrumb = ""
                    lfdNr = 0
                    Dim type As Integer = -1
                    Dim pvName As String = ""
                    Call splitBreadCrumbFullnameTo3(fullName, elemName, breadcrumb, lfdNr, type, pvName)


                    cphase = kvp.Value.getPhase(elemName, breadcrumb, lfdNr)

                    If Not IsNothing(cphase) Then
                        Try
                            startDate = cphase.getStartDate
                            endDate = cphase.getEndDate

                            atleastOne = True

                            If DateDiff(DateInterval.Day, startDate, earliestDate) > 0 Then
                                earliestDate = startDate
                                minCol = spalte
                            End If

                            If DateDiff(DateInterval.Day, latestDate, endDate) > 0 Then
                                latestDate = endDate
                                maxCol = spalte
                            End If

                            CType(.Cells(zeile, spalte), Excel.Range).Value = startDate.ToShortDateString & " - " & endDate.ToShortDateString

                        Catch ex As Exception
                            CType(.Cells(zeile, spalte), Excel.Range).Value = "?"

                        End Try
                    Else

                        CType(.Cells(zeile, spalte), Excel.Range).Value = "-"


                    End If

                    spalte = spalte + 1



                Next


                ' Meilensteine schreiben 

                Dim milestone As clsMeilenstein = Nothing

                For ix As Integer = 1 To milestoneList.Count

                    fullName = CStr(milestoneList.ElementAt(ix - 1).Value)
                    elemName = ""
                    breadcrumb = ""
                    lfdNr = 0
                    Dim type As Integer = -1
                    Dim pvName As String = ""
                    Call splitBreadCrumbFullnameTo3(fullName, elemName, breadcrumb, lfdNr, type, pvName)

                    milestone = kvp.Value.getMilestone(elemName, breadcrumb, lfdNr)

                    If Not IsNothing(milestone) Then
                        Try
                            startDate = milestone.getDate

                            atleastOne = True

                            If DateDiff(DateInterval.Day, startDate, earliestDate) > 0 Then
                                earliestDate = startDate
                                minCol = spalte
                            End If

                            If DateDiff(DateInterval.Day, latestDate, startDate) > 0 Then
                                latestDate = startDate
                                maxCol = spalte
                            End If

                            CType(.Cells(zeile, spalte), Excel.Range).Value = startDate


                        Catch ex As Exception
                            CType(.Cells(zeile, spalte), Excel.Range).Value = "?"
                            CType(.Cells(zeile, spalte), Excel.Range).Value = "?"
                        End Try
                    Else

                        CType(.Cells(zeile, spalte), Excel.Range).Value = "-"
                        CType(.Cells(zeile, spalte), Excel.Range).Value = "-"

                    End If

                    spalte = spalte + 1

                Next

                Dim dauerT As Long
                Dim dauerM As Double

                ' Dauer in Tagen schreiben 

                Try
                    If atleastOne Then
                        dauerT = DateDiff(DateInterval.Day, earliestDate, latestDate)
                        dauerM = 12 * dauerT / 365
                    Else
                        dauerT = 0
                        dauerM = 0.0
                    End If
                Catch ex As Exception
                    dauerT = 0
                    dauerM = 0.0
                End Try


                CType(.Cells(zeile, spalte), Excel.Range).Value = dauerT
                CType(.Cells(zeile, spalte + 1), Excel.Range).Value = dauerM

                ' jetzt einfärben, welche Daten zu der Dauer geführt haben 
                If minCol = maxCol And minCol > 0 Then
                    CType(.Cells(zeile, minCol), Excel.Range).Interior.Color = awinSettings.AmpelGruen
                Else
                    If minCol > 0 Then
                        CType(.Cells(zeile, minCol), Excel.Range).Interior.Color = awinSettings.AmpelNichtBewertet
                    End If
                    If maxCol > 0 Then
                        CType(.Cells(zeile, maxCol), Excel.Range).Interior.Color = awinSettings.AmpelGelb
                    End If

                End If



            End With

            zeile = zeile + 1
            spalte = 1

        Next

        'Dim expFName As String = awinPath & exportFilesOrdner & _
        '    "\Report_" & Date.Now.ToString.Replace(":", ".") & ".xlsx"

        Dim expFName As String = exportOrdnerNames(PTImpExp.rplan) & _
            "\Report_" & Date.Now.ToString.Replace(":", ".") & ".xlsx"

        Try
            appInstance.ActiveWorkbook.SaveAs(Filename:=expFName, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges)
        Catch ex As Exception

        End Try

        Try
            appInstance.ActiveWorkbook.Close(SaveChanges:=False)
        Catch ex As Exception

        End Try

        appInstance.EnableEvents = True



    End Sub
    ''' <summary>
    ''' erstellt die Vorlage für die InputDatei des Batch-Report
    ''' Input-Tabelle wird erzeugt, wie vom VISBO ReportGen erwartet
    ''' ReportProfile - Tabelle wird bestückt aus den vorhandenen ReportProfilen in Directory ReportProfile
    ''' ProjekteSzenarien - Tabelle wird bestückt aus Liste AlleProjekte (d.h. es müssen Projekte oder Szenarien geladen sein
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub createReportGenTemplate()

        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim tmpRange As Excel.Range

        Dim zeile As Integer = 1
        Dim spalte As Integer = 1

        appInstance.EnableEvents = False
        enableOnUpdate = False


        ' hier muss jetzt das entsprechende File aufgemacht werden ...
        ' das File 
        Try
            'appInstance.Workbooks.Open(awinPath & requirementsOrdner & excelExportVorlage)
            appInstance.Workbooks.Add()


        Catch ex As Exception
            appInstance.EnableEvents = formerEE
            enableOnUpdate = True
            Throw New ArgumentException("Excel Export nicht gefunden - Abbruch")
        End Try



        Dim wsName As Excel.Worksheet
        appInstance.Worksheets.Add()
        wsName = CType(appInstance.ActiveSheet, _
                                                Global.Microsoft.Office.Interop.Excel.Worksheet)
        wsName.Name = "ProjekteSzenarien"

        zeile = 1
        spalte = 1


        Dim anzahlProjekte As Integer = AlleProjekte.Count

        With wsName
            ' jetzt werden alle Spalten auf Breite 25 gesetzt 
            tmpRange = CType(.Range(.Cells(zeile, spalte), .Cells(zeile, spalte).offset(0, 200)), Excel.Range)
            tmpRange.ColumnWidth = 25


            ' jetzt wird der Header geschrieben 
            With CType(.Cells(zeile, spalte), Excel.Range)
                .Value = "Projekte "
                With .Font
                    .Name = "Arial"
                    .FontStyle = "Fett"
                    .Size = 11
                    .Strikethrough = False
                    .Superscript = False
                    .Subscript = False
                    .OutlineFont = False
                    .Shadow = False
                End With
            End With

            With CType(.Cells(zeile, spalte + 1), Excel.Range)
                .Value = "Varianten"
                With .Font
                    .Name = "Arial"
                    .FontStyle = "Fett"
                    .Size = 11
                    .Strikethrough = False
                    .Superscript = False
                    .Subscript = False
                    .OutlineFont = False
                    .Shadow = False
                End With
            End With

            spalte = spalte + 1
        End With


        zeile = 2
        spalte = 1

        For Each kvp As KeyValuePair(Of String, clsProjekt) In AlleProjekte.liste

            Dim projName As String = kvp.Value.name
            Dim variantName As String = kvp.Value.variantName

            With wsName


                ' Name schreiben 
                CType(.Cells(zeile, spalte), Excel.Range).Value = kvp.Value.name

                ' Varianten-Name schreiben 
                CType(.Cells(zeile, spalte + 1), Excel.Range).Value = kvp.Value.variantName



            End With

            zeile = zeile + 1
            spalte = 1

        Next


        zeile = zeile + 1   ' eine Leerzeile
        spalte = 1
        With wsName
            With CType(.Cells(zeile, spalte), Excel.Range)
                .Value = "Szenarien"
                With .Font
                    .Name = "Arial"
                    .FontStyle = "Fett"
                    .Size = 11
                    .Strikethrough = False
                    .Superscript = False
                    .Subscript = False
                    .OutlineFont = False
                    .Shadow = False
                End With
            End With
        End With

        zeile = zeile + 1   ' eine Leerzeile
        spalte = 1

        ' alle möglichen Szenario-Namen eintragen
        For Each kvp As KeyValuePair(Of String, clsConstellation) In projectConstellations.Liste

            Dim szenarioName As String = kvp.Value.constellationName

            With wsName


                ' SzenarioName schreiben 
                CType(.Cells(zeile, spalte), Excel.Range).Value = kvp.Value.constellationName

            End With

            zeile = zeile + 1
            spalte = 1

        Next

        Dim wsReportProfile As Excel.Worksheet
        appInstance.Worksheets.Add()
        wsReportProfile = CType(appInstance.ActiveSheet, _
                                              Global.Microsoft.Office.Interop.Excel.Worksheet)
        wsReportProfile.Name = "ReportProfile"

        zeile = 1
        spalte = 1

        With wsReportProfile

            ' jetzt wird der Header geschrieben 
            With CType(.Cells(zeile, spalte), Excel.Range)
                .ColumnWidth = 40
                .Value = "ReportProfile"
                With .Font
                    .Name = "Arial"
                    .FontStyle = "Fett"
                    .Size = 11
                    .Strikethrough = False
                    .Superscript = False
                    .Subscript = False
                    .OutlineFont = False
                    .Shadow = False
                End With
            End With

        End With

        zeile = 2
        spalte = 1

        Dim dateiName As String = ""

        Try

            With wsReportProfile

                Dim i As Integer
                Dim dirname As String = My.Computer.FileSystem.CombinePath(awinPath, ReportProfileOrdner)

                ' ReportProfile vom Directory lesen
                Dim listOfVorlagen As Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Computer.FileSystem.GetFiles(dirname)

                ' und in das Excel-File eintragen
                For i = 1 To listOfVorlagen.Count
                    Dim tmpstr() As String = Split(Dir(listOfVorlagen.Item(i - 1)), ".xml")
                    dateiName = tmpstr(0)
                    CType(.Cells(zeile, spalte), Excel.Range).Value = dateiName
                    zeile = zeile + 1

                Next i

            End With
        Catch ex As Exception

        End Try

        Dim wsInput As Excel.Worksheet
        appInstance.Worksheets.Add()
        wsInput = CType(appInstance.ActiveSheet, _
                                              Global.Microsoft.Office.Interop.Excel.Worksheet)
        wsInput.Name = "Input"

        zeile = 1
        spalte = 1

        With wsInput
            ' jetzt werden alle Spalten auf Breite 40 gesetzt 
            tmpRange = CType(.Range(.Cells(zeile, spalte), .Cells(zeile, spalte).offset(0, 200)), Excel.Range)
            With tmpRange
                .RowHeight = 20
                .HorizontalAlignment = XlHAlign.xlHAlignCenter
                .VerticalAlignment = XlVAlign.xlVAlignCenter

                With .Font
                    .Name = "Arial"
                    .FontStyle = "Fett"
                    .Size = 11
                    .Strikethrough = False
                    .Superscript = False
                    .Subscript = False
                    .OutlineFont = False
                    .Shadow = False
                End With
            End With

            ' jetzt wird der Header geschrieben 

            With CType(.Cells(zeile, spalte), Excel.Range)
                .Value = "Name des Reports"
                .ColumnWidth = 40
            End With

            With CType(.Cells(zeile, spalte + 1), Excel.Range)
                .Value = "SpeicherModus"
                .ColumnWidth = 15
            End With

            With CType(.Cells(zeile, spalte + 2), Excel.Range)
                .Value = "Name des ReportProfils"
                .ColumnWidth = 45
            End With

            With CType(.Cells(zeile, spalte + 3), Excel.Range)
                .Value = "Names des Portfolios / Projekts"
                .ColumnWidth = 30
            End With

            With CType(.Cells(zeile, spalte + 4), Excel.Range)
                .Value = "VariantenName"
                .ColumnWidth = 30
            End With

            With CType(.Cells(zeile, spalte + 5), Excel.Range)
                .Value = "TimeStamp"
                .ColumnWidth = 30
            End With

            With CType(.Cells(zeile, spalte + 6), Excel.Range)
                .Value = " von"
                .ColumnWidth = 18
            End With

            With CType(.Cells(zeile, spalte + 7), Excel.Range)
                .Value = "bis"
                .ColumnWidth = 18
            End With

        End With



        'Dim expFName As String = awinPath & exportFilesOrdner & _
        '    "\Report_" & Date.Now.ToString.Replace(":", ".") & ".xlsx"


        Dim expFName As String = exportOrdnerNames(PTImpExp.modulScen) & _
            "\ReportGenTemplate_" & Date.Now.ToString.Replace(":", ".") & ".xlsx"

        Try
            appInstance.ActiveWorkbook.SaveAs(Filename:=expFName, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges)
        Catch ex As Exception

        End Try

        Try
            appInstance.ActiveWorkbook.Close(SaveChanges:=False)
        Catch ex As Exception

        End Try

        appInstance.EnableEvents = True



    End Sub

    ''' <summary>
    ''' ruft das Formular auf, um Filter zu definieren
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub defineFilterDB()
        Dim auswahlFormular As New frmNameSelection
        Dim returnValue As DialogResult

        With auswahlFormular

            '.showModePortfolio = True
            .menuOption = PTmenue.filterdefinieren

            '.Show()
            returnValue = .ShowDialog
        End With

    End Sub


    ''' <summary>
    ''' zeichnet das Leistbarkeits-Chart 
    ''' </summary>
    ''' <param name="selCollection">Collection mit den Phasne-, Meilenstein, Rollen- oder Kostenarten</param>
    ''' <param name="chTyp">Typ: es handelt sich um Phasen, rollen, etc. </param>
    ''' <param name="chtop">auf welcher Höhe soll das Chart gezeichnet werden</param>
    ''' <param name="chleft">auf welcher x-Koordinate soll das Chart gezeichnet werden</param>
    ''' <remarks></remarks>
    Friend Sub zeichneLeistbarkeitsChart(ByVal selCollection As Collection, ByVal chTyp As String, ByVal oneChart As Boolean, _
                                              ByRef chtop As Double, ByRef chleft As Double, ByVal chwidth As Double, ByVal chHeight As Double)


        Dim repObj As Excel.ChartObject
        Dim myCollection As Collection


        '' Window Position festlegen 
        'chHeight = maxScreenHeight / 4 - 3
        'chWidth = maxScreenWidth / 5 - 3

        'chWidth = 265 + (showRangeRight - showRangeLeft - 12 + 1) * boxWidth + (showRangeRight - showRangeLeft) * screen_correct
        'chHeight = awinSettings.ChartHoehe1


        If oneChart = True Then


            ' alles in einem Chart anzeigen
            myCollection = New Collection
            For Each element As String In selCollection
                myCollection.Add(element, element)
            Next

            repObj = Nothing
            Call awinCreateprcCollectionDiagram(myCollection, repObj, chtop, chleft,
                                                              chWidth, chHeight, False, chTyp, False)


            'chtop = chtop + 7 + chHeight
            chtop = chtop + 2 + chHeight
            'chleft = chleft + 7
        Else
            ' für jedes ITEM ein eigenes Chart machen
            For Each element As String In selCollection
                ' es muss jedesmal eine neue Collection erzeugt werden - die Collection wird in DiagramList gemerkt
                ' wenn die mit Clear leer gemacht wird, funktioniert der Diagram Update nicht mehr ....
                myCollection = New Collection
                myCollection.Add(element, element)
                repObj = Nothing

                Call awinCreateprcCollectionDiagram(myCollection, repObj, chtop, chleft,
                                                                   chWidth, chHeight, False, chTyp, False)

                'chtop = chtop + 5
                'chleft = chleft + 7

                chtop = chtop + 2 + chHeight
            Next

        End If

    End Sub

    ''' <summary>
    ''' wird aus Formular NameSelection bzw. HrySelection aufgerufen
    ''' besetzt die Vorlagen Dropbox den entsprechenden Datei-NAmen
    ''' </summary>
    ''' <param name="menuOption"></param>
    ''' <param name="repVorlagenDropbox"></param>
    ''' <remarks></remarks>
    Public Sub frmHryNameReadPPTVorlagen(ByVal menuOption As Integer, ByRef repVorlagenDropbox As System.Windows.Forms.ComboBox, Optional ByVal mppreport As Boolean = False)


        Dim dirname As String
        Dim dateiName As String = ""


        If menuOption = PTmenue.multiprojektReport Or menuOption = PTmenue.einzelprojektReport Then

            If menuOption = PTmenue.multiprojektReport Then
                dirname = awinPath & RepPortfolioVorOrdner
            Else
                dirname = awinPath & RepProjectVorOrdner
            End If

            Dim listOfVorlagen As Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Computer.FileSystem.GetFiles(dirname)
            Try

                Dim i As Integer
                For i = 1 To listOfVorlagen.Count
                    dateiName = Dir(listOfVorlagen.Item(i - 1))
                    If dateiName.Contains("Typ II") Then
                        repVorlagenDropbox.Items.Add(dateiName)
                    End If

                Next i
            Catch ex As Exception

            End Try
        ElseIf menuOption = PTmenue.reportBHTC Or _
            menuOption = PTmenue.reportMultiprojektTafel Then

            If mppreport Then
                dirname = awinPath & RepPortfolioVorOrdner
            Else
                dirname = awinPath & RepProjectVorOrdner
            End If


            Dim listOfVorlagen As Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Computer.FileSystem.GetFiles(dirname)
            Try

                Dim i As Integer
                For i = 1 To listOfVorlagen.Count

                    dateiName = Dir(listOfVorlagen.Item(i - 1))
                    repVorlagenDropbox.Items.Add(dateiName)

                Next i
            Catch ex As Exception

            End Try

        End If

    End Sub
    ''' <summary>
    ''' wird aus Formular NameSelection bzw. HrySelection aufgerufen
    ''' besetzt die Filter-Auswahl Dropbox mit Filternamen aus Datenbank
    ''' </summary>
    ''' <param name="menuOption"></param>
    ''' <param name="filterDropbox"></param>
    ''' <remarks></remarks>
    Public Sub frmHryNameReadFilterVorlagen(ByVal menuOption As Integer, ByRef filterDropbox As System.Windows.Forms.ComboBox)


        ' einlesen und anzeigen der in der Datenbank definierten Filter
        If menuOption = PTmenue.filterdefinieren Then

            If Not noDB Then
                ' Filter mit Namen "fName" in DB speichern
                Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)


                ' Datenbank ist gestartet
                If request.pingMongoDb() Then

                    Dim listofDBFilter As SortedList(Of String, clsFilter) = request.retrieveAllFilterFromDB(False)
                    For Each kvp As KeyValuePair(Of String, clsFilter) In listofDBFilter
                        If Not filterDefinitions.Liste.ContainsKey(kvp.Key) Then
                            filterDefinitions.Liste.Add(kvp.Key, kvp.Value)
                        End If
                    Next
                Else
                    Call MsgBox(" Datenbank-Verbindung ist unterbrochen!" & vbLf & " Filter kann nicht in DB gespeichert werden")
                End If
            Else

            End If

        Else
            If menuOption = PTmenue.visualisieren Or _
                menuOption = PTmenue.multiprojektReport Or _
                menuOption = PTmenue.einzelprojektReport Or _
                menuOption = PTmenue.leistbarkeitsAnalyse Then

                If Not noDB Then

                    ' allee Filter aus DB lesen
                    Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)

                    ' Datenbank ist gestartet
                    If request.pingMongoDb() Then

                        Dim listofDBFilter As SortedList(Of String, clsFilter) = request.retrieveAllFilterFromDB(True)
                        For Each kvp As KeyValuePair(Of String, clsFilter) In listofDBFilter

                            If Not selFilterDefinitions.Liste.ContainsKey(kvp.Key) Then
                                selFilterDefinitions.Liste.Add(kvp.Key, kvp.Value)
                            End If

                        Next
                    Else
                        Call MsgBox(" Datenbank-Verbindung ist unterbrochen!" & vbLf & " Filter kann nicht in DB gespeichert werden")
                    End If

                End If
            End If
        End If


    End Sub

    ''' <summary>
    ''' führt die Aktionen Visualisieren, Leistbarkeit, Meilenstein Trendanalyse aus dem Hierarchie bzw. Namen-Auswahl Fenster durch 
    ''' 
    ''' </summary>
    ''' <param name="menueOption"></param>
    ''' <remarks></remarks>
    Public Sub frmHryNameActions(ByVal menueOption As Integer, _
                                 ByVal selectedPhases As Collection, ByVal selectedMilestones As Collection, _
                                 ByVal selectedRoles As Collection, ByVal selectedCosts As Collection, _
                                 ByVal oneChart As Boolean, ByVal filtername As String)

        Dim chTyp As String
        Dim validOption As Boolean

        If menueOption = PTmenue.visualisieren Or menueOption = PTmenue.einzelprojektReport Or _
            menueOption = PTmenue.excelExport Or menueOption = PTmenue.multiprojektReport Or _
            menueOption = PTmenue.vorlageErstellen Or menueOption = PTmenue.meilensteinTrendanalyse Then
            validOption = True
        ElseIf showRangeRight - showRangeLeft >= minColumns - 1 Then
            validOption = True
        Else
            validOption = False
        End If

        If menueOption = PTmenue.leistbarkeitsAnalyse Then

            Dim myCollection As New Collection

            If (selectedPhases.Count > 0 Or selectedMilestones.Count > 0 _
                    Or selectedRoles.Count > 0 Or selectedCosts.Count > 0) _
                    And validOption Then

                Dim formerSU As Boolean = appInstance.ScreenUpdating
                appInstance.ScreenUpdating = False

                ' Window Position festlegen
                'Dim chtop As Double = 50.0 + awinSettings.ChartHoehe1
                'Dim chleft As Double = (showRangeRight - 1) * boxWidth + 4
                Dim chtop As Double
                Dim chleft As Double
                Dim chwidth As Double
                Dim chHeight As Double


                'If visboZustaende.projectBoardMode = ptModus.graficboard Then
                '    chleft = (showRangeRight - 1) * boxWidth + 4
                'Else
                '    chleft = 5
                'End If

                '' um es im neuen Portfolio Chart Window anzuzeigen ... 
                'chtop = 3
                'chleft = 3

                Call bestimmeChartPositionAndSize(ptTables.mptPfCharts, chtop, chleft, chwidth, chHeight)


                If selectedPhases.Count > 0 Then
                    chTyp = DiagrammTypen(0)
                    Call zeichneLeistbarkeitsChart(selectedPhases, chTyp, oneChart, _
                                                   chtop, chleft, chwidth, chheight)
                End If

                If selectedMilestones.Count > 0 Then
                    chTyp = DiagrammTypen(5)
                    Call zeichneLeistbarkeitsChart(selectedMilestones, chTyp, oneChart, _
                                                   chtop, chleft, chwidth, chHeight)
                End If

                If selectedRoles.Count > 0 Then
                    chTyp = DiagrammTypen(1)
                    Call zeichneLeistbarkeitsChart(selectedRoles, chTyp, oneChart, _
                                                   chtop, chleft, chwidth, chHeight)
                End If

                If selectedCosts.Count > 0 Then
                    chTyp = DiagrammTypen(2)
                    Call zeichneLeistbarkeitsChart(selectedCosts, chTyp, oneChart, _
                                                   chtop, chleft, chwidth, chHeight)
                End If


                appInstance.ScreenUpdating = formerSU

            Else

            End If

        ElseIf menueOption = PTmenue.visualisieren Then


            If (selectedPhases.Count > 0 Or selectedMilestones.Count > 0 _
                    Or selectedRoles.Count > 0 Or selectedCosts.Count > 0) _
                    And validOption Then

                If (selectedPhases.Count > 0 Or selectedMilestones.Count > 0) And _
                    (selectedRoles.Count > 0 Or selectedCosts.Count > 0) Then
                    Call MsgBox("es können nur entweder Phasen / Meilensteine oder Rollen oder Kosten angezeigt werden")

                ElseIf selectedPhases.Count > 0 Or selectedMilestones.Count > 0 Then

                    If selectedPhases.Count > 0 Then
                        Call deleteBeschriftungen()
                        If roentgenBlick.isOn Then
                            Call awinNoshowProjectNeeds()
                            roentgenBlick.isOn = False
                        End If
                        Call awinZeichnePhasen(selectedPhases, False, True)

                        ' Selektion der selektierten Projekte wieder sichtbar machen
                        If selectedProjekte.Count > 0 Then
                            Call awinSelect()
                        End If
                    End If

                    If selectedMilestones.Count > 0 Then
                        ' Phasen anzeigen 
                        Dim farbID As Integer = 4
                        Call deleteBeschriftungen()
                        If roentgenBlick.isOn Then
                            Call awinNoshowProjectNeeds()
                            roentgenBlick.isOn = False
                        End If
                        Call awinZeichneMilestones(selectedMilestones, farbID, False, True)

                    End If

                ElseIf selectedRoles.Count > 0 Then

                    Call awinDeleteProjectChildShapes(0)
                    Call deleteBeschriftungen()
                    Call awinZeichneBedarfe(selectedRoles, DiagrammTypen(1))

                ElseIf selectedCosts.Count > 0 Then

                    Call awinDeleteProjectChildShapes(0)
                    Call deleteBeschriftungen()
                    Call awinZeichneBedarfe(selectedCosts, DiagrammTypen(2))

                Else
                    Call MsgBox("noch nicht implementiert")
                End If

            Else
                Call MsgBox("bitte mindestens ein Element aus einer der Kategorien selektieren  ")
            End If

            ' selektierte Projekte weiterhin als selektiert darstellen
            If selectedProjekte.Count > 0 Then
                Call awinSelect()
            End If

        ElseIf menueOption = PTmenue.filterdefinieren Then

            'Call MsgBox("ok, Filter gespeichert")

        ElseIf menueOption = PTmenue.sessionFilterDefinieren Then
            ' keine Message ausgeben ...

        ElseIf menueOption = PTmenue.excelExport Or menueOption = PTmenue.vorlageErstellen Then

            If (selectedPhases.Count > 0 Or selectedMilestones.Count > 0) _
                    And validOption Then

                Try
                    Call createDateiFromSelection(filtername, menueOption)
                    If menueOption = PTmenue.excelExport Then
                        Call MsgBox("ok, Excel File in " & exportOrdnerNames(PTImpExp.rplan) & " erzeugt")
                    Else
                        Call MsgBox("ok, Excel File in " & exportOrdnerNames(PTImpExp.modulScen) & " erzeugt")
                    End If

                Catch ex As Exception
                    Call MsgBox(ex.Message)
                End Try

            Else
                Call MsgBox("bitte mindestens ein Element aus einer der Kategorien Phasen / Meilensteine selektieren  ")
            End If
        ElseIf menueOption = PTmenue.meilensteinTrendanalyse Then


            If selectedMilestones.Count > 0 Then
                ' Window Position festlegen

                Call awinShowMilestoneTrend(selectedMilestones)
            Else
                Call MsgBox("Bitte Meilensteine auswählen! ")

            End If

        Else

            Call MsgBox("noch nicht unterstützt")

        End If

    End Sub


    Sub awinShowMilestoneTrend(ByVal selectedMilestones As Collection)

        Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
        Dim singleShp As Excel.Shape
        Dim listOfItems As New Collection
        Dim nameList As New SortedList(Of Date, String)
        Dim title As String = "Meilensteine auswählen"
        Dim hproj As clsProjekt
        Dim awinSelection As Excel.ShapeRange
        Dim selektierteProjekte As New clsProjekte
        Dim top As Double, left As Double, height As Double, width As Double
        Dim repObj As Excel.ChartObject = Nothing

        Dim pName As String, vglName As String = " "
        Dim variantName As String

        Call projektTafelInit()

        Try
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If request.pingMongoDb() Then

            If Not awinSelection Is Nothing Then

                ' eingangs-prüfung, ob auch nur ein Element selektiert wurde ...
                If awinSelection.Count = 1 Then

                    ' Aktion durchführen ...

                    singleShp = awinSelection.Item(1)

                    Try

                        hproj = ShowProjekte.getProject(singleShp.Name, True)
                        nameList = hproj.getMilestones
                        listOfItems = hproj.getElemIdsOf(selectedMilestones, True)


                        ' jetzt muss die ProjektHistorie aufgebaut werden 
                        With hproj
                            pName = .name
                            variantName = .variantName
                        End With

                        If Not projekthistorie Is Nothing Then
                            If projekthistorie.Count > 0 Then
                                vglName = projekthistorie.First.getShapeText
                            End If
                        Else
                            projekthistorie = New clsProjektHistorie
                        End If

                        If vglName <> hproj.getShapeText Then

                            ' projekthistorie muss nur dann neu bestimmt werden, wenn sie nicht bereits für dieses Projekt geholt wurde
                            projekthistorie.liste = request.retrieveProjectHistoryFromDB(projectname:=pName, variantName:=variantName, _
                                                                                storedEarliest:=StartofCalendar, storedLatest:=Date.Now)
                            projekthistorie.Add(Date.Now, hproj)


                        Else
                            ' der aktuelle Stand hproj muss hinzugefügt werden 
                            Dim lastElem As Integer = projekthistorie.Count - 1
                            projekthistorie.RemoveAt(lastElem)
                            projekthistorie.Add(Date.Now, hproj)
                        End If



                        With singleShp
                            top = .Top + boxHeight + 5
                            left = .Left - 5
                        End With

                        height = 2 * ((nameList.Count - 1) * 20 + 110)
                        width = System.Math.Max(hproj.anzahlRasterElemente * boxWidth + 10, 24 * boxWidth + 10)


                        Call createMsTrendAnalysisOfProject(hproj, repObj, listOfItems, top, left, height, width)


                    Catch ex As Exception
                        Call MsgBox(ex.Message)
                    End Try

                Else
                    Call MsgBox("bitte nur ein Projekt selektieren ...")
                End If
            Else
                Call MsgBox("vorher ein Projekt selektieren ...")
            End If

        Else
            Call MsgBox(" Datenbank-Verbindung ist unterbrochen!" & vbLf & " Projekthistorie kann nicht geladen werden")
            'projekthistorie.clear()
        End If
        enableOnUpdate = True
        appInstance.EnableEvents = True





    End Sub


    ''' <summary>
    ''' erstellt das Excel Export bzw. Vorlagen  File für die angegebenen Phasen, Meilensteine, Rollen und Kosten
    ''' vorläufig nur für Phasen und Meilensteine realisiert
    ''' </summary>
    ''' <param name="filterName">gibt den Namen des Filters an, der die Collections enthält </param>
    ''' <remarks></remarks>
    Friend Sub createDateiFromSelection(ByVal filterName As String, ByVal menueOption As Integer)

        Dim earliestDate As Date, latestDate As Date
        Dim phaseList As New SortedList(Of String, String)
        Dim milestonelist As New SortedList(Of String, String)

        Dim selphases As New Collection
        Dim selMilestones As New Collection
        Dim selRoles As New Collection
        Dim selCosts As New Collection
        Dim selBUs As New Collection
        Dim selTyps As New Collection

        Call retrieveSelections(filterName, menueOption, selBUs, selTyps, _
                                 selphases, selMilestones, selRoles, selCosts)

        ' initialisieren 
        earliestDate = StartofCalendar.AddMonths(-12)
        latestDate = StartofCalendar.AddMonths(1200)

        Dim anteil As Double = 0.0
        Dim anzahlProjekte As Integer = ShowProjekte.Count
        Dim currentIX As Integer
        Dim hproj As clsProjekt
        Dim pName As String, msName As String

        Dim anzPlanobjekte As Integer = selphases.Count + selMilestones.Count
        Dim bestproj As String = ""
        Dim startFaktor As Double = 1.0
        Dim durationFaktor As Double = 0.000001
        Dim correctFaktor As Double = 0.00000001
        Dim korrFaktor As Double
        Dim refLaenge As Integer
        Dim fullName As String = ""
        Dim breadcrumb As String = ""
        Dim listName As String = ""

        ' die selphases und selMilestones enthalten jetzt 

        currentIX = 1
        Do While currentIX <= anzahlProjekte

            hproj = ShowProjekte.getProject(currentIX)

            If currentIX = 1 Then
                korrFaktor = 1.0
                refLaenge = hproj.dauerInDays
            Else
                Try
                    korrFaktor = hproj.dauerInDays / refLaenge
                Catch ex As Exception
                    korrFaktor = 1.0
                End Try

            End If

            ' es wird einfach der Reihenfolge nach eingetragen
            ' eine vorherige Überprüfung, welche Meilensteine grundsätzlich vorne stehen, wird nicht mehr gemacht 

            For Each pObject As Object In selphases

                pName = ""
                breadcrumb = ""
                fullName = CStr(pObject)
                Dim type As Integer = -1
                Dim pvName As String = ""
                Call splitHryFullnameTo2(fullName, pName, breadcrumb, type, pvName)

                ' jetzt muss eine Schleife gemacht werden über alle Vorkommen dieses Namens
                Dim anzahlElements As Integer = hproj.hierarchy.getPhaseIndices(pName, breadcrumb).Length

                For ce As Integer = 1 To anzahlElements

                    listName = fullName & "#" & ce.ToString("00#")

                    If phaseList.ContainsKey(listName) Then
                        ' nichts tun, dann ist sie schon eingeordnet 
                    Else

                        ' schlüssel kann gar nicht mehrfach vorkommen) 
                        phaseList.Add(listName, listName)


                    End If
                Next
            Next


            For Each pObject As Object In selMilestones

                msName = ""
                breadcrumb = ""
                fullName = CStr(pObject)
                Dim type As Integer = -1
                Dim pvName As String = ""
                Call splitHryFullnameTo2(fullName, msName, breadcrumb, type, pvName)

                ' jetzt muss eine Schleife gemacht werden über alle Vorkommen dieses Namens
                Dim anzahlElements As Integer = CInt(hproj.hierarchy.getMilestoneIndices(msName, breadcrumb).Length / 2)


                For ce As Integer = 1 To anzahlElements

                    listName = fullName & "#" & ce.ToString("00#")

                    If milestonelist.ContainsKey(listName) Then
                        ' nichts tun, dann ist sie schon eingeordnet 
                    Else

                        ' schlüssel kann gar nicht mehrfach vorkommen) 
                        milestonelist.Add(listName, listName)

                    End If

                Next

                ' alt 

            Next


            currentIX = currentIX + 1

        Loop

        ' jetzt sind die Elemente in der richtigen Reihenfolge eingeordnet 
        ' jetzt werden sie rausgeschrieben 
        Try
            If menueOption = PTmenue.excelExport Then
                Call exportSelectionToExcel(phaseList, milestonelist)
            ElseIf menueOption = PTmenue.vorlageErstellen Then
                Call createVorlageFromSelection(phaseList, milestonelist)
            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try



    End Sub
    ''' <summary>
    ''' speichert den letzten Filter unter "fname" und setzt die temporären Collections wieder zurück 
    ''' </summary>
    ''' <remarks></remarks>
    '''
    Public Sub storeFilter(ByVal fName As String, ByVal menuOption As Integer, _
                                              ByVal fBU As Collection, ByVal fTyp As Collection, _
                                              ByVal fPhase As Collection, ByVal fMilestone As Collection, _
                                              ByVal fRole As Collection, ByVal fCost As Collection, _
                                              ByVal calledFromHry As Boolean)

        Dim lastFilter As clsFilter

        If menuOption = PTmenue.filterdefinieren Or _
            menuOption = PTmenue.sessionFilterDefinieren Or _
            menuOption = PTmenue.filterAuswahl Then

            If calledFromHry Then
                Dim nameLastFilter As clsFilter = filterDefinitions.retrieveFilter("Last")

                If Not IsNothing(nameLastFilter) Then
                    With nameLastFilter
                        lastFilter = New clsFilter(fName, .BUs, .Typs, fPhase, fMilestone, .Roles, .Costs)
                    End With
                Else
                    lastFilter = New clsFilter(fName, fBU, fTyp, _
                                      fPhase, fMilestone, _
                                     fRole, fCost)
                End If


            Else
                lastFilter = New clsFilter(fName, fBU, fTyp, _
                                      fPhase, fMilestone, _
                                     fRole, fCost)
            End If



            filterDefinitions.storeFilter(fName, lastFilter)

            If Not noDB Then


                ' Filter mit Namen "fName" in DB speichern
                Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)

                ' Datenbank ist gestartet
                If request.pingMongoDb() Then

                    Dim filterToStoreInDB As clsFilter = filterDefinitions.retrieveFilter(fName)
                    Dim returnvalue As Boolean = request.storeFilterToDB(filterToStoreInDB, False)
                    If returnvalue = False Then
                        Call MsgBox("Fehler bei Schreiben Filter: " & fName)
                    End If
                Else
                    Call MsgBox(" Datenbank-Verbindung ist unterbrochen!" & vbLf & " Filter kann nicht in DB gespeichert werden")
                End If


            End If

        Else        ' nicht menuOption = PTmenue.filterdefinieren

            If calledFromHry Then
                Dim nameLastFilter As clsFilter = selFilterDefinitions.retrieveFilter("Last")

                If Not IsNothing(nameLastFilter) Then
                    With nameLastFilter
                        lastFilter = New clsFilter(fName, .BUs, .Typs, fPhase, fMilestone, .Roles, .Costs)
                    End With
                Else
                    lastFilter = New clsFilter(fName, fBU, fTyp, _
                                      fPhase, fMilestone, _
                                     fRole, fCost)
                End If


            Else
                lastFilter = New clsFilter(fName, fBU, fTyp, _
                                      fPhase, fMilestone, _
                                     fRole, fCost)
            End If

            selFilterDefinitions.storeFilter(fName, lastFilter)

            If Not noDB Then

                ' Filter mit Namen "fName" in DB speichern
                Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)

                ' Datenbank ist gestartet
                If request.pingMongoDb() Then

                    Dim filterToStoreInDB As clsFilter = selFilterDefinitions.retrieveFilter(fName)
                    Dim returnvalue As Boolean = request.storeFilterToDB(filterToStoreInDB, True)
                Else
                    Call MsgBox(" Datenbank-Verbindung ist unterbrochen!" & vbLf & " Filter kann nicht in DB gespeichert werden")
                End If

            End If

        End If

    End Sub

    ''' <summary>
    ''' löscht das angegebene Projekt mit Name pName inkl all seiner Varianten 
    ''' </summary>
    ''' <param name="pName">
    ''' gibt an , ob es der erste Aufruf war
    ''' wenn ja, kommt erst der Bestätigungs-Dialog 
    ''' wenn nein, wird ohne Aufforderung zur Bestätigung gelöscht 
    ''' </param>
    ''' <remarks></remarks>
    Public Sub awinDeleteProjectInSession(ByVal pName As String,
                                          Optional ByVal considerDependencies As Boolean = False, _
                                          Optional ByVal upDateDiagrams As Boolean = False, _
                                          Optional ByVal vName As String = Nothing)


        Dim hproj As clsProjekt

        Dim tmpCollection As New Collection

        Dim formerEOU As Boolean = enableOnUpdate
        enableOnUpdate = False


        If ShowProjekte.contains(pName) Then

            ' Aktuelle Konstellation ändert sich dadurch
            If currentConstellationName <> calcLastSessionScenarioName() Then
                currentConstellationName = calcLastSessionScenarioName()
            End If

            hproj = ShowProjekte.getProject(pName)
            If IsNothing(vName) Or vName = hproj.variantName Then
                Call putProjectInNoShow(hproj.name, considerDependencies, upDateDiagrams)
            End If


        End If

        ' jetzt müssen alle oder die ausgewählte Variante aus AlleProjekte gelöscht werden 
        If IsNothing(vName) Then
            AlleProjekte.RemoveAllVariantsOf(pName)
        Else
            Dim key As String = calcProjektKey(pName, vName)
            If AlleProjekte.Containskey(key) Then
                AlleProjekte.Remove(key)
            End If
        End If

        enableOnUpdate = formerEOU

    End Sub


    ''' <summary>
    ''' nimmt das angegebene Projekt aus ShowProjekte heraus
    ''' löscht das Projekt auf der Plan-Tafel und schicbt die restlichen Projekte weiter nach oben 
    ''' wenn considerDependencies=true: dann werden alle abhängigen Projekte, die ebenfalls im ShowProjekte sind, auch rausgenommen
    ''' wenn upDateDiagrams=true: alle Diagramme werde neu gezeichnet  
    ''' 
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <param name="considerDependencies"></param>
    ''' <param name="upDateDiagrams"></param>
    ''' <remarks></remarks>
    Public Sub putProjectInNoShow(ByVal pName As String, ByVal considerDependencies As Boolean, ByVal upDateDiagrams As Boolean)

        Dim pZeile As Integer
        Dim tmpCollection As New Collection
        Dim anzahlZeilen As Integer = 1

        If ShowProjekte.contains(pName) Then

            Dim hproj As clsProjekt = ShowProjekte.getProject(pName)
            pZeile = calcYCoordToZeile(projectboardShapes.getCoord(pName)(0))

            If hproj.extendedView Then
                anzahlZeilen = _
                    hproj.calcNeededLines(tmpCollection, tmpCollection, awinSettings.drawphases Or hproj.extendedView, False)
            End If

            'pZeile = ShowProjekte.getPTZeile(selectedProjectName)
            'Call MsgBox("Zeile: " & pZeile.ToString)

            Call clearProjektinPlantafel(pName)

            ShowProjekte.Remove(pName)

            Call moveShapesUp(pZeile + 1, anzahlZeilen, True)

        End If

        ' jetzt muss noch geprüft werden , ob considerDependencies true ist 
        If considerDependencies Then
            ' ggf. die Projekte einblenden, von denen dieses Projekt abhängt 
            Dim toDoListe As Collection = allDependencies.activeListe(pName, PTdpndncyType.inhalt)
            If toDoListe.Count > 0 Then
                For Each dprojectName As String In toDoListe
                    Call putProjectInNoShow(dprojectName, considerDependencies, False)
                Next

            End If
        Else
            ' nichts tun 
        End If

        If upDateDiagrams Then
            ' jetzt müssen die Portfolio Diagramme neu gezeichnet werden 
            Call awinNeuZeichnenDiagramme(2)
        End If


    End Sub

    ''' <summary>
    ''' bringt die angegebene Projekt-Variante ins Show ... 
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <param name="vNAme"></param>
    ''' <param name="considerDependencies"></param>
    ''' <param name="upDateDiagrams"></param>
    ''' <remarks></remarks>
    Public Sub putProjectInShow(ByVal pName As String, ByVal vName As String, _
                                    ByVal considerDependencies As Boolean, _
                                    ByVal upDateDiagrams As Boolean, _
                                    ByVal myConstellation As clsConstellation, _
                                    Optional ByVal parentChoice As Boolean = False, _
                                    Optional pZeile As Integer = -1)

        Dim key As String = calcProjektKey(pName, vName)
        Dim hproj As clsProjekt = AlleProjekte.getProject(key)


        If IsNothing(hproj) And parentChoice Then
            Dim variantNames As Collection = AlleProjekte.getVariantNames(pName, False)
            vName = CStr(variantNames.Item(1))
            key = calcProjektKey(pName, vName)
            hproj = AlleProjekte.getProject(key)
        End If

        ' wenn immer noch Nothing, nichts tun ... 
        If IsNothing(hproj) Then
            Exit Sub
        End If

        Dim anzahlZeilen As Integer = 1

        If Not ShowProjekte.contains(pName) Then
            ShowProjekte.Add(hproj)
            If pZeile < 2 Then
                'pZeile = ShowProjekte.getPTZeile(pName)
                pZeile = myConstellation.getBoardZeile(pName)
            End If

            Dim tmpCollection As New Collection

            If hproj.extendedView Then
                anzahlZeilen = _
                    hproj.calcNeededLines(tmpCollection, tmpCollection, awinSettings.drawphases Or hproj.extendedView, False)
            End If

            If pZeile > 0 Then
                Call moveShapesDown(tmpCollection, pZeile, anzahlZeilen, 0)
                Call ZeichneProjektinPlanTafel(tmpCollection, pName, pZeile, tmpCollection, tmpCollection)
            End If
        End If

        ' jetzt muss das Projekt neu gezeichnet werden ; 
        ' dazu muss die Einfügestelle bestimmt werden, dann alle anderen Shapes nach unten verschoben werden 
        ' hier muss die Zeile über Showprojekte bestimmt werden, einfach nach der Sortier-Reihenfolge 
        ' das kann später dann noch angepasst werden 

        'Dim pZeile2 As Integer = node.Index
        'Call MsgBox("Zeile: " & pZeile.ToString)



        ' jetzt muss noch geprüft werden , ob considerDependencies true ist 
        If considerDependencies Then
            ' ggf. die Projekte einblenden, von denen dieses Projekt abhängt 
            Dim toDoListe As Collection = allDependencies.passiveListe(pName, PTdpndncyType.inhalt)
            If toDoListe.Count > 0 Then
                For Each mprojectName As String In toDoListe
                    Call putProjectInShow(pName:=mprojectName, _
                                          vName:="", considerDependencies:=considerDependencies, _
                                          upDateDiagrams:=False, _
                                          myConstellation:=myConstellation, parentChoice:=True)
                Next

            End If
        Else
            ' nichts tun 
        End If

        If upDateDiagrams Then
            ' jetzt müssen die Portfolio Diagramme neu gezeichnet werden 
            Call awinNeuZeichnenDiagramme(2)
        End If

    End Sub

    ''' <summary>
    ''' verallgemeinerte Import Routine, ähnlich wie BMWimport
    ''' wenn treatAsPhases = true, werden die einzelnen Pläne als Sammelvorgänge innerhalb ein und desselben Projektes aufgefasst  
    ''' </summary>
    ''' <param name="myCollection"></param>
    ''' <param name="isVorlage"></param>
    ''' <remarks></remarks>
    Public Sub planExcelImport(ByRef myCollection As Collection, ByVal isVorlage As Boolean, ByVal dateiname As String)

        Dim phaseHierarhy(9) As String
        Dim currentHierarchy As Integer = 0
        Dim zeile As Integer, spalte As Integer
        Dim pName As String = " "
        Dim phaseName As String = " "
        Dim currentDateiName As String
        Dim isMilestone As Boolean

        Dim lastRow As Integer

        Dim hproj As clsProjekt
        Dim vproj As clsProjektvorlage
        Dim vglName As String = ""
        Dim vglProj As New clsProjekt
        Dim geleseneProjekte As Integer
        Dim projektFarbe As Object
        Dim anfang As Integer, ende As Integer
        Dim cphase As clsPhase
        Dim cmilestone As clsMeilenstein
        Dim cbewertung As clsBewertung
        Dim ix As Integer
        Dim tmpStr(20) As String
        Dim completeName As String
        Dim nameSopTyp As String = " "
        Dim nameProduktlinie As String = ""
        Dim defaultBU As String = ""

        Dim startDate As Date, endDate As Date
        Dim startoffset As Long, duration As Long
        Dim vorlagenName As String = ""
        Dim variantenName As String = ""

        Dim itemName As String = ""
        Dim zufall As New Random(10)
        Dim itemDauer As Integer
        Dim colProtocol As Integer

        Dim schriftGroesse As Integer
        Dim schriftfarbe As Long

        ' Kennungen für die BMW Projekte
        Dim typKennung As String = ""
        Dim anlaufKennung As String = ""
        Dim anzProcessedElements As Integer = 0
        Dim anzSubstituted As Integer = 0
        Dim anzIgnored As Integer = 0
        Dim anzCorrect As Integer = 0

        ' 
        Dim logMessage As String = ""

        ' ur: 1.12.2015: wird nun Public awinSettings.fullProtokoll As Boolean = True  
        ' und damit global definiert, da auch in RXFImport benötigt.
        ' Dim fullProtocol As Boolean = True


        Dim milestoneIX As Integer = MilestoneDefinitions.Count + 1
        Dim phaseIX As Integer = PhaseDefinitions.Count + 1
        ' wird benötigt, um bei Phasen, die als doppelt erkannt wurden alle darunter liegenden Elemente auch zu ignorieren 
        Dim lastDuplicateIndent As Integer = 1000000

        ' bestimmen, des eventuell benötigten VariantenName. Dieser wird aus dem Dateinamen erstellt
        Dim tmpStrNew() As String
        tmpStrNew = Split(dateiname, "\", -1)
        variantenName = tmpStrNew(tmpStrNew.Length - 1)


        ' Vorbedingung: das Excel File. das importiert werden soll , ist bereits geöffnet 

        Dim colName As Integer
        Dim colAnfang As Integer
        Dim colEnde As Integer
        Dim colDauer As Integer = -1
        Dim colProduktlinie As Integer = -1
        Dim colAbbrev As Integer = -1
        Dim colVorgangsKlasse As Integer = -1
        Dim colDescription As Integer = -1

        Dim pDescription As String = ""
        Dim firstZeile As Excel.Range
        Dim protocolRange As Excel.Range


        Dim suchstr(8) As String
        suchstr(ptPlanNamen.Name) = "Name"
        suchstr(ptPlanNamen.Anfang) = "Start"
        suchstr(ptPlanNamen.Ende) = "End"
        suchstr(ptPlanNamen.Beschreibung) = "Description"
        suchstr(ptPlanNamen.Vorgangsklasse) = "Appearance"
        suchstr(ptPlanNamen.BusinessUnit) = "Business Unit"
        suchstr(ptPlanNamen.Protocol) = "Übernommen als"
        suchstr(ptPlanNamen.Dauer) = "Duration"
        suchstr(ptPlanNamen.Abkuerzung) = "Abbreviation"


        zeile = 2
        spalte = 5
        geleseneProjekte = 0

        ' wie lautet der aktuelle Dateiname ? 
        currentDateiName = CType(appInstance.ActiveWorkbook, Excel.Workbook).Name

        ' wie lautet ggf der Default Produktlinien Name ? 
        Dim i As Integer
        Dim found As Boolean = False
        Dim tmpName As String
        i = 1
        While i <= businessUnitDefinitions.Count And Not found

            tmpName = businessUnitDefinitions.ElementAt(i - 1).Value.name
            If currentDateiName.Contains(tmpName) Then
                defaultBU = tmpName
                found = True
            Else
                i = i + 1
            End If

        End While



        Dim aktivesSheet As Excel.Worksheet = CType(appInstance.ActiveWorkbook.ActiveSheet, _
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet)

        With aktivesSheet
            firstZeile = CType(.Rows(1), Excel.Range)
        End With



        ' diese Daten müssen vorhanden sein - andernfalls Abbruch 
        Try
            colName = firstZeile.Find(What:=suchstr(ptPlanNamen.Name), LookAt:=XlLookAt.xlWhole).Column
            colAnfang = firstZeile.Find(What:=suchstr(ptPlanNamen.Anfang), LookAt:=XlLookAt.xlWhole).Column
            colEnde = firstZeile.Find(What:=suchstr(ptPlanNamen.Ende), LookAt:=XlLookAt.xlWhole).Column

        Catch ex As Exception
            Throw New ArgumentException("Fehler im Datei Aufbau ..." & vbLf & ex.Message)
        End Try

        Try
            colDauer = firstZeile.Find(What:=suchstr(ptPlanNamen.Dauer), LookAt:=XlLookAt.xlWhole).Column
        Catch ex As Exception
            colDauer = -1
        End Try


        Try
            colProduktlinie = firstZeile.Find(What:=suchstr(ptPlanNamen.BusinessUnit), LookAt:=XlLookAt.xlWhole).Column
        Catch ex As Exception
            colProduktlinie = -1
        End Try


        Try
            colAbbrev = firstZeile.Find(What:=suchstr(ptPlanNamen.Abkuerzung), LookAt:=XlLookAt.xlWhole).Column
        Catch ex As Exception
            colAbbrev = -1
        End Try

        Try
            colDescription = firstZeile.Find(What:=suchstr(ptPlanNamen.Beschreibung), LookAt:=XlLookAt.xlWhole).Column
        Catch ex As Exception
            colAbbrev = -1
        End Try

        Try
            colVorgangsKlasse = firstZeile.Find(What:=suchstr(ptPlanNamen.Vorgangsklasse), LookAt:=XlLookAt.xlWhole).Column
        Catch ex As Exception

        End Try


        With aktivesSheet

            lastRow = System.Math.Max(CType(.Cells(40000, colName), Global.Microsoft.Office.Interop.Excel.Range).End(Excel.XlDirection.xlUp).Row, _
                                          CType(.Cells(40000, colAnfang), Global.Microsoft.Office.Interop.Excel.Range).End(Excel.XlDirection.xlUp).Row)
        End With




        ' Hier wird die Stelle und die Informationen für das Visbo Protocoll ermittelt und gesetzt 
        Dim protocolCellName As String = "VISBO_Protocol"
        Dim pCell As Excel.Range

        With aktivesSheet
            Try
                colProtocol = .Range(protocolCellName).Column
                protocolRange = CType(.Range(.Cells(1, colProtocol - 3), .Cells(lastRow + 10, colProtocol + 200)), Excel.Range)
                protocolRange.Clear()
                protocolRange.Interior.Color = RGB(255, 255, 255)
                protocolRange.ClearFormats()

            Catch ex As Exception
                Try
                    colProtocol = CType(.Cells(1, 2000), Global.Microsoft.Office.Interop.Excel.Range).End(Excel.XlDirection.xlToLeft).Column + 4
                Catch ex1 As Exception
                    colProtocol = 20
                End Try
                pCell = .Range(.Cells(1, colProtocol), .Cells(1, colProtocol))
                appInstance.ActiveWorkbook.Names.Add(Name:=protocolCellName, RefersToR1C1:=pCell)

                ' dann müssen auch die Spaltenbreiten gesetzt werden 
                Dim tmpRange As Excel.Range
                With aktivesSheet

                    For i = -3 To 9
                        tmpRange = CType(aktivesSheet.Columns(colProtocol + i), Excel.Range)
                        tmpRange.ColumnWidth = 40
                    Next


                End With

            End Try


        End With


        ' Die Überschriften für das Protokoll werden alle wieder gesetzt 
        With aktivesSheet


            If awinSettings.fullProtocol Then

                CType(.Cells(1, colProtocol), Excel.Range).Value = "Projekt"
                CType(.Cells(1, colProtocol + 1), Excel.Range).Value = "Hierarchie"
                CType(.Cells(1, colProtocol + 2), Excel.Range).Value = "Plan-Element"
                CType(.Cells(1, colProtocol + 3), Excel.Range).Value = "Klasse"
                CType(.Cells(1, colProtocol + 4), Excel.Range).Value = "Abkürzung"
                CType(.Cells(1, colProtocol + 5), Excel.Range).Value = "Quelle"
                CType(.Cells(1, colProtocol + 8), Excel.Range).Value = "PT Hierarchie"
                CType(.Cells(1, colProtocol + 9), Excel.Range).Value = "PT Klasse"
            End If

            ' wird immer geschrieben 
            CType(.Cells(1, colProtocol + 6), Excel.Range).Value = suchstr(ptPlanNamen.Protocol)
            CType(.Cells(1, colProtocol + 7), Excel.Range).Value = "Grund"

        End With

        Try

            With aktivesSheet

                Try
                    projektFarbe = CType(aktivesSheet.Cells(zeile, 1), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color
                    ' das Folgende wird nur für die Projekt-Vorlagen benötigt (isVorlage = true) 
                    schriftfarbe = CLng(CType(aktivesSheet.Cells(zeile, 1), Global.Microsoft.Office.Interop.Excel.Range).Font.Color)
                    schriftGroesse = CInt(CType(aktivesSheet.Cells(zeile, 1), Global.Microsoft.Office.Interop.Excel.Range).Font.Size)

                Catch ex As Exception
                    projektFarbe = CType(aktivesSheet.Cells(zeile, 1), Excel.Range).Interior.ColorIndex
                End Try

                ' jetzt kommt der Check, ob Blanks als Indent verwendet werden oder echte Excel Indents
                Dim stdIndent As Boolean = True
                Dim stdIndentedRows As Integer = 0
                Dim blankIndentedRows As Integer = 0
                For ik As Integer = 1 To lastRow
                    If CType(.Cells(ik, colName), Excel.Range).IndentLevel > 0 Then
                        stdIndentedRows = stdIndentedRows + 1
                    End If
                    Dim tstString As String = CStr(CType(.Cells(ik, colName), Excel.Range).Value)
                    If tstString.StartsWith(" ") Then
                        blankIndentedRows = blankIndentedRows + 1
                    End If
                Next

                If stdIndentedRows > blankIndentedRows Then
                    stdIndent = True
                Else
                    stdIndent = False
                End If

                ' zeile ist an der Stelle 2
                While zeile <= lastRow

                    ' wenn es mit einem neuen Projekt beginnt, muss der lastDuplicateIndent zurückgesetzt sein 
                    lastDuplicateIndent = 1000000

                    ix = zeile + 1

                    Dim zellenFarbe As Long = CLng(CType(.Cells(ix, 1), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color)
                    Do While zellenFarbe <> CLng(projektFarbe) And (ix <= lastRow)
                        ix = ix + 1
                        zellenFarbe = CLng(CType(.Cells(ix, 1), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color)
                    Loop

                    anfang = zeile + 1
                    ende = ix - 1

                    ' hier wird Name, Typ, SOP, Business Unit, vname, Start-Datum, Dauer der Phase(1) ausgelesen  

                    ' ur: 24.06.2016:testweise auskomentiert
                    ' '' ''endDate = CDate(.Cells(RowIndex:=zeile, ColumnIndex:=colEnde).value)
                    ' '' ''startDate = CDate(.Cells(RowIndex:=zeile, ColumnIndex:=colAnfang).value)

                    ' '' ''completeName = CStr(.Cells(RowIndex:=zeile, ColumnIndex:=colName).value)

                    startDate = CDate(CType(.Cells(zeile, colAnfang), Global.Microsoft.Office.Interop.Excel.Range).Value)
                    endDate = CDate(CType(.Cells(zeile, colEnde), Global.Microsoft.Office.Interop.Excel.Range).Value)
                    completeName = CStr(CType(.Cells(zeile, colName), Global.Microsoft.Office.Interop.Excel.Range).Value)

                    ' andere Informationen auslesen ... 
                    pDescription = ""
                    If colDescription > 0 Then
                        pDescription = CStr(CType(.Cells(zeile, colDescription), Global.Microsoft.Office.Interop.Excel.Range).Value)
                    End If

                    defaultBU = ""
                    If colProduktlinie > 0 Then

                        Try
                            Dim tmpBU As String
                            If colProduktlinie > 0 Then
                                tmpBU = CStr(CType(.Cells(zeile, colProduktlinie), Global.Microsoft.Office.Interop.Excel.Range).Value)
                            Else
                                tmpBU = ""
                            End If


                            ' gibt es die Business Unit ? 
                            found = False
                            Dim bix As Integer = 1

                            If tmpBU.Length > 0 Then
                                While bix <= businessUnitDefinitions.Count And Not found
                                    If businessUnitDefinitions.ElementAt(bix - 1).Value.name = tmpBU Then

                                        found = True
                                        defaultBU = tmpBU

                                    Else
                                        bix = bix + 1
                                    End If
                                End While
                            End If


                            If Not found Then

                                CType(aktivesSheet.Cells(zeile, colProduktlinie), Excel.Range).Interior.Color = awinSettings.AmpelRot

                            End If

                        Catch ex1 As Exception

                        End Try

                    End If

                    Dim tmpvalue As String
                    Dim tmp2Str() As String

                    If colDauer > 0 Then

                        Try
                            tmpvalue = CStr(CType(aktivesSheet.Cells(zeile, colDauer), Excel.Range).Value).Trim
                            tmp2Str = tmpvalue.Trim.Split(New Char() {CChar(" ")}, 5)
                            itemDauer = CInt(tmp2Str(0))
                        Catch ex As Exception
                            itemDauer = -1
                        End Try
                    End If


                    duration = DateDiff(DateInterval.Day, startDate, endDate) + 1
                    If duration < 0 Then
                        startDate = endDate
                        duration = -1 * duration
                        endDate = startDate.AddDays(duration)
                    End If

                    tmpStr = completeName.Trim.Split(New Char() {CChar("["), CChar("]")}, 5)

                    ' PT-71 Änderung 22.1.15 (tk) Der Projekt-Name soll der RPLAN Name sein 
                    'pName = tmpStr(0).Trim
                    ' damit alt: 
                    ' jetzt doch wieder hereingenommen, weil sich von einem Monat auf den anderen ein und dasselbe Projekte im SOP ändert .... 
                    Dim doADD As Boolean = False


                    pName = tmpStr(0)



                    ' prüfen, ob das Projekt überhaupt vollständig im Kalender liegt 
                    ' wenn nein, dann nicht importieren 
                    If DateDiff(DateInterval.Day, StartofCalendar, startDate) < 0 Then


                        Call MsgBox("Projekt liegt vor dem Kalender-Anfang und wird deshalb nicht importiert")


                    Else
                        '
                        ' jetzt wird das Projekt angelegt 
                        '
                        hproj = New clsProjekt


                        Try

                            hproj.name = pName
                            hproj.startDate = startDate
                            ' Projekte sollten erstmal nicht verschoben werden können
                            ' dazu muss eine Variante erzeugt werden , die kann dann verschoben werden 
                            hproj.Status = ProjektStatus(1)

                            If DateDiff(DateInterval.Month, startDate, Date.Now) <= 0 Then
                                hproj.earliestStartDate = hproj.startDate.AddMonths(hproj.earliestStart)
                                hproj.latestStartDate = hproj.startDate.AddMonths(hproj.latestStart)
                            Else
                                hproj.earliestStartDate = startDate
                                hproj.latestStartDate = startDate
                            End If

                            hproj.StrategicFit = 5
                            hproj.Risiko = 5
                            hproj.businessUnit = defaultBU
                            hproj.description = pDescription

                            hproj.Erloes = 0.0


                        Catch ex As Exception
                            Throw New Exception("in erstelle Import Excel Projekte: " & vbLf & ex.Message)
                        End Try

                        ' jetzt wird die Import Hierarchie angelegt 
                        Dim pHierarchy As New clsImportFileHierarchy
                        Dim origHierarchy As New clsImportFileHierarchy

                        ' jetzt wird die Projekt-Hierarchie neu angelegt 
                        ' die erste Phase, die sogenannte Root Phase hat immer diesen Namen: 

                        ' jetzt werden all die Phasen angelegt , beginnend mit der ersten 
                        cphase = New clsPhase(parent:=hproj)
                        cphase.nameID = rootPhaseName
                        startoffset = 0
                        duration = DateDiff(DateInterval.Day, startDate, endDate) + 1
                        cphase.changeStartandDauer(startoffset, duration)

                        hproj.AddPhase(cphase)

                        Try
                            pHierarchy.add(cphase, rootPhaseName, 0)
                            origHierarchy.add(cphase, rootPhaseName, 0)
                        Catch ex As Exception

                        End Try

                        Dim itemStartDate As Date
                        Dim itemEndDate As Date
                        Dim ok As Boolean = True

                        Dim curZeile As Integer
                        Dim txtVorgangsKlasse As String
                        Dim origVorgangsKlasse As String
                        Dim txtAbbrev As String
                        ' ist notwendig um anhand der führenden Blanks die Hierarchie Stufe zu bestimmen 
                        Dim origItem As String = ""

                        ' 
                        ' Schleife, um alle Elemente des Projektes auszulesen
                        ' hier werden jetzt die einzelnen Zeilen = Phasen oder Meilensteine ausgelesen 
                        For curZeile = anfang To ende

                            origVorgangsKlasse = ""
                            txtVorgangsKlasse = ""
                            txtAbbrev = ""
                            logMessage = ""

                            Dim indentLevel As Integer

                            Try

                                Dim tmpName2 As String = CStr(CType(.Cells(curZeile, colName), Excel.Range).Value)

                                tmpStr = tmpName2.Split(New Char() {CChar("["), CChar("]")}, 5)
                                origItem = tmpStr(0)

                                If origItem.Trim.Length = 0 Then

                                    'CType(aktivesSheet.Cells(curZeile, colProtocol + 7), Excel.Range).Value = _
                                    '            "leerer String wird ignoriert .."
                                    logMessage = "leerer String wird ignoriert .."
                                    ok = False

                                Else

                                    If stdIndent Then
                                        indentLevel = CType(.Cells(curZeile, colName), Excel.Range).IndentLevel
                                    Else
                                        indentLevel = pHierarchy.getLevel(origItem)
                                    End If

                                    ' hier checken, ob indentlevel > lastduplicateIndent; 
                                    ' wenn ja, dann protokollieren, Next for und lastduplicateIndent wieder auf hohen Wert setzen

                                    If indentLevel > lastDuplicateIndent Then
                                        ' Skip , weil es sich dann um Elemente handelt, deren Parent Phase als Duplikat ignoriert wurde 
                                        ' Protokollieren ...

                                        'CType(aktivesSheet.Cells(curZeile, colProtocol + 7), Excel.Range).Value = _
                                        '            "ist Kind eines doppelten/nicht zugelassenen Elements und wird ignoriert"

                                        logMessage = "ist Kind eines doppelten/nicht zugelassenen Elements und wird ignoriert"
                                        ok = False

                                    Else
                                        lastDuplicateIndent = 1000000

                                        itemName = origItem.Trim

                                        anzProcessedElements = anzProcessedElements + 1


                                        If awinSettings.fullProtocol Then

                                            CType(aktivesSheet.Cells(curZeile, colProtocol + 2), Excel.Range).Value = origItem.Trim
                                            CType(aktivesSheet.Cells(curZeile, colProtocol), Excel.Range).Value = completeName
                                            CType(aktivesSheet.Cells(curZeile, colProtocol + 5), Excel.Range).Value = currentDateiName
                                        End If



                                        ' Änderung 26.1.15 Ignorieren 

                                        itemStartDate = CDate(CType(.Cells(curZeile, colAnfang), Excel.Range).Value)
                                        itemEndDate = CDate(CType(.Cells(curZeile, colEnde), Excel.Range).Value)

                                        If IsNothing(CType(.Cells(curZeile, colAnfang), Excel.Range).Value) Then
                                            isMilestone = True
                                            itemStartDate = itemEndDate
                                        ElseIf CStr(CType(.Cells(curZeile, colAnfang), Excel.Range).Value).Trim = "" Then
                                            isMilestone = True
                                            itemStartDate = itemEndDate
                                        Else
                                            If DateDiff(DateInterval.Minute, itemStartDate, itemEndDate) = 0 Then
                                                isMilestone = True
                                            Else
                                                isMilestone = False
                                            End If
                                        End If



                                        ' jetzt prüfen, ob es sich um ein grundsätzlich zu ignorierendes Element handelt .. 
                                        If isMilestone Then
                                            If MilestoneDefinitions.Contains(itemName) Then
                                                ok = True
                                            ElseIf milestoneMappings.tobeIgnored(itemName) Then
                                                'CType(aktivesSheet.Cells(curZeile, colProtocol + 7), Excel.Range).Value = _
                                                '                "nicht zugelassen (lt. Wörterbuch ignorieren)"

                                                logMessage = "nicht zugelassen (lt. Wörterbuch ignorieren)"
                                                ok = False
                                                lastDuplicateIndent = indentLevel
                                            Else
                                                ok = True
                                            End If


                                        Else

                                            If PhaseDefinitions.Contains(itemName) Then
                                                ok = True
                                            ElseIf phaseMappings.tobeIgnored(itemName) Then
                                                'CType(aktivesSheet.Cells(curZeile, colProtocol + 7), Excel.Range).Value = _
                                                '                "nicht zugelassen (lt. Wörterbuch ignorieren)"
                                                logMessage = "nicht zugelassen (lt. Wörterbuch ignorieren)"
                                                lastDuplicateIndent = indentLevel
                                                ok = False
                                            Else
                                                ok = True

                                            End If

                                        End If

                                    End If

                                End If

                            Catch ex As Exception
                                itemName = ""
                                ok = False
                            End Try


                            If ok Then


                                startoffset = DateDiff(DateInterval.Day, hproj.startDate, itemStartDate)
                                duration = DateDiff(DateInterval.Day, itemStartDate, itemEndDate) + 1


                                ' jetzt werden vorgangsklasse und Abkürzung rausgelesen 
                                If colVorgangsKlasse > 0 Then
                                    Try

                                        origVorgangsKlasse = CStr((CType(.Cells(curZeile, colVorgangsKlasse), Excel.Range).Value)).Trim
                                        If duration > 1 Then
                                            txtVorgangsKlasse = mapToAppearance(origVorgangsKlasse, False)
                                            'CType(activeWSListe.Cells(curZeile, protocolColumn + 2), Excel.Range).Value = _
                                            '        "auf folgende Phasen Darstellungsklasse abgebildet: " & txtVorgangsKlasse.Trim
                                        Else
                                            txtVorgangsKlasse = mapToAppearance(origVorgangsKlasse, True)
                                            'CType(activeWSListe.Cells(curZeile, protocolColumn + 2), Excel.Range).Value = _
                                            '        "auf folgende Meilenstein Darstellungsklasse abgebildet: " & txtVorgangsKlasse.Trim
                                        End If




                                    Catch ex As Exception

                                        'CType(activeWSListe.Cells(curZeile, protocolColumn + 2), Excel.Range).Value = _
                                        '            "Fehler bei Abbildung auf Darstellungsklasse ... " & txtVorgangsKlasse.Trim

                                    End Try
                                End If


                                ' jetzt wird die Abkürzung rausgelesen 
                                If colAbbrev > 0 Then
                                    Try

                                        txtAbbrev = CStr((CType(.Cells(curZeile, colAbbrev), Excel.Range).Value)).Trim

                                    Catch ex As Exception
                                        txtAbbrev = ""
                                    End Try
                                End If

                                '
                                ' jetzt muss protokolliert werden 
                                Dim oLevel As Integer
                                oLevel = origHierarchy.getLevel(origItem)
                                Dim oBreadCrumb As String = origHierarchy.getFootPrint(oLevel)


                                If awinSettings.fullProtocol Then

                                    ' Original Footprint
                                    CType(aktivesSheet.Cells(curZeile, colProtocol + 1), Excel.Range).Value = oBreadCrumb
                                    ' Textvorgangsklasse
                                    CType(aktivesSheet.Cells(curZeile, colProtocol + 3), Excel.Range).Value = origVorgangsKlasse
                                    ' Abkürzung
                                    CType(aktivesSheet.Cells(curZeile, colProtocol + 4), Excel.Range).Value = txtAbbrev
                                End If


                                ' jetzt muss ggf die Phase in die Orig Hierarchie aufgenommen werden 
                                If Not isMilestone Then

                                    Dim ophase As clsPhase
                                    ophase = New clsPhase(parent:=hproj)
                                    ophase.nameID = calcHryElemKey(origItem.Trim, False)
                                    'ophase.changeStartandDauer(startoffset, duration)

                                    Try
                                        origHierarchy.add(ophase, "dummy", oLevel)
                                    Catch ex As Exception

                                    End Try


                                End If

                                Dim stdName As String
                                Dim parentElemName As String
                                Dim parentNodeID As String
                                Dim elemID As String

                                ' If duration > 1 Or itemDauer > 0 Then
                                If duration > 1 Then
                                    ' es handelt sich um eine Phase 


                                    parentElemName = pHierarchy.getPhaseBeforeLevel(indentLevel).name
                                    ' das folgende wurde am 31.3. ergänzt, um die Hierarchie aufbauen zu können
                                    parentNodeID = pHierarchy.getIDBeforeLevel(indentLevel)

                                    ' Plausibilitäts-Check: die beiden müssen identisch sein !!
                                    ' tk Debug: 27.11.15
                                    'If elemNameOfElemID(parentNodeID) <> parentElemName Then
                                    '    Call MsgBox("nicht konsistent in bmwImportProjekteITO15, zeile 663")
                                    'End If


                                    ' jetzt den tatsächlichen Namen bestimmen , ggf wird dazu der Parent Phase Name benötigt 
                                    Try

                                        If Not PhaseDefinitions.Contains(itemName) Then
                                            stdName = phaseMappings.mapToStdName(parentElemName, itemName)
                                        Else
                                            stdName = itemName
                                        End If

                                    Catch ex As Exception
                                        stdName = itemName
                                    End Try


                                    Dim ok1 As Boolean = True


                                    'Dim breadcrumb As String = pHierarchy.getFootPrint(indentLevel, "#")
                                    Dim parentPhase As clsPhase = pHierarchy.getPhaseBeforeLevel(indentLevel)
                                    Dim parentphaseName As String = ""

                                    If Not IsNothing(parentPhase) Then
                                        parentphaseName = parentPhase.name
                                    End If


                                    ' sollen Duplikate eliminiert werden ?
                                    If awinSettings.eliminateDuplicates And hproj.hierarchy.containsKey(calcHryElemKey(stdName, False)) Then
                                        ' nur dann kann es Duplikate geben 
                                        If hproj.isCloneToParent(stdName, parentPhase.nameID, itemStartDate, itemEndDate, 0.97) Then
                                            ok1 = False
                                            logMessage = stdName & " ist Duplikat zu Parent " & parentPhase.name & " und wird ignoriert "

                                        Else
                                            Dim duplicateSiblingID As String = hproj.getDuplicatePhaseSiblingID(stdName, parentPhase.nameID, _
                                                                                                                 itemStartDate, itemEndDate, 0.97)

                                            If duplicateSiblingID = "" Then
                                                ok1 = True
                                            Else
                                                ok1 = False
                                                logMessage = stdName & " ist Duplikat zu Geschwister " & elemNameOfElemID(duplicateSiblingID) & _
                                                             " und wird ignoriert "
                                            End If
                                        End If



                                    End If



                                    ' jetzt muss geprüft werden, ob das Element in Std Definitions aufgenommen werden muss 
                                    Dim ok2 As Boolean = True
                                    If Not PhaseDefinitions.Contains(stdName) And ok1 Then

                                        Dim hphaseDef As clsPhasenDefinition
                                        hphaseDef = New clsPhasenDefinition

                                        hphaseDef.darstellungsKlasse = txtVorgangsKlasse
                                        hphaseDef.shortName = txtAbbrev
                                        hphaseDef.name = stdName
                                        hphaseDef.UID = phaseIX
                                        phaseIX = phaseIX + 1


                                        If isVorlage And awinSettings.alwaysAcceptTemplateNames Then
                                            ' in die Phase-Definitions aufnehmen 
                                            Try
                                                PhaseDefinitions.Add(hphaseDef)
                                            Catch ex As Exception
                                            End Try
                                        Else
                                            ' in Abhängigkeit vom Setting die Elemente aufnehmen oder nicht 
                                            Try
                                                If awinSettings.importUnknownNames Then
                                                    ok2 = True
                                                Else
                                                    ok2 = False
                                                    logMessage = "ist nicht in der Liste der zugelassenen Elemente enthalten"
                                                End If
                                                missingPhaseDefinitions.Add(hphaseDef)
                                            Catch ex As Exception
                                            End Try


                                        End If

                                    End If

                                    ' hier muss noch der letzte Check rein 

                                    If ok1 And ok2 Then

                                        ' hier muss jetzt überprüft werden, ob es Geschwister mit gleichen Namen gibt
                                        ' wenn ja , wird an den stdName solange eine ldfNR Ergänzung rangemacht, bis der NAme innerhalb der 
                                        ' Geschwistergruppe eindeutig ist

                                        ' Bestimmung des eindeutigen Namens innerhalb der Geschwister, unterschieden nach Meilensten  und Phase 
                                        If awinSettings.createUniqueSiblingNames Then
                                            stdName = hproj.hierarchy.findUniqueGeschwisterName(parentNodeID, stdName, False)
                                        End If

                                        elemID = hproj.hierarchy.findUniqueElemKey(stdName, False)

                                        ' das muss auf alle Fälle gemacht werden 
                                        cphase = New clsPhase(parent:=hproj)

                                        ' Änderung tk: jetzt muss die elemID in den Phasen Namen 
                                        cphase.nameID = elemID
                                        cphase.changeStartandDauer(startoffset, duration)

                                        ' der Aufbau der Hierarchie erfolgt in addphase
                                        hproj.AddPhase(cphase, origName:=origItem.Trim, _
                                                       parentID:=pHierarchy.getIDBeforeLevel(indentLevel))

                                        ' wird übernommen als 
                                        CType(aktivesSheet.Cells(curZeile, colProtocol + 6), Excel.Range).Value = stdName

                                        Dim PTBreadCrumb As String = hproj.hierarchy.getBreadCrumb(elemID)


                                        If awinSettings.fullProtocol Then

                                            CType(aktivesSheet.Cells(curZeile, colProtocol + 8), Excel.Range).Value = PTBreadCrumb
                                            CType(aktivesSheet.Cells(curZeile, colProtocol + 9), Excel.Range).Value = txtVorgangsKlasse
                                        End If
                                        ' neuer Breadcrumb 
                                        'Dim PTBreadCrumb As String = pHierarchy.getFootPrint(indentLevel)

                                        If stdName.Trim <> origItem.Trim Then
                                            ' es hat eine Ersetzung stattgefunden 
                                            CType(aktivesSheet.Cells(curZeile, colProtocol + 6), Excel.Range).Interior.Color = awinSettings.AmpelGelb
                                            anzSubstituted = anzSubstituted + 1
                                        ElseIf PhaseDefinitions.Contains(stdName.Trim) Then
                                            CType(aktivesSheet.Cells(curZeile, colProtocol + 6), Excel.Range).Interior.Color = awinSettings.AmpelGruen
                                            anzCorrect = anzCorrect + 1
                                        Else
                                            CType(aktivesSheet.Cells(curZeile, colProtocol + 6), Excel.Range).Interior.Color = awinSettings.AmpelNichtBewertet
                                        End If

                                        ' nur wenn es aufgenommen ist, sollte es in die Hierarchie aufgenommen werden 
                                        Try
                                            pHierarchy.add(cphase, elemID, indentLevel)
                                        Catch ex As Exception

                                        End Try

                                    Else

                                        CType(aktivesSheet.Cells(curZeile, colProtocol + 6), Excel.Range).Interior.Color = awinSettings.AmpelRot
                                        CType(aktivesSheet.Cells(curZeile, colProtocol + 7), Excel.Range).Value = logMessage
                                        lastDuplicateIndent = indentLevel

                                        anzIgnored = anzIgnored + 1

                                    End If


                                ElseIf duration = 1 Then
                                    ' hier kommt die Behandlung eines Meilensteins


                                    Try

                                        Dim bewertungsAmpel As Integer = 0
                                        Dim explanation As String = ""

                                        ' hole die Parentphase
                                        cphase = pHierarchy.getPhaseBeforeLevel(indentLevel)
                                        cmilestone = New clsMeilenstein(parent:=cphase)
                                        cbewertung = New clsBewertung


                                        ' damit Kriterien auch eingelesen werden, wenn noch keine Bewertung existiert ...
                                        With cbewertung
                                            '.bewerterName = resultVerantwortlich
                                            .colorIndex = bewertungsAmpel
                                            .datum = Date.Now
                                            .description = explanation
                                        End With


                                        parentElemName = cphase.name
                                        ' jetzt den tatsächlichen Namen bestimmen , ggf wird dazu der Parent Phase Name benötigt 

                                        Try
                                            If Not MilestoneDefinitions.Contains(itemName) Then
                                                stdName = milestoneMappings.mapToStdName(parentElemName, itemName)
                                            Else
                                                stdName = itemName
                                            End If

                                        Catch ex As Exception
                                            stdName = itemName
                                        End Try

                                        Dim ok1 As Boolean = True

                                        If awinSettings.eliminateDuplicates And hproj.hierarchy.containsKey(calcHryElemKey(stdName, True)) Then
                                            ' nur dann kann es Duplikate geben 
                                            Dim duplicateSiblingID As String = hproj.getDuplicateMsSiblingID(stdName, cphase.nameID, _
                                                                                                                 itemStartDate, 0)

                                            If duplicateSiblingID = "" Then
                                                ok1 = True
                                            Else
                                                ok1 = False
                                                logMessage = stdName & " ist Duplikat zu Geschwister " & elemNameOfElemID(duplicateSiblingID) & _
                                                             " und wird ignoriert "
                                            End If

                                        End If


                                        ' jetzt muss geprüft werden, ob stdName bereits aufgenommen ist
                                        Dim ok2 As Boolean = True
                                        If Not MilestoneDefinitions.Contains(stdName) And ok1 Then

                                            Dim hMilestoneDef As New clsMeilensteinDefinition

                                            With hMilestoneDef
                                                .name = stdName
                                                .belongsTo = parentElemName
                                                .shortName = txtAbbrev
                                                .darstellungsKlasse = txtVorgangsKlasse
                                                .UID = milestoneIX
                                            End With

                                            milestoneIX = milestoneIX + 1

                                            If isVorlage And awinSettings.alwaysAcceptTemplateNames Then
                                                ' in die Milestone-Definitions aufnehmen 
                                                Try
                                                    MilestoneDefinitions.Add(hMilestoneDef)
                                                Catch ex As Exception
                                                End Try

                                            Else

                                                logMessage = "ist nicht in der Liste der zugelassenen Elemente enthalten"

                                                ' in die Missing Milestone-Definitions aufnehmen 
                                                Try
                                                    ' das Element aufnehmen, in Abhängigkeit vom Setting 
                                                    If awinSettings.importUnknownNames Then
                                                        ok2 = True
                                                    Else
                                                        ok2 = False
                                                    End If

                                                    missingMilestoneDefinitions.Add(hMilestoneDef)
                                                Catch ex As Exception
                                                End Try
                                            End If


                                        End If

                                        If ok1 And ok2 Then


                                            ' Bestimmung des eindeutigen Namens innerhalb der Geschwister, unterschieden nach Meilenstein und Phase 
                                            If awinSettings.createUniqueSiblingNames Then
                                                stdName = hproj.hierarchy.findUniqueGeschwisterName(cphase.nameID, stdName, True)
                                            End If

                                            elemID = hproj.hierarchy.findUniqueElemKey(stdName, True)


                                            With cmilestone
                                                .nameID = elemID
                                                .setDate = itemEndDate
                                                If Not cbewertung Is Nothing Then
                                                    .addBewertung(cbewertung)
                                                End If
                                            End With

                                            If IsNothing(cphase.getMilestone(cmilestone.nameID)) Then

                                                With cphase
                                                    .addMilestone(cmilestone, origName:=origItem.Trim)
                                                End With

                                                ' Protokollieren
                                                CType(aktivesSheet.Cells(curZeile, colProtocol + 6), Excel.Range).Value = stdName.Trim

                                                ' neuer Breadcrumb 
                                                'Dim PTBreadCrumb As String = pHierarchy.getFootPrint(indentLevel)
                                                Dim PTBreadCrumb As String = hproj.hierarchy.getBreadCrumb(elemID)


                                                If awinSettings.fullProtocol Then

                                                    CType(aktivesSheet.Cells(curZeile, colProtocol + 8), Excel.Range).Value = PTBreadCrumb
                                                    CType(aktivesSheet.Cells(curZeile, colProtocol + 9), Excel.Range).Value = txtVorgangsKlasse
                                                End If

                                                If stdName.Trim <> origItem.Trim Then
                                                    ' es hat eine Ersetzung stattgefunden 
                                                    CType(aktivesSheet.Cells(curZeile, colProtocol + 6), Excel.Range).Interior.Color = awinSettings.AmpelGelb
                                                    anzSubstituted = anzSubstituted + 1
                                                ElseIf MilestoneDefinitions.Contains(stdName.Trim) Then
                                                    CType(aktivesSheet.Cells(curZeile, colProtocol + 6), Excel.Range).Interior.Color = awinSettings.AmpelGruen
                                                    anzCorrect = anzCorrect + 1
                                                Else
                                                    CType(aktivesSheet.Cells(curZeile, colProtocol + 6), Excel.Range).Interior.Color = awinSettings.AmpelGruen

                                                End If


                                            Else

                                                ' Meilenstein existiert in dieser Phase bereits .... 
                                                CType(aktivesSheet.Cells(curZeile, colProtocol + 7), Excel.Range).Value = _
                                                        stdName.Trim & " existiert bereits: Datum 1: " & cphase.getMilestone(stdName).getDate.ToShortDateString & _
                                                        "   , Datum 2: " & cmilestone.getDate.ToShortDateString

                                            End If
                                        Else

                                            CType(aktivesSheet.Cells(curZeile, colProtocol + 7), Excel.Range).Value = logMessage
                                            CType(aktivesSheet.Cells(curZeile, colProtocol + 6), Excel.Range).Interior.Color = awinSettings.AmpelRot
                                            anzIgnored = anzIgnored + 1

                                        End If


                                    Catch ex As Exception
                                        CType(aktivesSheet.Cells(curZeile, colProtocol + 7), Excel.Range).Value = _
                                                            "Fehler in Zeile " & zeile & ", Item-Name: " & itemName
                                        CType(aktivesSheet.Cells(curZeile, colProtocol + 6), Excel.Range).Interior.Color = awinSettings.AmpelRot
                                    End Try


                                End If

                            Else
                                CType(aktivesSheet.Cells(curZeile, colProtocol + 7), Excel.Range).Value = logMessage
                                CType(aktivesSheet.Cells(curZeile, colProtocol + 6), Excel.Range).Interior.Color = awinSettings.AmpelRot
                                anzIgnored = anzIgnored + 1
                            End If

                        Next


                        If Not isVorlage Then

                            ' das ist BMW spezifisch und wird jetzt de-aktiviert .... 
                            'Try
                            '    Dim sopDate As Date = hproj.getMilestone("SOP").getDate

                            '    If DateDiff(DateInterval.Month, StartofCalendar, sopDate) > 0 Then
                            '        Dim sopMonth As Integer = sopDate.Month
                            '        If sopMonth >= 3 And sopMonth <= 6 Then
                            '            anlaufKennung = "03"
                            '        ElseIf sopMonth >= 7 And sopMonth <= 10 Then
                            '            anlaufKennung = "07"
                            '        Else
                            '            anlaufKennung = "11"
                            '        End If
                            '    Else
                            '        anlaufKennung = "?"
                            '    End If

                            'Catch ex As Exception
                            '    anlaufKennung = "?"
                            'End Try

                            ' jetzt wird die Vorlagen Kennung bestimmt 
                            'Dim tstphase As clsPhase = Nothing
                            'Dim relNr As String
                            'tstphase = hproj.getPhase("Systemgestaltung")

                            'If IsNothing(tstphase) Then
                            '    tstphase = hproj.getPhase("I500")
                            '    If IsNothing(tstphase) Then
                            '        tstphase = hproj.getPhase("I300")
                            '        If IsNothing(tstphase) Then
                            '            relNr = "rel 4 "
                            '        Else
                            '            relNr = "rel 5 "
                            '        End If
                            '    Else
                            '        relNr = "rel 5 "
                            '    End If
                            'Else
                            '    relNr = "rel 5 "
                            'End If

                            'vorlagenName = relNr & typKennung & "-" & anlaufKennung
                            'Try
                            '    vorlagenName = vorlagenName.Trim
                            'Catch ex As Exception
                            '    vorlagenName = "unknown"
                            'End Try

                            'If Projektvorlagen.Contains(vorlagenName) Then
                            '    hproj.VorlagenName = vorlagenName
                            'Else
                            '    hproj.VorlagenName = vorlagenName & "*"
                            'End If

                            vorlagenName = ""
                            'If Projektvorlagen.Count >= 1 Then
                            '    vorlagenName = Projektvorlagen.getProject(0).VorlagenName
                            '    hproj.VorlagenName = vorlagenName
                            'End If

                        End If

                        Try

                            If isVorlage Then
                                hproj.farbe = projektFarbe
                                hproj.Schrift = schriftGroesse
                                hproj.Schriftfarbe = schriftfarbe
                            Else

                                If Projektvorlagen.Contains(vorlagenName) Then
                                    vproj = Projektvorlagen.getProject(vorlagenName)

                                    hproj.farbe = vproj.farbe
                                    hproj.Schrift = vproj.Schrift
                                    hproj.Schriftfarbe = vproj.Schriftfarbe
                                    hproj.earliestStart = vproj.earliestStart
                                    hproj.latestStart = vproj.latestStart

                                    'ElseIf Projektvorlagen.Contains("unknown") Then
                                    '    vproj = Projektvorlagen.getProject("unknown")
                                Else
                                    'Throw New Exception("es gibt weder die Vorlage 'unknown' noch die Vorlage " & vorlagenName)
                                    'hproj.farbe = awinSettings.AmpelNichtBewertet
                                    Try
                                        hproj.farbe = CInt(iProjektFarbe)
                                    Catch ex As Exception
                                        hproj.farbe = awinSettings.AmpelNichtBewertet
                                    End Try
                                    hproj.Schrift = Projektvorlagen.getProject(0).Schrift
                                    hproj.Schriftfarbe = RGB(10, 10, 10)
                                    hproj.earliestStart = 0
                                    hproj.latestStart = 0

                                End If




                            End If

                        Catch ex As Exception
                            Throw New Exception(ex.Message)
                        End Try


                        If Not isVorlage And awinSettings.fullProtocol Then

                            ' jetzt werden Projekt-Name, Business Unit und Vorlagen-Kennung weggeschreiben 
                            CType(aktivesSheet.Cells(anfang - 1, colProtocol - 3), Excel.Range).Value = hproj.name
                            CType(aktivesSheet.Cells(anfang - 1, colProtocol - 2), Excel.Range).Value = hproj.VorlagenName
                            CType(aktivesSheet.Cells(anfang - 1, colProtocol - 1), Excel.Range).Value = hproj.businessUnit
                        End If

                        ' jetzt muss das Projekt eingetragen werden 
                        ' ####################################################################
                        ' prüfen ob das Projekt bereits in Session oder Datenbank existiert 

                        vglName = calcProjektKey(hproj.name, hproj.variantName)
                        ' 
                        If ImportProjekte.Containskey(vglName) Then

                            ' dann existiert es bereits in der Session

                            vglProj = ImportProjekte.getProject(vglName)
                            If IsNothing(vglProj) Then
                                ' dieser Fall kann eigentlich gar nicht auftreten ... ? 
                                Call MsgBox("Fehler mit " & vglName)

                            Else
                                ' prüfen, ob es unterschiedlich ist; 
                                ' wenn ja , dann wird es unter dem Varianten Namen Datei-Name angelegt
                                ' wenn der auch schon existiert, dann Fehler udn nichts anlegen ...
                                Dim unterschiede As Collection = hproj.listOfDifferences(vglProj, True, 0)

                                If unterschiede.Count > 0 Then
                                    '' '' '' es gibt Unterschiede, also muss eine Variante angelegt werden 
                                    If hproj.variantName <> variantenName Then
                                        hproj.variantName = variantenName
                                        vglName = calcProjektKey(hproj.name, hproj.variantName)

                                        ' wenn die Variante bereits in der Session existiert ..
                                        ' wird die bisherige gelöscht , die neue über ImportProjekte neu aufgenommen  
                                        If AlleProjekte.Containskey(vglName) Then
                                            AlleProjekte.Remove(vglName)
                                        End If

                                    Else
                                        ' in diesem Fall wird die Variante über hproj neu angelegt 
                                        AlleProjekte.Remove(vglName)
                                    End If

                                    Call replaceProjectVariant(hproj.name, hproj.variantName, False, True, hproj.tfZeile)

                                    Try
                                        myCollection.Add(vglName, vglName)
                                    Catch ex As Exception

                                    End Try

                                Else
                                    ' Projekt in der Form existiert bereits , keine Neu-Anlage
                                    ' es muss sichergestellt sein, dass es angezeigt wird und die Portfolio Definition entsprechend angepasst wird 
                                    ok = False
                                    hproj = vglProj

                                    Call replaceProjectVariant(hproj.name, hproj.variantName, False, True, hproj.tfZeile)

                                    Try
                                        myCollection.Add(vglName, vglName)
                                    Catch ex As Exception

                                    End Try
                                End If
                            End If


                        End If



                        If Not ImportProjekte.Containskey(calcProjektKey(hproj)) Then
                            ImportProjekte.Add(hproj, False)
                            myCollection.Add(calcProjektKey(hproj))

                        End If

                        zeile = ende + 1

                    End If

                End While

                ' jetzt wird die Statistik geschreiben ....
                'CType(activeWSListe.Cells(1, colProtocol + 10), Excel.Range).Value = "Anzahl Insgesamt"
                'CType(activeWSListe.Cells(2, colProtocol + 10), Excel.Range).Value = anzProcessedElements

                'CType(activeWSListe.Cells(1, colProtocol + 11), Excel.Range).Value = "Original Namen"
                'CType(activeWSListe.Cells(2, colProtocol + 11), Excel.Range).Value = anzCorrect

                'CType(activeWSListe.Cells(1, colProtocol + 12), Excel.Range).Value = "Korrigierte Namen"
                'CType(activeWSListe.Cells(2, colProtocol + 12), Excel.Range).Value = anzSubstituted

                'CType(activeWSListe.Cells(1, colProtocol + 13), Excel.Range).Value = "Ignorierte Namen"
                'CType(activeWSListe.Cells(2, colProtocol + 13), Excel.Range).Value = anzIgnored

                '
                ' jetzt werden die Missing Phase- und Milestone Definitions noch weggeschrieben 
                '

                ' aber nur, wenn awinSettings.fullProtokoll = true 


                If awinSettings.fullProtocol Then


                    Dim tmpzeile As Integer
                    tmpzeile = 1

                    Dim wsName As String = "unbekannte Phasen"
                    Dim txtrange As Excel.Range
                    Dim tmpWS As Excel.Worksheet

                    If missingPhaseDefinitions.Count > 0 Then
                        Try
                            tmpWS = CType(appInstance.ActiveWorkbook.Worksheets(wsName), Excel.Worksheet)
                            With tmpWS
                                txtrange = .Range(.Cells(1, 1), .Cells(5000, 8))
                            End With
                            txtrange.Clear()
                        Catch ex As Exception
                            tmpWS = CType(appInstance.ActiveWorkbook.Worksheets.Add(After:=aktivesSheet), Excel.Worksheet)
                            tmpWS.Name = wsName
                        End Try


                        CType(tmpWS.Cells(tmpzeile, 1), Excel.Range).Value = "Phasen-Name"
                        CType(tmpWS.Cells(tmpzeile, 6), Excel.Range).Value = "Abkürzung"
                        CType(tmpWS.Cells(tmpzeile, 7), Excel.Range).Value = "Darstellungsklasse"


                        Dim phDef As clsPhasenDefinition
                        For i = 1 To missingPhaseDefinitions.Count

                            phDef = missingPhaseDefinitions.getPhaseDef(i)
                            CType(tmpWS.Cells(tmpzeile + i, 1), Excel.Range).Value = phDef.name
                            CType(tmpWS.Cells(tmpzeile + i, 6), Excel.Range).Value = phDef.shortName
                            CType(tmpWS.Cells(tmpzeile + i, 7), Excel.Range).Value = phDef.darstellungsKlasse

                        Next
                    End If



                    '
                    ' jetzt werden die Missing Milestone Definitions noch weggeschrieben 
                    '
                    If missingMilestoneDefinitions.Count > 0 Then

                        tmpzeile = 1

                        wsName = "unbekannte Meilensteine"

                        Try
                            tmpWS = CType(appInstance.ActiveWorkbook.Worksheets(wsName), Excel.Worksheet)
                            With tmpWS
                                txtrange = .Range(.Cells(1, 1), .Cells(5000, 8))
                            End With
                            txtrange.Clear()
                        Catch ex As Exception
                            tmpWS = CType(appInstance.ActiveWorkbook.Worksheets.Add(After:=aktivesSheet), Excel.Worksheet)
                            tmpWS.Name = wsName
                        End Try


                        CType(tmpWS.Cells(tmpzeile, 1), Excel.Range).Value = "Meilenstein-Name"
                        CType(tmpWS.Cells(tmpzeile, 5), Excel.Range).Value = "Bezug"
                        CType(tmpWS.Cells(tmpzeile, 6), Excel.Range).Value = "Abkürzung"
                        CType(tmpWS.Cells(tmpzeile, 7), Excel.Range).Value = "Darstellungsklasse"


                        Dim msDef As clsMeilensteinDefinition
                        For i = 1 To missingMilestoneDefinitions.Count

                            msDef = missingMilestoneDefinitions.getMilestoneDef(i)
                            If Not IsNothing(msDef) Then
                                CType(tmpWS.Cells(tmpzeile + i, 1), Excel.Range).Value = msDef.name
                                CType(tmpWS.Cells(tmpzeile + i, 5), Excel.Range).Value = msDef.belongsTo
                                CType(tmpWS.Cells(tmpzeile + i, 6), Excel.Range).Value = msDef.shortName
                                CType(tmpWS.Cells(tmpzeile + i, 7), Excel.Range).Value = msDef.darstellungsKlasse
                            End If


                        Next


                    End If

                End If

                If appInstance.ActiveSheet.name <> aktivesSheet.Name Then
                    aktivesSheet.Activate()
                End If

            End With
        Catch ex As Exception
            Throw New Exception("Fehler in Datei " & vbLf & ex.Message & vbLf & _
                                 currentDateiName & vbLf)
        End Try


    End Sub

    ''' <summary>
    ''' exportiert das angegebene Projekt in die bereits geöffnete Datei 
    ''' Das Schreiben beginnt ab "zeile"
    ''' </summary>
    ''' <param name="hproj"></param>
    ''' <param name="zeile"></param>
    ''' <remarks></remarks>
    Public Sub planExportProject(ByVal hproj As clsProjekt, ByRef zeile As Integer)

        Dim ip As Integer, im As Integer
        Dim startdate As Date, endDate As Date
        Dim curName As String
        Dim color As Long
        Dim ws As Excel.Worksheet
        Dim spalte As Integer = 1
        Dim cphase As clsPhase
        Dim cmilestone As clsMeilenstein
        Dim indentlevel As Integer = 0
        Dim indentDelta As Integer = 3

        ' diese Datei muss offen sein und das aktive Workbook
        ' wenn nein, dann aktivieren ! 
        Try
            If appInstance.ActiveWorkbook.Name <> excelExportVorlage Then
                appInstance.Workbooks(excelExportVorlage).Activate()
            End If
        Catch ex As Exception
            Throw New ArgumentException("Export Vorlage ist nicht die aktive Excel Datei")
        End Try

        ' bestimme die Farbe - sie steht im Excel Ausgabe File in der Zeile 2, Spalte 1 
        ws = CType(appInstance.ActiveWorkbook.Worksheets("Export VISBO Projekttafel"), Excel.Worksheet)

        Dim suchstr(8) As String
        suchstr(ptPlanNamen.Name) = "Name"
        suchstr(ptPlanNamen.Anfang) = "Start"
        suchstr(ptPlanNamen.Ende) = "End"
        suchstr(ptPlanNamen.Beschreibung) = "Description"
        suchstr(ptPlanNamen.Vorgangsklasse) = "Appearance"
        suchstr(ptPlanNamen.BusinessUnit) = "Business Unit"
        suchstr(ptPlanNamen.Protocol) = "Übernommen als"
        suchstr(ptPlanNamen.Dauer) = "Duration"
        suchstr(ptPlanNamen.Abkuerzung) = "Abbreviation"

        ' jetzt werden die Spaltenüberschriften geschrieben 
        Dim colName As Integer = 1
        CType(ws.Cells(1, colName), Excel.Range).Value = suchstr(ptPlanNamen.Name)
        Dim colStart As Integer = 2
        CType(ws.Cells(1, colStart), Excel.Range).Value = suchstr(ptPlanNamen.Anfang)
        Dim colEnde As Integer = 3
        CType(ws.Cells(1, colEnde), Excel.Range).Value = suchstr(ptPlanNamen.Ende)
        Dim colBU As Integer = 4
        CType(ws.Cells(1, colBU), Excel.Range).Value = suchstr(ptPlanNamen.BusinessUnit)
        Dim colDescription As Integer = 5
        CType(ws.Cells(1, colDescription), Excel.Range).Value = suchstr(ptPlanNamen.Beschreibung)
        Dim colAppearance As Integer = 6
        CType(ws.Cells(1, colAppearance), Excel.Range).Value = suchstr(ptPlanNamen.Vorgangsklasse)
        Dim colAbbrev As Integer = 7
        CType(ws.Cells(1, colAbbrev), Excel.Range).Value = suchstr(ptPlanNamen.Abkuerzung)

        color = CLng(CType(ws.Cells(2, 1), Excel.Range).Interior.Color)

        ' jetzt wird das Projekt geschrieben 
        CType(ws.Cells(zeile, colName), Excel.Range).Value = hproj.getShapeText
        CType(ws.Cells(zeile, colStart), Excel.Range).Value = hproj.startDate.ToShortDateString
        CType(ws.Cells(zeile, colEnde), Excel.Range).Value = hproj.endeDate.ToShortDateString
        CType(ws.Cells(zeile, colBU), Excel.Range).Value = hproj.businessUnit
        CType(ws.Cells(zeile, colDescription), Excel.Range).Value = hproj.description
        CType(ws.Rows(zeile), Excel.Range).Interior.Color = color

        Dim indentPhase As String = "   "
        'Dim indentMS As String = "      "

        ' die erste Phase kann auch Meilensteine haben !
        cphase = hproj.getPhase(1)
        indentlevel = hproj.hierarchy.getIndentLevel(cphase.nameID)

        For im = 1 To cphase.countMilestones
            zeile = zeile + 1
            cmilestone = cphase.getMilestone(im)
            startdate = cmilestone.getDate
            
            curName = cmilestone.name

            indentlevel = hproj.hierarchy.getIndentLevel(cmilestone.nameID)
            CType(ws.Cells(zeile, colName), Excel.Range).Value = erzeugeIndent(indentlevel) & curName

            If DateDiff(DateInterval.Day, StartofCalendar, startdate) > 0 Then
                CType(ws.Cells(zeile, colStart), Excel.Range).Value = ""
                CType(ws.Cells(zeile, colEnde), Excel.Range).Value = startdate.ToShortDateString
            Else
                CType(ws.Cells(zeile, colStart), Excel.Range).Value = "Fehler !"
                CType(ws.Cells(zeile, colEnde), Excel.Range).Value = "Fehler !"
            End If

            ' jetzt Vorgangsklasse und Abbrev schreiben, falls vorhanden 
            Dim tmpAbbrev As String = MilestoneDefinitions.getAbbrev(curName)
            Dim tmpAppearance As String = MilestoneDefinitions.getAppearance(curName)

            CType(ws.Cells(zeile, colAbbrev), Excel.Range).Value = tmpAbbrev
            CType(ws.Cells(zeile, colAppearance), Excel.Range).Value = tmpAppearance

        Next



        For ip = 2 To hproj.AllPhases.Count
            zeile = zeile + 1
            cphase = hproj.getPhase(ip)
            startdate = cphase.getStartDate
            endDate = cphase.getEndDate
            curName = cphase.name

            indentlevel = hproj.hierarchy.getIndentLevel(cphase.nameID)
            CType(ws.Cells(zeile, spalte), Excel.Range).Value = erzeugeIndent(indentlevel) & curName

            If DateDiff(DateInterval.Day, StartofCalendar, startdate) > 0 Then
                CType(ws.Cells(zeile, colStart), Excel.Range).Value = startdate.ToShortDateString
            Else
                CType(ws.Cells(zeile, colStart), Excel.Range).Value = "Fehler !"
            End If

            If DateDiff(DateInterval.Day, StartofCalendar, endDate) > 0 Then
                CType(ws.Cells(zeile, colEnde), Excel.Range).Value = endDate.ToShortDateString
            Else
                CType(ws.Cells(zeile, colEnde), Excel.Range).Value = "Fehler !"
            End If

            ' jetzt Vorgangsklasse und Abbrev schreiben, falls vorhanden 
            Dim tmpAbbrev As String = PhaseDefinitions.getAbbrev(curName)
            Dim tmpAppearance As String = PhaseDefinitions.getAppearance(curName)

            CType(ws.Cells(zeile, colAbbrev), Excel.Range).Value = tmpAbbrev
            CType(ws.Cells(zeile, colAppearance), Excel.Range).Value = tmpAppearance

            For im = 1 To cphase.countMilestones
                zeile = zeile + 1
                cmilestone = cphase.getMilestone(im)
                startdate = cmilestone.getDate
                

                curName = cmilestone.name
                indentlevel = hproj.hierarchy.getIndentLevel(cmilestone.nameID)
                CType(ws.Cells(zeile, spalte), Excel.Range).Value = erzeugeIndent(indentlevel) & curName

                If DateDiff(DateInterval.Day, StartofCalendar, startdate) > 0 Then
                    CType(ws.Cells(zeile, colStart), Excel.Range).Value = ""
                    CType(ws.Cells(zeile, colEnde), Excel.Range).Value = startdate.ToShortDateString
                Else
                    CType(ws.Cells(zeile, colStart), Excel.Range).Value = "Fehler !"
                    CType(ws.Cells(zeile, colEnde), Excel.Range).Value = "Fehler !"
                End If

                ' jetzt Vorgangsklasse und Abbrev schreiben, falls vorhanden 
                tmpAbbrev = MilestoneDefinitions.getAbbrev(curName)
                tmpAppearance = MilestoneDefinitions.getAppearance(curName)

                CType(ws.Cells(zeile, colAbbrev), Excel.Range).Value = tmpAbbrev
                CType(ws.Cells(zeile, colAppearance), Excel.Range).Value = tmpAppearance
            Next

        Next

        ' jetzt muss um eine Zeile weitergeschaltet werden, damit immer auf eine freie Zeile geschrieben wird
        zeile = zeile + 1

    End Sub


    ''' <summary>
    ''' Einlesen eines RXF-Files (XML-Ausleitung von RPLAN) und dazu ein Protokoll in Tabellenblatt 'xmlfilename'protokoll in Datei Logfile
    ''' </summary>
    ''' <param name="myCollection"></param>
    ''' <param name="xmlfilename"></param>Name des RXF-Files
    ''' <param name="isVorlage"></param>Ist Vorlage, oder nicht
    ''' <remarks></remarks>
    Sub RXFImport(ByRef myCollection As Collection, ByVal xmlfilename As String, _
                  ByVal isVorlage As Boolean, ByRef protokollliste As SortedList(Of Integer, clsProtokoll))
        ' akt. Name zum Zweck des Fehlersuchens
        Dim aktuellerName As String = ""

        'Variablen-Definitionen für Projectboard 

        Dim hproj As clsProjekt

        Dim vproj As clsProjektvorlage
        Dim vorlagenName As String = ""

        Dim ProjektdauerinDays As Integer = 0
        Dim cphase As clsPhase = Nothing

        Dim parentphase As clsPhase = Nothing
        Dim lastphase As clsPhase = Nothing

        Dim parentelemID As String = ""
        Dim lastelemID As String = ""

        Dim cBewertung As clsBewertung = Nothing

        Dim milestoneName As String = ""

        ' Ersetzen eines bestimmten Strings in der kompletten Datei 'xmlfilename'
        ' Zurückgeben des Namens der neuen Datei 'newXMLfilename'

        Dim newXMLfilename As String = replaceStringInFile(xmlfilename, "xsi:type=""subscribedtask""", "")



        ' XML-Datei Öffnen
        ' A FileStream is needed to read the XML document.
        Dim fs As New FileStream(newXMLfilename, FileMode.Open)

        ' Declare an object variable of the type to be deserialized.
        Dim Rplan As New rxf            ' Class rxf wird in clsRplanRXF.vb definiert

        Try

            ' Create an instance of the XmlSerializer class;
            ' specify the type of object to be deserialized.
            Dim deserializer As New XmlSerializer(GetType(rxf), "http://www.actano.de/2007/rxf")


            ' If the XML document has been altered with unknown
            ' nodes or attributes, handle them with the
            ' UnknownNode and UnknownAttribute events.

            ' Änderung tk: die beiden deserializer Kommandos müssen wieder aktiviert werden !
            'Call MsgBox("hier wurde RXF Import massgeblich verändert !!" & vbLf & _
            '             " lief bei Windows 10/Excel 2016 nicht")
            AddHandler deserializer.UnknownNode, AddressOf deserializer_UnknownNode
            AddHandler deserializer.UnknownAttribute, AddressOf deserializer_UnknownAttribute


            ' Einlesen des kompletten XML-Dokument im die Klasse rxf
            ' Use the Deserialize method to restore the object's state with
            ' data from the XML document. 
            Rplan = CType(deserializer.Deserialize(fs), rxf)

            ' Tabellenblatt "xmlfilename" im logfile.xlsx erzeugen fürs Protokoll (xmlfilename ohne ".rxf" Extension)

            Dim tstr As String() = Split(xmlfilename, "\", -1)
            Dim hstr As String = tstr(tstr.Length - 1)
            Dim quelle As String = hstr
            tstr = Split(hstr, ".", 2)

            Dim tabblattname As String = tstr(0)
            Dim wslogbuch As Excel.Worksheet = Nothing


            Dim protokollLine As New clsProtokoll("", quelle)
            Dim zeile As Integer = 3


            ' Projekt suchen; VISBO Projekt suchen unter der RPLANTasks mit gegebenen MainProject
            For i = 0 To Rplan.task.Length - 1

                If Not IsNothing(Rplan.task(i).mainProject) Then
                    ' akt. Task ist Projekt 

                    aktuellerName = Rplan.task(i).name

                    Dim aktTask_i As rxfTask = Rplan.task(i)
                    hproj = New clsProjekt

                    hproj.name = aktTask_i.name
                    hproj.VorlagenName = ""
                    hproj.leadPerson = aktTask_i.owner
                    hproj.startDate = aktTask_i.actualDate.start.Value
                    ProjektdauerinDays = calcDauerIndays(aktTask_i.actualDate.start.Value, aktTask_i.actualDate.finish.Value)

                    ' Protokollzeile bestücken
                    protokollLine = New clsProtokoll(hproj.name, quelle)


                    ' ProjektPhase erzeugen
                    cphase = New clsPhase(parent:=hproj)

                    With cphase
                        .nameID = rootPhaseName

                        Dim Duration As Long = calcDauerIndays(aktTask_i.actualDate.start.Value, aktTask_i.actualDate.finish.Value)
                        Dim offset As Integer = DateDiff(DateInterval.Day, hproj.startDate, aktTask_i.actualDate.start.Value)

                        ' für die rootPhase muss gelten: offset = startoffset = 0 und duration = ProjektdauerIndays

                        Dim startOffset As Integer = 0
                        .changeStartandDauer(startOffset, Duration)
                        Dim phaseStartdate As Date = .getStartDate
                        Dim phaseEnddate As Date = .getEndDate

                    End With

                    ' ProjektPhase wird hinzugefügt
                    Dim hrchynode As New clsHierarchyNode
                    hrchynode.elemName = cphase.name
                    hrchynode.parentNodeKey = ""
                    hproj.AddPhase(cphase, parentID:=hrchynode.parentNodeKey)
                    parentphase = cphase
                    parentelemID = cphase.nameID
                    lastphase = cphase
                    lastelemID = cphase.nameID

                    ' Alle Tasks zu diesem Projekt mit deren Kinder und KindesKinder in hproj eintragen
                    Try
                        Call findAllTasksandInsert(aktTask_i, parentelemID, hproj, Rplan, protokollLine, zeile, protokollliste)
                    Catch ex As Exception
                        Dim a As Integer = 0
                    End Try


                    '
                    '' '' Bestimmung der BMW-Vorlage des jeweiligen Projektes
                    '' '' muss noch genauer herausgefunden werden, welche Vorlage für das jeweilige Projekt verwendet werden muss
                    '
                    vorlagenName = findBMWVorlagenName(hproj)
                    '
                    ''''    Ende Bestimmung der BMW-Vorlage zu diesem Projekt
                    '


                    If Projektvorlagen.Contains(vorlagenName) Then
                        vproj = Projektvorlagen.getProject(vorlagenName)

                        hproj.VorlagenName = vorlagenName
                        hproj.farbe = vproj.farbe
                        hproj.Schrift = vproj.Schrift
                        hproj.Schriftfarbe = vproj.Schriftfarbe
                        hproj.earliestStart = vproj.earliestStart
                        hproj.latestStart = vproj.latestStart

                        'ElseIf Projektvorlagen.Contains("unknown") Then
                        '    vproj = Projektvorlagen.getProject("unknown")
                    Else
                        'Throw New Exception("es gibt weder die Vorlage 'unknown' noch die Vorlage " & vorlagenName)
                        hproj.VorlagenName = ""
                        hproj.farbe = awinSettings.AmpelNichtBewertet
                        hproj.Schrift = Projektvorlagen.getProject(0).Schrift
                        hproj.Schriftfarbe = RGB(10, 10, 10)
                        hproj.earliestStart = 0
                        hproj.latestStart = 0

                    End If

                    Dim msPHdefcount As Integer = missingPhaseDefinitions.Count
                    Dim msMSdefcount As Integer = missingMilestoneDefinitions.Count

                    ' jetzt muss das Projekt eingetragen werden in die Listen Importierte Projekte und myCollection
                    ' Änderung tk: falls es das Projekt unter diesem Namen bereits gibt, wird eine Variante angelegt ... 
                    Dim lfdNr As Integer = 2
                    Do While ImportProjekte.Containskey(calcProjektKey(hproj))
                        hproj.variantName = lfdNr.ToString
                        lfdNr = lfdNr + 1
                    Loop

                    Dim hlptxt As String = ""
                    If lfdNr - 2 > 0 Then
                        If lfdNr - 2 = 1 Then
                            hlptxt = "es wurde eine Variante angelegt"
                        Else
                            hlptxt = "es wurden " & lfdNr - 2 & " Varianten angelegt."
                        End If
                        Call MsgBox("Projekt " & hproj.name & " kommt mehrmals vor! " & vbLf & hlptxt)
                    End If


                    ' jetzt ist sichergestellt, dass calcProjektKey nicht mehr vorkommt 
                    ImportProjekte.Add(hproj, False)
                    myCollection.Add(calcProjektKey(hproj))


                Else
                    ' aktuelle Task ist kein Projekt
                End If
            Next i

            '' '' Protokolldatei sichern
            ' ''Call writeProtokoll(protokollliste, tabblattname)


            ' RXF-Datei (entspricht XML-Datei) Schliessen
            fs.Close()

        Catch ex As Exception
            Call logfileSchreiben(ex.Message & vbLf & "Fehler bei Name " & CStr(aktuellerName), aktuellerName, anzFehler)
            Throw New ArgumentException("Fehler bei Name " & CStr(aktuellerName))

            ' RXF-Datei (entspricht XML-Datei) Schliessen
            fs.Close()
        End Try


    End Sub


    ''' <summary>
    ''' sucht zu der rxfTask 'task' alle Kinder und KindesKinder  und trägt diese in das Projekt 'hproj' ein 
    ''' dazu wird diese Routine rekursiv aufgerufen
    ''' </summary>
    ''' <param name="task"></param>rxfTask 'task', die Parent aller gesuchten Tasks ist
    ''' <param name="parentelemID"></param>Parent dieser rxfTask 'task'
    ''' <param name="hproj"></param>aktuelles aufzubauendes Projekt
    ''' <param name="RPLAN"></param>Komplette eingelesene rxf-Struktur 
    ''' <remarks></remarks>
    Private Sub findAllTasksandInsert(ByVal task As rxfTask, ByVal parentelemID As String, ByRef hproj As clsProjekt, ByVal RPLAN As rxf, ByRef prtLine As clsProtokoll, ByRef zeile As Integer, ByRef prtliste As SortedList(Of Integer, clsProtokoll))


        Dim cphase As clsPhase = Nothing
        Dim cmilestone As clsMeilenstein
        Dim parentphase As clsPhase = hproj.getPhaseByID(parentelemID)
        Dim lastphase As clsPhase = Nothing
        Dim lastelemID As String = ""

        Dim phaseNameID As String = ""
        Dim cBewertung As clsBewertung = Nothing

        Dim origMSname As String = ""
        Dim milestonedate As Date
        Dim isNotDuplikate As Boolean = True
        Dim isUnkownName As Boolean = False


        ' weitere Tasks finden, die zu diesem Projekt (mit ID=aktTask.id) gehören, d.h. ID muss als Parent auftreten
        For j = 0 To RPLAN.task.Length - 1

            isUnkownName = False            ' hier ist noch unklar, ob kown oder unkown Task

            If RPLAN.task(j).parent = task.id Then

                ' Änderung tk am 10.2.16 - wenn ein Fehler bei einem einzelnen Element auftritt, soll das nicht dazu führen, 
                ' dass der Import aller anderen abgebrochen wird ... eine entsprechende Fehlermeldung soll ins Protokoll kommen 
                ' alle anderen Elemente sollen importiert werden 
                Try

                    Dim aktTask_j As rxfTask = RPLAN.task(j)
                    Dim isMilestone As Boolean

                    Dim isKnownMsName As Boolean = MilestoneDefinitions.Contains(aktTask_j.name) Or _
                                                missingMilestoneDefinitions.Contains(aktTask_j.name)

                    Dim isKnownPhName As Boolean = PhaseDefinitions.Contains(aktTask_j.name) Or _
                                                missingPhaseDefinitions.Contains(aktTask_j.name)

                    Dim taskdauerinDays As Long = calcDauerIndays(aktTask_j.actualDate.start.Value, aktTask_j.actualDate.finish.Value)
                    ' Herausfinden, ob aktTask_j Phase oder Meilenstein ist 

                    If taskdauerinDays > 1 Then
                        isMilestone = False

                        If aktTask_j.taskType.type = "MILESTONE" Then
                            Call logfileSchreiben("Korrektur, RXFImport: Phasen-Element mit verschiedenen Start- und Ende-Daten war als Meilenstein deklariert:", _
                                                        aktTask_j.name & ": " & aktTask_j.actualDate.start.Value.ToShortDateString & " versus " & _
                                                        aktTask_j.actualDate.finish.Value.ToShortDateString & vbLf & _
                                                        "Projekt: " & hproj.name, _
                                                        anzFehler)
                        End If

                    ElseIf aktTask_j.taskType.type = "MILESTONE" Then
                        isMilestone = True

                    ElseIf isKnownMsName And Not isKnownPhName Then
                        isMilestone = True
                        If aktTask_j.taskType.type <> "MILESTONE" Then
                            Call logfileSchreiben("Korrektur, RXFImport: bekanntes Meilenstein-Element  mit falscher Typ-Zuordnung:", _
                                                        aktTask_j.name & " mit Typ " & aktTask_j.taskType.type & vbLf & _
                                                        "Projekt: " & hproj.name, _
                                                        anzFehler)
                        End If

                    ElseIf isKnownPhName And Not isKnownMsName Then
                        isMilestone = False


                    Else
                        isMilestone = True
                    End If

                    If Not isMilestone Then

                        ''''''  ist PHASE

                        If aktTask_j.name = "Projektphasen" Then

                            For i = 0 To aktTask_j.customvalue.Length - 1
                                If aktTask_j.customvalue(i).name = "UsA_SERVICE_SPALTE_A" Then
                                    hproj.businessUnit = aktTask_j.customvalue(i).Value
                                End If

                                If aktTask_j.customvalue(i).name = "UsA_SERVICE_SPALTE_B" Then
                                    hproj.VorlagenName = aktTask_j.customvalue(i).Value
                                End If
                            Next i

                        End If

                        ' überprüfen, ob die Phase evt. ignoriert werden soll (wird im  CustomizationFile in Tabelle Phase-Mappings definiert)
                        If Not phaseMappings.tobeIgnored(aktTask_j.name) Then
                            Dim mappedPhasename As String = ""

                            prtLine.planelement = aktTask_j.name
                            prtLine.hgColor = awinSettings.AmpelNichtBewertet

                            If PhaseDefinitions.Contains(aktTask_j.name) Then

                                mappedPhasename = aktTask_j.name
                                prtLine.hgColor = awinSettings.AmpelGruen

                            Else
                                ' aktTask_j.name existiert nicht in den PhaseDefinitions

                                'wenn der PhasenName gemappt werden kann und dieser dann in phasedefinitions enthalten ist, so wird phasename ersetzt
                                mappedPhasename = phaseMappings.mapToStdName(elemNameOfElemID(parentelemID), aktTask_j.name)

                                If PhaseDefinitions.Contains(mappedPhasename) Then
                                    ' neuer aktueller Name der Task
                                    prtLine.hgColor = awinSettings.AmpelGelb

                                Else
                                    ' PhasenName ist nicht bekannt
                                    isUnkownName = True


                                    Dim newPhaseDef As New clsPhasenDefinition


                                    ' Änderung tk 6.12.15: das muss auf den mappedPhasename gesetzt werdne, da sonst Eltern-Ersetzungen, die noch nicht 
                                    ' in der phasedefinitions sind , nicht in der Liste der unbekannten aufgenommen werden ... 
                                    'newPhaseDef.name = aktTask_j.name
                                    'mappedPhasename = aktTask_j.name

                                    newPhaseDef.name = mappedPhasename
                                    newPhaseDef.shortName = aktTask_j.remark

                                    newPhaseDef.darstellungsKlasse = mapToAppearance(aktTask_j.taskType.Value, False)
                                    newPhaseDef.UID = PhaseDefinitions.Count + 1
                                    ' muss in missingPhaseDefinitions noch eingetragen werden
                                    ' in add wird abgefragt, ob der Name schon existiert, wenn ja, wird nix gemacht 
                                    missingPhaseDefinitions.Add(newPhaseDef)

                                    ' Änderung tk: wird auskommentiert, das steht ja im Protokoll
                                    'Call logfileSchreiben(("Achtung, RXFImport: Phase '" & aktTask_j.name & "' existiert im CustomizationFile nicht!"), hproj.name, anzFehler)

                                End If
                            End If

                            ' Phase nur aufnehmen in das aktuelle Projekt, wenn 
                            ' awinSettings.importUnkownNames=true ist oder auch isUnkownName = false

                            If Not isUnkownName Or awinSettings.importUnknownNames Then

                                Dim phaseStartdate As Date
                                Dim phaseEnddate As Date
                                cphase = New clsPhase(hproj)

                                With cphase

                                    Dim Duration As Integer = calcDauerIndays(aktTask_j.actualDate.start.Value, aktTask_j.actualDate.finish.Value)
                                    Dim offset As Integer = DateDiff(DateInterval.Day, hproj.startDate, aktTask_j.actualDate.start.Value)

                                    .changeStartandDauer(offset, Duration)
                                    phaseStartdate = .getStartDate
                                    phaseEnddate = .getEndDate

                                    isNotDuplikate = True
                                    ' sollen Duplikate eliminiert werden ?
                                    If awinSettings.eliminateDuplicates And hproj.hierarchy.containsKey(calcHryElemKey(mappedPhasename, False)) Then
                                        ' nur dann kann es Duplikate geben 
                                        If hproj.isCloneToParent(mappedPhasename, parentphase.nameID, phaseStartdate, phaseEnddate, 0.97) Then
                                            isNotDuplikate = False
                                            prtLine.planelement = aktTask_j.name
                                            prtLine.hgColor = awinSettings.AmpelRot
                                            prtLine.grund = "Phase wurde eliminiert: Duplikat zur Parent-Phase"
                                            'Call logfileSchreiben("Fehler in RXFImport: " & mappedPhasename & " ist Duplikat zu Parent " & parentphase.name & " und wird ignoriert ", hproj.name, anzFehler)

                                        Else
                                            Dim duplicateSiblingID As String = hproj.getDuplicatePhaseSiblingID(mappedPhasename, parentphase.nameID, _
                                                                                                                phaseStartdate, phaseEnddate, 0.97)

                                            If duplicateSiblingID = "" Then
                                                isNotDuplikate = True
                                            Else
                                                isNotDuplikate = False
                                                prtLine.planelement = aktTask_j.name
                                                prtLine.hgColor = awinSettings.AmpelRot
                                                prtLine.grund = "Phase wurde eliminiert: Duplikat zur Geschwister-Phase"
                                                'Call logfileSchreiben(" Fehler in RXFImport: " & mappedPhasename & " ist Duplikat zu Geschwister " & elemNameOfElemID(duplicateSiblingID) & _
                                                '" und wird ignoriert ", hproj.name, anzFehler)
                                            End If
                                        End If

                                    End If

                                End With

                                If isNotDuplikate Then

                                    ' hier muss für gleiche PhasenNamen als Geschwister noch eine lfdNummer angehängt werden
                                    ' es muss überprüft werden, ob es Geschwister mit gleichem Namen gibt:
                                    ' wenn ja, wird an den mappedPhaseName eine LFdNr. ergänzt,bis der Name innerhalb der Geschwistergruppe eindeutig ist.

                                    If awinSettings.createUniqueSiblingNames Then
                                        mappedPhasename = hproj.hierarchy.findUniqueGeschwisterName(parentelemID, mappedPhasename, False)
                                    End If

                                    cphase.nameID = hproj.hierarchy.findUniqueElemKey(mappedPhasename, False)

                                    ' Phase wird ins Projekt mitaufgenommen

                                    Dim phrchynode As New clsHierarchyNode
                                    phrchynode.elemName = cphase.name
                                    phrchynode.parentNodeKey = parentelemID

                                    hproj.AddPhase(cphase, origName:=aktTask_j.name, parentID:=phrchynode.parentNodeKey)
                                    phrchynode.indexOfElem = hproj.AllPhases.Count

                                    ' merken von letzem Element (Knoten,Phase,Meilenstein)
                                    'lasthrchynode = phrchynode
                                    lastelemID = cphase.nameID
                                    lastphase = cphase

                                    prtLine.hierarchie = hproj.hierarchy.getBreadCrumb(cphase.nameID)
                                    prtLine.PThierarchie = hproj.hierarchy.getBreadCrumb(cphase.nameID)
                                    prtLine.planelement = aktTask_j.name
                                    prtLine.abkürzung = PhaseDefinitions.getAbbrev(cphase.name)
                                    prtLine.planeleÜbern = cphase.name

                                    prtLine.klasse = aktTask_j.taskType.Value
                                    prtLine.PTklasse = mapToAppearance(aktTask_j.taskType.Value, False)

                                    prtliste.Add(zeile, prtLine)
                                    zeile = zeile + 1
                                    'prtLine.writeLog(zeile)

                                    Dim quelle As String = prtLine.quelle

                                    prtLine = New clsProtokoll(hproj.name, quelle)
                                    prtLine.actDate = ""


                                    Call findAllTasksandInsert(aktTask_j, lastelemID, hproj, RPLAN, prtLine, zeile, prtliste)

                                Else
                                    prtliste.Add(zeile, prtLine)
                                    zeile = zeile + 1

                                    Dim quelle As String = prtLine.quelle
                                    prtLine = New clsProtokoll(hproj.name, quelle)
                                    prtLine.actDate = ""
                                End If

                            Else
                                prtLine.planelement = aktTask_j.name
                                prtLine.hgColor = awinSettings.AmpelRot
                                prtLine.grund = "Phase wurde ignoriert: unbekannter Bezeichner"

                                prtliste.Add(zeile, prtLine)
                                zeile = zeile + 1

                                Dim quelle As String = prtLine.quelle
                                prtLine = New clsProtokoll(hproj.name, quelle)
                                prtLine.actDate = ""
                            End If

                        Else
                            prtLine.planelement = aktTask_j.name
                            prtLine.hgColor = awinSettings.AmpelRot
                            prtLine.grund = "Phase wurde ignoriert: gemäß Eintrag TOBEIGNORED im Wörterbuch"

                            prtliste.Add(zeile, prtLine) ' Protokollzeile in Liste eintragen
                            zeile = zeile + 1

                            Dim quelle As String = prtLine.quelle
                            prtLine = New clsProtokoll(hproj.name, quelle) ' neue Protokollzeile
                            prtLine.actDate = ""

                        End If       'Ende of tobeignored phase


                    Else
                        ' ist MEILENSTEIN

                        Dim mappedMSname As String = ""

                        If Not milestoneMappings.tobeIgnored(aktTask_j.name) Then

                            If MilestoneDefinitions.Contains(aktTask_j.name) Then

                                mappedMSname = aktTask_j.name
                                prtLine.hgColor = awinSettings.AmpelGruen

                            Else
                                'wenn der MeilensteinName gemappt werden kann und dieser dann in milestonedefinitions enthalten ist, so wird Meilensteinname ersetzt
                                mappedMSname = milestoneMappings.mapToStdName(elemNameOfElemID(parentelemID), aktTask_j.name)
                                If MilestoneDefinitions.Contains(mappedMSname) Then

                                    prtLine.hgColor = awinSettings.AmpelGelb
                                Else

                                    isUnkownName = True

                                    Dim msDef As New clsMeilensteinDefinition


                                    ' Änderung tk 6.12.15: das muss auf den mappedMSNamen gesetzt werdne, da sonst Eltern-Ersetzungen, die noch nicht 
                                    ' in der milestonedefinitions sind , nicht in der Liste der unbekannten aufgenommen werden ... 
                                    'msDef.name = aktTask_j.name
                                    'mappedMSname = aktTask_j.name

                                    msDef.name = mappedMSname
                                    msDef.schwellWert = 0
                                    msDef.belongsTo = parentphase.name
                                    msDef.shortName = aktTask_j.remark

                                    msDef.darstellungsKlasse = mapToAppearance(aktTask_j.taskType.Value, True)
                                    msDef.UID = MilestoneDefinitions.Count + 1

                                    Try
                                        missingMilestoneDefinitions.Add(msDef)

                                        'Call logfileSchreiben(("Achtung, RXFImport: Meilenstein '" & aktTask_j.name & "' existiert im CustomizationFile nicht!"), hproj.name, anzFehler)

                                    Catch ex As Exception
                                    End Try

                                End If
                            End If

                            ' Meilenstein wird nur in das aktuelle Projekt aufgenommen, wenn awinSettings.importUnkownNames = true 
                            ' und der Name bekannt ist (isUnkownName = false)

                            If Not isUnkownName Or awinSettings.importUnknownNames Then

                                cmilestone = New clsMeilenstein(parent:=parentphase)
                                cBewertung = New clsBewertung

                                origMSname = aktTask_j.name

                                If DateDiff(DateInterval.Day, aktTask_j.actualDate.start.Value, aktTask_j.actualDate.finish.Value) = 0 Then
                                    milestonedate = aktTask_j.actualDate.start.Value
                                Else
                                    Throw New Exception("Fehler, RXFImport: Der Meilenstein hat verschiedene Start- und End-Daten:" & vbLf & _
                                                        aktTask_j.actualDate.start.Value.ToShortDateString & " versus " & _
                                                        aktTask_j.actualDate.finish.Value.ToShortDateString & vbLf & _
                                                        "Projekt: " & hproj.name)
                                End If



                                ' wenn der freefloat nicht zugelassen ist und der Meilenstein ausserhalb der Phasen-Grenzen liegt 
                                ' muss abgebrochen werden 

                                If Not awinSettings.milestoneFreeFloat And _
                                    (DateDiff(DateInterval.Day, parentphase.getStartDate, milestonedate) < 0 Or _
                                     DateDiff(DateInterval.Day, parentphase.getEndDate, milestonedate) > 0) Then

                                    'Call logfileSchreiben(("Fehler, RXFImport: Der Meilenstein liegt ausserhalb seiner Phase" & vbLf & _
                                    '                    origMSname & " nicht innerhalb " & parentphase.name & vbLf & _
                                    '                         "Korrigieren Sie bitte diese Inkonsistenz in der Datei '"), hproj.name, anzFehler)
                                    Throw New Exception("Fehler, RXFImport: Der Meilenstein liegt ausserhalb seiner Phase" & vbLf & _
                                                        origMSname & " nicht innerhalb " & parentphase.name & vbLf & _
                                                             "Korrigieren Sie bitte diese Inkonsistenz in der Datei '" & vbLf & hproj.name & ".xlsx'")
                                End If

                                Dim resultVerantwortlich As String = aktTask_j.owner
                                Dim bewertungsAmpel As Integer = 0
                                Dim explanation As String = aktTask_j.note

                                ' Ergänzung tk 2.11 deliverables ergänzt 
                                Dim deliverables As String = ""

                                If bewertungsAmpel < 0 Or bewertungsAmpel > 3 Then
                                    ' es gibt keine Bewertung
                                    bewertungsAmpel = 0
                                End If
                                ' damit Kriterien auch eingelesen werden, wenn noch keine Bewertung existiert ...
                                With cBewertung
                                    '.bewerterName = resultVerantwortlich
                                    .colorIndex = bewertungsAmpel
                                    .datum = Date.Now
                                    .description = explanation
                                    ' deliverables sind jetzt Bestandteil von clsMeilenstein (List (of String))  
                                    '.deliverables = deliverables
                                End With

                                isNotDuplikate = True
                                If awinSettings.eliminateDuplicates And hproj.hierarchy.containsKey(calcHryElemKey(mappedMSname, True)) Then
                                    ' nur dann kann es Duplikate geben 
                                    Dim duplicateSiblingID As String = hproj.getDuplicateMsSiblingID(mappedMSname, parentphase.nameID, _
                                                                                                         milestonedate, 0)

                                    If duplicateSiblingID = "" Then
                                        isNotDuplikate = True
                                    Else
                                        isNotDuplikate = False
                                        prtLine.planelement = aktTask_j.name
                                        prtLine.hgColor = awinSettings.AmpelRot
                                        prtLine.grund = "Meilenstein wurde eliminiert: Duplikat zur Geschwister-Phase"
                                        'Call logfileSchreiben("Fehler, RXFImport:" & mappedMSname & " ist Duplikat zu Geschwister " & elemNameOfElemID(duplicateSiblingID) & _
                                        '" und wird ignoriert ", hproj.name, anzFehler)
                                    End If

                                End If

                                If isNotDuplikate Then

                                    With cmilestone
                                        .setDate = milestonedate
                                        '.verantwortlich = resultVerantwortlich

                                        ' hier muss für gleiche PhasenNamen als Geschwister noch eine lfdNummer angehängt werden
                                        ' es muss überprüft werden, ob es Geschwister mit gleichem Namen gibt:
                                        ' wenn ja, wird an den mappedPhaseName eine LFdNr. ergänzt,bis der Name innerhalb der Geschwistergruppe eindeutig ist.

                                        If awinSettings.createUniqueSiblingNames Then
                                            mappedMSname = hproj.hierarchy.findUniqueGeschwisterName(parentelemID, mappedMSname, True)
                                        End If

                                        .nameID = hproj.hierarchy.findUniqueElemKey(mappedMSname, True)
                                        If Not cBewertung Is Nothing Then
                                            .addBewertung(cBewertung)
                                        End If
                                    End With

                                    With parentphase
                                        .addMilestone(cmilestone, origName:=origMSname)
                                    End With

                                    prtLine.hierarchie = hproj.hierarchy.getBreadCrumb(cmilestone.nameID)
                                    prtLine.PThierarchie = hproj.hierarchy.getBreadCrumb(cmilestone.nameID)
                                    prtLine.planelement = aktTask_j.name
                                    prtLine.abkürzung = MilestoneDefinitions.getAbbrev(cmilestone.name)
                                    prtLine.planeleÜbern = cmilestone.name

                                    prtLine.klasse = aktTask_j.taskType.Value
                                    prtLine.PTklasse = mapToAppearance(aktTask_j.taskType.Value, True)

                                    prtliste.Add(zeile, prtLine)
                                    zeile = zeile + 1

                                    Dim quelle As String = prtLine.quelle
                                    prtLine = New clsProtokoll(hproj.name, quelle) ' neue Protokollzeile
                                    prtLine.actDate = ""

                                Else
                                    prtliste.Add(zeile, prtLine) ' Protokollzeile in Liste eintragen
                                    zeile = zeile + 1

                                    Dim quelle As String = prtLine.quelle
                                    prtLine = New clsProtokoll(hproj.name, quelle) ' neue Protokollzeile
                                    prtLine.actDate = ""

                                End If

                            Else
                                prtLine.planelement = aktTask_j.name
                                prtLine.hgColor = awinSettings.AmpelRot
                                prtLine.grund = "Meilenstein wurde ignoriert: unbekannter Bezeichner"

                                prtliste.Add(zeile, prtLine) ' Protokollzeile in Liste eintragen
                                zeile = zeile + 1

                                Dim quelle As String = prtLine.quelle
                                prtLine = New clsProtokoll(hproj.name, quelle) ' neue Protokollzeile
                                prtLine.actDate = ""
                            End If

                        Else
                            prtLine.planelement = aktTask_j.name
                            prtLine.hgColor = awinSettings.AmpelRot
                            prtLine.grund = "Meilenstein wurde ignoriert gemäß Eintrag im Wörterbuch"

                            prtliste.Add(zeile, prtLine) ' Protokollzeile in Liste eintragen
                            zeile = zeile + 1

                            Dim quelle As String = prtLine.quelle
                            prtLine = New clsProtokoll(hproj.name, quelle) ' neue Protokollzeile
                            prtLine.actDate = ""

                        End If     ' Ende: Meilenstein soll ignoriert werden



                    End If      '  Ende: ist MEILENSTEIN

                Catch ex As Exception

                    Call logfileSchreiben(ex.Message, hproj.name, anzFehler)


                End Try


            End If

        Next j    ' Ende Schleife über alle Tasks
    End Sub

    ''' <summary>
    ''' nach BMW-Vorgaben:
    ''' bestimmt aus dem übergebenen VorlagenNamen ( =  der CustomValue "UsA_SERVICE_SPALTE_B" aus Phase "Projektphasen" ) 
    ''' den tatsächlichen VorlagenNamen des Projekts 
    '''     ''' </summary>
    ''' <param name="hproj"></param>aktuelles zu lesendes Projekt
    ''' <returns></returns>fertig zusammengesetzter VorlagenName des Projekts (gemäß BMW vorschriften
    ''' <remarks></remarks>
    Private Function findBMWVorlagenName(ByVal hproj As clsProjekt) As String


        Dim vorNam1 = "rel 4"
        Dim typkennung As String = hproj.VorlagenName      ' hier ist aber nur enthalten, eA, wA, E  usw.
        Dim anlaufkennung As String = "03"
        Dim firstMS As Integer = hproj.hierarchy.getIndexOf1stMilestone


        Dim hrchyhproj As clsHierarchy = hproj.hierarchy
        For phi = 1 To firstMS - 1
            Dim phID As String = hrchyhproj.getIDAtIndex(phi)
            Dim phName As String = elemNameOfElemID(phID)
            If phName.Contains("I-Stufen") Then
                Dim pharray() As String = Split(phName, " ", 5)
                vorNam1 = "rel 5"
            End If
        Next phi
        For msi = firstMS To hrchyhproj.count - 1
            Dim msID As String = hrchyhproj.getIDAtIndex(msi)
            Dim msName As String = elemNameOfElemID(msID)
            If msName.Contains("SOP") Then
                Dim msarray() As String = Split(msName, " ", 5)
                Try
                    Dim sopdate As Date = hproj.getMilestoneDate(msID)

                    If DateDiff(DateInterval.Month, StartofCalendar, sopdate) > 0 Then
                        Dim sopMonth As Integer = sopdate.Month
                        If sopMonth >= 3 And sopMonth <= 6 Then
                            anlaufkennung = "03"
                        ElseIf sopMonth >= 7 And sopMonth <= 10 Then
                            anlaufkennung = "07"
                        Else
                            anlaufkennung = "11"
                        End If
                    Else
                        anlaufkennung = sopdate.Month.ToString("D2")       ' Monat mindestens zweistellig angeben
                    End If

                Catch ex As Exception
                    anlaufkennung = "?"
                End Try

            End If
        Next
        Try
            If Not IsNothing(typkennung) Then

                If typkennung.Contains("SB") Then
                    typkennung = "SBWE"
                ElseIf typkennung.Contains("eA") Then
                    typkennung = "eA"
                ElseIf typkennung.Contains("wA") Then
                    typkennung = "wA"
                ElseIf typkennung.Contains("E") Then
                    typkennung = "E"
                Else
                    typkennung = "?"
                End If
            Else
                typkennung = "?"
            End If
        Catch ex As Exception
            typkennung = "?"
        End Try

        findBMWVorlagenName = vorNam1 & " " & typkennung & "-" & anlaufkennung

    End Function


    ''' <summary>
    ''' Behandelt den Fehler UnkonwnNode beim Einlesen eines XML-Files (oder RXF-Files)
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub deserializer_UnknownNode(sender As Object, e As XmlNodeEventArgs)
        Call MsgBox(("XMLImport: Unknown Node:" & e.Name & ControlChars.Tab & e.Text))
    End Sub 'serializer_UnknownNode


    ''' <summary>
    ''' Behandelt den Fehler UnkonwnAttribute beim Einlesen eines XML-Files (oder RXF-Files)
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub deserializer_UnknownAttribute(sender As Object, e As XmlAttributeEventArgs)
        Dim attr As System.Xml.XmlAttribute = e.Attr
        Call MsgBox(("XMLImport: Unknown attribute " & attr.Name & "='" & attr.Value & "'"))
    End Sub 'serializer_UnknownAttribute

    ''' <summary>
    ''' in der ganzen Datei sfilename wird der String searchstr durch replacestr ersetzt
    ''' </summary>
    ''' <param name="sfilename"></param>Name der Datei, in der die Ersetzung erfolgen soll
    ''' <param name="searchstr"></param>zu ersetzender String
    ''' <param name="replacestr"></param>neuer String
    ''' <returns></returns>Name der neuen Datei
    ''' <remarks></remarks>
    Private Function replaceStringInFile(ByVal sfilename As String, ByVal searchstr As String, ByVal replacestr As String) As String

        'Declare ALL of your variables :)
        Const ForReading = 1    '
        Dim fileToRead As String = sfilename  ' the path of the file to read
        Dim tstr() As String = Split(sfilename, ".", 2)
        Dim fileToWrite As String = tstr(0) & ".new"     ' the path of a new file
        Dim FSO As Object
        Dim readFile As Object  'the file you will READ
        Dim writeFile As Object 'the file you will CREATE
        Dim repLine As Object   'the array of lines you will WRITE
        Dim ln As Object
        Dim l As Long

        FSO = CreateObject("Scripting.FileSystemObject")
        readFile = FSO.OpenTextFile(fileToRead, ForReading, False)
        writeFile = FSO.CreateTextFile(fileToWrite, True, False)

        '# Read entire file into an array & close it
        repLine = Split(readFile.ReadAll, vbNewLine)
        readFile.Close()

        '# iterate the array and do the replacement line by line

        For Each ln In repLine
            ln = Replace(ln, searchstr, replacestr)
            repLine(l) = ln
            l = l + 1
        Next

        '# Write to the array items to the file
        writeFile.Write(Join(repLine, vbNewLine))
        writeFile.Close()

        '# clean up
        readFile = Nothing
        writeFile = Nothing
        FSO = Nothing
        replaceStringInFile = fileToWrite

    End Function



    ' '' ''' <summary>
    ' '' ''' initialisert das Logfile
    ' '' ''' </summary>
    ' '' ''' <remarks></remarks>
    ' ''Sub logfileInit()

    ' ''    Try

    ' ''        With CType(xlsLogfile.Worksheets(1), Excel.Worksheet)
    ' ''            .Name = "logBuch"
    ' ''            CType(.Cells(1, 1), Excel.Range).Value = "logfile erzeugt " & Date.Now.ToString
    ' ''            CType(.Columns(1), Excel.Range).ColumnWidth = 100
    ' ''            CType(.Columns(2), Excel.Range).ColumnWidth = 50
    ' ''            CType(.Columns(3), Excel.Range).ColumnWidth = 20
    ' ''        End With
    ' ''    Catch ex As Exception

    ' ''    End Try


    ' ''End Sub
    ' '' ''' <summary>
    ' '' ''' schreibt in das logfile 
    ' '' ''' </summary>
    ' '' ''' <param name="text"></param>
    ' '' ''' <param name="addOn"></param>
    ' '' ''' <remarks></remarks>
    ' ''Sub logfileSchreiben(ByVal text As String, ByVal addOn As String, ByRef anzFehler As Long)

    ' ''    Dim obj As Object

    ' ''    Try
    ' ''        obj = CType(CType(xlsLogfile.Worksheets("logBuch"), Excel.Worksheet).Rows(1), Excel.Range).Insert(Excel.XlInsertShiftDirection.xlShiftDown)

    ' ''        With CType(xlsLogfile.Worksheets("logBuch"), Excel.Worksheet)
    ' ''            CType(.Cells(1, 1), Excel.Range).Value = text
    ' ''            CType(.Cells(1, 2), Excel.Range).Value = addOn
    ' ''            CType(.Cells(1, 3), Excel.Range).Value = Date.Now
    ' ''        End With
    ' ''        anzFehler = anzFehler + 1


    ' ''    Catch ex As Exception

    ' ''    End Try

    ' ''End Sub
    ' '' ''' <summary>
    ' '' ''' öffnet das LogFile
    ' '' ''' </summary>
    ' '' ''' <remarks></remarks>
    ' ''Sub logfileOpen()

    ' ''    appInstance.ScreenUpdating = False

    ' ''    ' aktives Workbook merken im Variable actualWB
    ' ''    Dim actualWB As String = appInstance.ActiveWorkbook.Name

    ' ''    If My.Computer.FileSystem.FileExists(awinPath & logFileName) Then
    ' ''        Try
    ' ''            xlsLogfile = appInstance.Workbooks.Open(awinPath & logFileName)
    ' ''            myLogfile = appInstance.ActiveWorkbook.Name
    ' ''        Catch ex As Exception

    ' ''            logmessage = "Öffnen von " & logFileName & " fehlgeschlagen" & vbLf & _
    ' ''                                            "falls die Datei bereits geöffnet ist: Schließen Sie sie bitte"
    ' ''            'Call logfileSchreiben(logMessage, " ")
    ' ''            Throw New ArgumentException(logmessage)

    ' ''        End Try

    ' ''    Else
    ' ''        ' Logfile neu anlegen 
    ' ''        xlsLogfile = appInstance.Workbooks.Add
    ' ''        Call logfileInit()
    ' ''        xlsLogfile.SaveAs(awinPath & logFileName)
    ' ''        myLogfile = xlsLogfile.Name

    ' ''    End If

    ' ''    ' Workbook, das vor dem öffnen des Logfiles aktiv war, wieder aktivieren
    ' ''    appInstance.Workbooks(actualWB).Activate()

    ' ''End Sub



    ' '' ''' <summary>
    ' '' ''' schliesst  das logfile 
    ' '' ''' </summary>  
    ' '' ''' <remarks></remarks>
    ' ''Sub logfileSchliessen()

    ' ''    appInstance.EnableEvents = False
    ' ''    Try

    ' ''        appInstance.Workbooks(myLogfile).Close(SaveChanges:=True)

    ' ''    Catch ex As Exception
    ' ''        Call MsgBox("Fehler beim Schließen des Logfiles")
    ' ''    End Try
    ' ''    appInstance.EnableEvents = True
    ' ''End Sub

    ''' <summary>
    ''' initialisert im Inputfile die Tabelle 'Logbuch'
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Sub InitProtokoll(ByRef wslogbuch As Excel.Worksheet, ByVal tabblattname As String)

        ' diese Variable sagt, ob das Tabellenblatt existiert hat; wenn nein, müssen die Spalten-Breiten gesetzt werden 
        Dim didntExist As Boolean

        Try
            wslogbuch = CType(xlsLogfile.Worksheets(tabblattname), _
               Global.Microsoft.Office.Interop.Excel.Worksheet)


            If Not IsNothing(wslogbuch) Then

                ' Änderung tk: 16.1.16
                ' es reicht die Inhalte zu löschen ...  
                wslogbuch.Cells.Clear()
                didntExist = False
                'xlsLogfile.Worksheets.Application.DisplayAlerts = False
                'wslogbuch.Delete()
                'xlsLogfile.Worksheets.Application.DisplayAlerts = True

                'wslogbuch = CType(xlsLogfile.Worksheets.Add(), _
                '   Global.Microsoft.Office.Interop.Excel.Worksheet)
                'wslogbuch.Name = tabblattname
            End If
        Catch ex As Exception
            'wsLogbuch = CType(xlsInput.Worksheets.Add(After:=xlsInput.Worksheets.Count), _
            '   Global.Microsoft.Office.Interop.Excel.Worksheet)
            wslogbuch = CType(xlsLogfile.Worksheets.Add(), _
                Global.Microsoft.Office.Interop.Excel.Worksheet)
            wslogbuch.Name = tabblattname
            didntExist = True
        End Try


        With wslogbuch

            If didntExist Then
                .Rows.RowHeight = 15
                CType(.Rows(1), Excel.Range).RowHeight = 30
                CType(.Rows(1), Excel.Range).Font.Bold = True
            End If


            If awinSettings.fullProtocol Then
                CType(.Cells(1, 1), Excel.Range).Value() = "Datum"
                CType(.Cells(1, 2), Excel.Range).Value() = "Projekt"
                CType(.Cells(1, 3), Excel.Range).Value() = "Hierarchie"
                CType(.Cells(1, 4), Excel.Range).Value() = "Plan-Element"
                CType(.Cells(1, 5), Excel.Range).Value() = "Klasse"
                CType(.Cells(1, 6), Excel.Range).Value() = "Abkürzung"
                CType(.Cells(1, 7), Excel.Range).Value() = "Quelle"
                CType(.Cells(1, 8), Excel.Range).Value() = "Übernommen als"
                CType(.Cells(1, 9), Excel.Range).Value() = "Grund"
                CType(.Cells(1, 10), Excel.Range).Value() = "PT Hierarchie"
                CType(.Cells(1, 11), Excel.Range).Value() = "PT Klasse"

                ' nur verändern, wenn es nicht vorher schon existiert hat ... 
                ' falls der Anwender sich die Breiten so hingerichtet hat , wie er es gerne hätte, 
                ' sollte das nicht verändert werden 
                If didntExist Then
                    CType(.Columns(1), Excel.Range).ColumnWidth = 10
                    CType(.Columns(2), Excel.Range).ColumnWidth = 40
                    CType(.Columns(3), Excel.Range).ColumnWidth = 40
                    CType(.Columns(4), Excel.Range).ColumnWidth = 40
                    CType(.Columns(5), Excel.Range).ColumnWidth = 40
                    CType(.Columns(6), Excel.Range).ColumnWidth = 40
                    CType(.Columns(7), Excel.Range).ColumnWidth = 40
                    CType(.Columns(8), Excel.Range).ColumnWidth = 40
                    CType(.Columns(9), Excel.Range).ColumnWidth = 40
                    CType(.Columns(10), Excel.Range).ColumnWidth = 40
                    CType(.Columns(11), Excel.Range).ColumnWidth = 40
                End If

            Else
                CType(.Cells(1, 1), Excel.Range).Value() = "Datum"
                CType(.Cells(1, 2), Excel.Range).Value() = "Projekt"
                CType(.Cells(1, 4), Excel.Range).Value() = "Plan-Element"
                CType(.Cells(1, 8), Excel.Range).Value() = "Übernommen als"
                CType(.Cells(1, 9), Excel.Range).Value() = "Grund"

                ' nur verändern, wenn es nicht vorher schon existiert hat ... 
                ' falls der Anwender sich die Breiten so hingerichtet hat , wie er es gerne hätte, 
                ' sollte das nicht verändert werden 
                If didntExist Then
                    CType(.Columns(1), Excel.Range).ColumnWidth = 18
                    CType(.Columns(2), Excel.Range).ColumnWidth = 35
                    CType(.Columns(3), Excel.Range).ColumnWidth = 5
                    CType(.Columns(4), Excel.Range).ColumnWidth = 40
                    CType(.Columns(5), Excel.Range).ColumnWidth = 10
                    CType(.Columns(6), Excel.Range).ColumnWidth = 10
                    CType(.Columns(7), Excel.Range).ColumnWidth = 10
                    CType(.Columns(8), Excel.Range).ColumnWidth = 40
                    CType(.Columns(9), Excel.Range).ColumnWidth = 40
                End If

            End If

        End With

    End Sub
    ''' <summary>
    ''' schreibt das Protokoll in das Tabellenblatt
    ''' es wird eine Range definiert, die soviele Zeilen enthält wie öt 
    ''' </summary>
    ''' <param name="prtliste"></param>
    ''' <param name="tabblattname"></param>
    ''' <remarks></remarks>
    Sub writeProtokoll(ByRef prtliste As SortedList(Of Integer, clsProtokoll), ByVal tabblattname As String)

        Dim zelle As Excel.Range = Nothing
        Dim zeile As Integer

        Dim anzZeilen As Integer = prtliste.Count

        Dim wsLogbuch As Excel.Worksheet = Nothing

        Try
            Call InitProtokoll(wsLogbuch, tabblattname) ' Tabelle Logbuch wird initialisiert
            If Not IsNothing(xlsLogfile) Then
                xlsLogfile.Save()
            End If


        Catch ex As Exception

            Call MsgBox("Fehler beim Initialisieren des Protokolls")
        End Try

        Dim protokollRange As Excel.Range = wsLogbuch.Cells


        For Each prtline As KeyValuePair(Of Integer, clsProtokoll) In prtliste
            Try
                'rowOffset = CType(CType(xlsLogfile.Worksheets(Me.tabblattname), Excel.Worksheet).Cells(20000, 1), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlUp).Row
                zeile = prtline.Key

                With wsLogbuch

                    ' Änderung tk: das dauert sehr lange ... 
                    'If awinSettings.fullProtocol Then
                    '    CType(.Cells(zeile, 1), Excel.Range).Value() = prtline.Value.actDate
                    '    CType(.Cells(zeile, 2), Excel.Range).Value() = prtline.Value.Projekt
                    '    CType(.Cells(zeile, 3), Excel.Range).Value() = prtline.Value.hierarchie
                    '    CType(.Cells(zeile, 4), Excel.Range).Value() = prtline.Value.planelement
                    '    CType(.Cells(zeile, 5), Excel.Range).Value() = prtline.Value.klasse
                    '    CType(.Cells(zeile, 6), Excel.Range).Value() = prtline.Value.abkürzung
                    '    CType(.Cells(zeile, 7), Excel.Range).Value() = prtline.Value.quelle
                    '    CType(.Cells(zeile, 8), Excel.Range).Value() = prtline.Value.planeleÜbern
                    '    CType(.Cells(zeile, 8), Excel.Range).Interior.Color = prtline.Value.hgColor
                    '    CType(.Cells(zeile, 9), Excel.Range).Value() = prtline.Value.grund
                    '    CType(.Cells(zeile, 10), Excel.Range).Value() = prtline.Value.PThierarchie
                    '    CType(.Cells(zeile, 11), Excel.Range).Value() = prtline.Value.PTklasse
                    'Else
                    '    CType(.Cells(zeile, 1), Excel.Range).Value() = prtline.Value.actDate
                    '    CType(.Cells(zeile, 2), Excel.Range).Value() = prtline.Value.Projekt
                    '    CType(.Cells(zeile, 4), Excel.Range).Value() = prtline.Value.planelement
                    '    CType(.Cells(zeile, 8), Excel.Range).Value() = prtline.Value.planeleÜbern
                    '    CType(.Cells(zeile, 8), Excel.Range).Interior.Color = prtline.Value.hgColor
                    '    CType(.Cells(zeile, 9), Excel.Range).Value() = prtline.Value.grund
                    'End If

                    If awinSettings.fullProtocol Then
                        protokollRange.Cells(zeile, 1).Value = prtline.Value.actDate
                        protokollRange.Cells(zeile, 2).Value = prtline.Value.Projekt
                        protokollRange.Cells(zeile, 3).Value = prtline.Value.hierarchie
                        protokollRange.Cells(zeile, 4).Value = prtline.Value.planelement
                        protokollRange.Cells(zeile, 5).Value = prtline.Value.klasse
                        protokollRange.Cells(zeile, 6).Value = prtline.Value.abkürzung
                        protokollRange.Cells(zeile, 7).Value = prtline.Value.quelle
                        protokollRange.Cells(zeile, 8).Value = prtline.Value.planeleÜbern
                        protokollRange.Cells(zeile, 8).Interior.Color = prtline.Value.hgColor
                        protokollRange.Cells(zeile, 9).Value = prtline.Value.grund
                        protokollRange.Cells(zeile, 10).Value = prtline.Value.PThierarchie
                        protokollRange.Cells(zeile, 11).Value = prtline.Value.PTklasse
                    Else
                        protokollRange.Cells(zeile, 1).Value = prtline.Value.actDate
                        protokollRange.Cells(zeile, 2).Value = prtline.Value.Projekt
                        protokollRange.Cells(zeile, 4).Value = prtline.Value.planelement
                        protokollRange.Cells(zeile, 8).Value = prtline.Value.planeleÜbern
                        protokollRange.Cells(zeile, 8).Interior.Color = prtline.Value.hgColor
                        protokollRange.Cells(zeile, 9).Value = prtline.Value.grund
                    End If

                End With
            Catch ex As Exception

            End Try

        Next

        ' Logbuch sichern
        If Not IsNothing(xlsLogfile) Then
            xlsLogfile.Save()
        End If

    End Sub



    Public Sub XMLExportReportProfil(ByVal profil As clsReport)

        Dim dirname As String = awinPath & ReportProfileOrdner
        Dim xmlfilename As String = dirname & "\" & profil.name & ".xml"

        Try

            If Not My.Computer.FileSystem.DirectoryExists(dirname) Then
                Try
                    My.Computer.FileSystem.CreateDirectory(dirname)
                Catch ex As Exception

                End Try
            End If

            Dim serializer = New DataContractSerializer(GetType(clsReport))

            ' ''Dim file As New FileStream(xmlfilename, FileMode.Create)
            ' ''serializer.WriteObject(file, profil)
            ' ''file.Close()

            Dim settings As New XmlWriterSettings()
            settings.Indent = True
            settings.IndentChars = (ControlChars.Tab)
            settings.OmitXmlDeclaration = True

            Dim writer As XmlWriter = XmlWriter.Create(xmlfilename, settings)
            serializer.WriteObject(writer, profil)
            writer.Flush()
            writer.Close()

        Catch ex As Exception

            Call MsgBox("Beim Schreiben der XML-Datei '" & xmlfilename & "' ist ein Fehler aufgetreten !")

        End Try

    End Sub


    Public Sub XMLExportReportProfil(ByVal profil As clsReportAll)

        Dim dirname As String = awinPath & ReportProfileOrdner
        Dim xmlfilename As String = dirname & "\" & profil.name & ".xml"

        Try

            If Not My.Computer.FileSystem.DirectoryExists(dirname) Then
                Try
                    My.Computer.FileSystem.CreateDirectory(dirname)
                Catch ex As Exception

                End Try
            End If

            Dim serializer = New DataContractSerializer(GetType(clsReportAll))

            ' ''Dim file As New FileStream(xmlfilename, FileMode.Create)
            ' ''serializer.WriteObject(file, profil)
            ' ''file.Close()

            Dim settings As New XmlWriterSettings()
            settings.Indent = True
            settings.IndentChars = (ControlChars.Tab)
            settings.OmitXmlDeclaration = True

            Dim writer As XmlWriter = XmlWriter.Create(xmlfilename, settings)
            serializer.WriteObject(writer, profil)
            writer.Flush()
            writer.Close()

        Catch ex As Exception

            Call MsgBox("Beim Schreiben der XML-Datei '" & xmlfilename & "' ist ein Fehler aufgetreten !")

        End Try

    End Sub


    Public Function XMLImportReportProfil(ByVal profilName As String) As clsReportAll

        Dim ergprofil As New clsReportAll
        Dim aktfile As FileStream = Nothing

        Dim xmlfilename As String = awinPath & ReportProfileOrdner & "\" & profilName & ".xml"
        Try
            ' ur: 31.03.2017 von nun an wird auch bei BHTC mit der Struktur clsReportAll agiert.
            '                alte ReportProfile von BHTC können trotzdem noch gelesen werden, siehe Catch-fall
            Dim serializer = New DataContractSerializer(GetType(clsReportAll))
            Dim profil As New clsReportAll

            ' XML-Datei Öffnen
            ' A FileStream is needed to read the XML document.
            Dim file As New FileStream(xmlfilename, FileMode.Open)
            aktfile = file
            profil = serializer.ReadObject(file)
            file.Close()
            ergprofil = profil

            XMLImportReportProfil = ergprofil


        Catch ex As Exception

            ' ur: 18.05.2017: oben geöffnete Datei zunächst schließen, da falsche Format
            If Not IsNothing(aktfile) Then
                aktfile.Close()
            End If


            ' ur: 31.03.2017 neu eingefügt
            Try
                Dim serializer = New DataContractSerializer(GetType(clsReport))
                Dim profil As New clsReport

                ' XML-Datei Öffnen
                ' A FileStream is needed to read the XML document.
                Dim file As New FileStream(xmlfilename, FileMode.Open)
                profil = serializer.ReadObject(file)
                file.Close()
                profil.CopyTo(ergprofil)

                XMLImportReportProfil = ergprofil


            Catch ex2 As Exception

                Call MsgBox("Beim Lesen der XML-Datei '" & xmlfilename & "' ist ein Fehler aufgetreten !")
                XMLImportReportProfil = Nothing
            End Try

        End Try

    End Function

    Public Function XMLImportReportAllProfil(ByVal profilName As String) As clsReportAll

        Dim profil As New clsReportAll

        Dim serializer = New DataContractSerializer(GetType(clsReportAll))
        Dim xmlfilename As String = awinPath & ReportProfileOrdner & "\" & profilName & ".xml"
        Try

            ' XML-Datei Öffnen
            ' A FileStream is needed to read the XML document.
            Dim file As New FileStream(xmlfilename, FileMode.Open)
            profil = serializer.ReadObject(file)
            file.Close()

            XMLImportReportAllProfil = profil

        Catch ex As Exception

            Call MsgBox("Beim Lesen der XML-Datei '" & xmlfilename & "' ist ein Fehler aufgetreten !")
            XMLImportReportAllProfil = Nothing
        End Try

    End Function

    Public Sub retrieveProfilSelection(ByVal profilName As String, ByVal menuOption As Integer, _
                                     ByRef selectedBUs As Collection, ByRef selectedTyps As Collection, _
                                     ByRef selectedPhases As Collection, ByRef selectedMilestones As Collection, _
                                     ByRef selectedRoles As Collection, ByRef selectedCosts As Collection, ByRef reportProfil As clsReportAll)
        Try
            If menuOption = PTmenue.reportBHTC Then

                ' Datumsangaben sichern
                Dim vondate_sav As Date = reportProfil.VonDate
                Dim bisdate_sav As Date = reportProfil.BisDate
                Dim PPTvondate_sav As Date = reportProfil.CalendarVonDate
                Dim PPTbisdate_sav As Date = reportProfil.CalendarBisDate

                ' Projekte sichern
                Dim projects_sav As New SortedList(Of Double, String)
                For Each kvp As KeyValuePair(Of Double, String) In reportProfil.Projects
                    projects_sav.Add(kvp.Key, kvp.Value)
                Next

                ' Datumsangaben zurücksichern
                reportProfil.CalendarVonDate = PPTvondate_sav
                reportProfil.CalendarBisDate = PPTbisdate_sav
                reportProfil.calcRepVonBis(vondate_sav, bisdate_sav)



                ' für BHTC immer true
                reportProfil.ExtendedMode = True
                ' für BHTC immer false
                reportProfil.Ampeln = False
                reportProfil.AllIfOne = False
                reportProfil.FullyContained = False
                reportProfil.SortedDauer = False
                reportProfil.ProjectLine = False
                reportProfil.UseOriginalNames = False

                ' Projekte zurücksichern
                reportProfil.Projects.Clear()
                For Each kvp As KeyValuePair(Of Double, String) In projects_sav
                    reportProfil.Projects.Add(kvp.Key, kvp.Value)
                Next

            Else
                '  menuOption = PTmenue.reportMultiprojektTafel


                ' Einlesen des ausgewählten ReportProfils
                reportProfil = XMLImportReportProfil(profilName)


            End If


            '  und bereitstellen der Auswahl für Hierarchieselection
            selectedPhases = copySortedListtoColl(reportProfil.Phases)
            selectedMilestones = copySortedListtoColl(reportProfil.Milestones)
            selectedRoles = copySortedListtoColl(reportProfil.Roles)
            selectedCosts = copySortedListtoColl(reportProfil.Costs)
            selectedBUs = copySortedListtoColl(reportProfil.BUs)
            selectedTyps = copySortedListtoColl(reportProfil.Typs)

        Catch ex As Exception
            Throw New ArgumentException("Fehler beim Lesen des ReportProfils: retrieveProfilSelection")
        End Try

    End Sub


    Public Sub storeReportProfil(ByVal menuOption As Integer, _
                                     ByVal selectedBUs As Collection, ByVal selectedTyps As Collection, _
                                     ByVal selectedPhases As Collection, ByVal selectedMilestones As Collection, _
                                     ByVal selectedRoles As Collection, ByVal selectedCosts As Collection, ByVal reportProfil As clsReportAll)



        '  und bereitstellen der Auswahl für Hierarchieselection
        reportProfil.Phases = copyColltoSortedList(selectedPhases)
        reportProfil.Milestones = copyColltoSortedList(selectedMilestones)
        reportProfil.Roles = copyColltoSortedList(selectedRoles)
        reportProfil.Costs = copyColltoSortedList(selectedCosts)
        reportProfil.BUs = copyColltoSortedList(selectedBUs)
        reportProfil.Typs = copyColltoSortedList(selectedTyps)


        With awinSettings

            ' tk : wird für Darstellung Projekt auf Multiprojekt Tafel verwendet; hier nicht setzen ! 
            '.drawProjectLine = True
            reportProfil.ExtendedMode = .mppExtendedMode
            reportProfil.OnePage = .mppOnePage
            reportProfil.AllIfOne = .mppShowAllIfOne
            reportProfil.Ampeln = .mppShowAmpel
            reportProfil.Legend = .mppShowLegend
            reportProfil.MSDate = .mppShowMsDate
            reportProfil.MSName = .mppShowMsName
            reportProfil.PhDate = .mppShowPhDate
            reportProfil.PhName = .mppShowPhName
            reportProfil.ProjectLine = .mppShowProjectLine
            reportProfil.SortedDauer = .mppSortiertDauer
            reportProfil.VLinien = .mppVertikalesRaster
            reportProfil.FullyContained = .mppFullyContained
            reportProfil.ShowHorizontals = .mppShowHorizontals
            reportProfil.UseAbbreviation = .mppUseAbbreviation
            reportProfil.UseOriginalNames = .mppUseOriginalNames
            reportProfil.KwInMilestone = .mppKwInMilestone

            If menuOption = PTmenue.reportMultiprojektTafel Then
                reportProfil.projectsWithNoMPmayPass = .mppProjectsWithNoMPmayPass
            Else
                ' dann gilt: menuOption = PTmenue.reportBHTC
                reportProfil.projectsWithNoMPmayPass = Nothing
                reportProfil.description = ""
            End If

        End With


        ' Schreiben des ausgewählten ReportProfils
        Call XMLExportReportProfil(reportProfil)

    End Sub
    ''' <summary>
    ''' synchronisiert die globalen mit den lokalen Konfigurations-Dateien 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub synchronizeGlobalToLocalFolder()


        Dim srcFile As String
        Dim destFile As String
        Dim destdir As String

        Try


            'Prüfen, ob der Globale Folder existiert
            If Not My.Computer.FileSystem.DirectoryExists(globalPath) Then

                Throw New ArgumentException("Globaler Requirementsordner " & globalPath & " existiert nicht!")

            Else
                ' '' ''Prüfen, ob der Lokale Folder existiert
                '' ''If Not My.Computer.FileSystem.DirectoryExists(awinPath) Then

                '' ''    Call MsgBox("lokaler Requirementsordner " & awinPath & " existiert nicht!")

                ' Lokaler Requirementsordner wird erzeugt, mit allen Unterdirectories
                Try


                    My.Computer.FileSystem.CreateDirectory(awinPath)
                    My.Computer.FileSystem.CreateDirectory(awinPath & requirementsOrdner)

                    For Each gdir In My.Computer.FileSystem.GetDirectories(globalPath & requirementsOrdner)

                        ' Name des lokalen Directories zusammensetzen
                        Dim hstr() As String
                        hstr = gdir.Split(New Char() {CChar("\")})

                        ' Name des destinationDirectories zusammen setzen
                        destdir = awinPath & requirementsOrdner

                        destdir = destdir & hstr(hstr.Length - 1)

                        My.Computer.FileSystem.CreateDirectory(destdir)

                    Next

                    ' ''My.Computer.FileSystem.CreateDirectory(awinPath & projektVorlagenOrdner)
                    ' ''My.Computer.FileSystem.CreateDirectory(awinPath & modulVorlagenOrdner)
                    ' ''My.Computer.FileSystem.CreateDirectory(awinPath & projektRessOrdner)
                    ' ''My.Computer.FileSystem.CreateDirectory(awinPath & RepProjectVorOrdner)
                    ' ''My.Computer.FileSystem.CreateDirectory(awinPath & RepPortfolioVorOrdner)
                    ' ''My.Computer.FileSystem.CreateDirectory(awinPath & ReportProfileOrdner)


                    '' ''importOrdnerNames(PTImpExp.visbo) = awinPath & "Import\VISBO Steckbriefe"
                    '' ''importOrdnerNames(PTImpExp.rplan) = awinPath & "Import\RPLAN-Excel"
                    '' ''importOrdnerNames(PTImpExp.msproject) = awinPath & "Import\MSProject"
                    '' ''importOrdnerNames(PTImpExp.simpleScen) = awinPath & "Import\einfache Szenarien"
                    '' ''importOrdnerNames(PTImpExp.modulScen) = awinPath & "Import\modulare Szenarien"
                    '' ''importOrdnerNames(PTImpExp.addElements) = awinPath & "Import\addOn Regeln"
                    '' ''importOrdnerNames(PTImpExp.rplanrxf) = awinPath & "Import\RXF Files"

                    '' ''exportOrdnerNames(PTImpExp.visbo) = awinPath & "Export\VISBO Steckbriefe"
                    '' ''exportOrdnerNames(PTImpExp.rplan) = awinPath & "Export\RPLAN-Excel"
                    '' ''exportOrdnerNames(PTImpExp.msproject) = awinPath & "Export\MSProject"
                    '' ''exportOrdnerNames(PTImpExp.simpleScen) = awinPath & "Export\einfache Szenarien"
                    '' ''exportOrdnerNames(PTImpExp.modulScen) = awinPath & "Export\modulare Szenarien"

                    '' ''My.Computer.FileSystem.CreateDirectory(importOrdnerNames(PTImpExp.visbo))
                    '' ''My.Computer.FileSystem.CreateDirectory(importOrdnerNames(PTImpExp.rplan))
                    '' ''My.Computer.FileSystem.CreateDirectory(importOrdnerNames(PTImpExp.msproject))
                    '' ''My.Computer.FileSystem.CreateDirectory(importOrdnerNames(PTImpExp.simpleScen))
                    '' ''My.Computer.FileSystem.CreateDirectory(importOrdnerNames(PTImpExp.modulScen))
                    '' ''My.Computer.FileSystem.CreateDirectory(importOrdnerNames(PTImpExp.addElements))
                    '' ''My.Computer.FileSystem.CreateDirectory(importOrdnerNames(PTImpExp.rplanrxf))

                    '' ''My.Computer.FileSystem.CreateDirectory(exportOrdnerNames(PTImpExp.visbo))
                    '' ''My.Computer.FileSystem.CreateDirectory(exportOrdnerNames(PTImpExp.rplan))
                    '' ''My.Computer.FileSystem.CreateDirectory(exportOrdnerNames(PTImpExp.msproject))
                    '' ''My.Computer.FileSystem.CreateDirectory(exportOrdnerNames(PTImpExp.simpleScen))
                    '' ''My.Computer.FileSystem.CreateDirectory(exportOrdnerNames(PTImpExp.modulScen))

                Catch ex As Exception

                End Try
                '' ''Else


                Dim dirItem As String = globalPath & requirementsOrdner

                ' lokaler RequirementsOrdner existiert

                ' RequirementsOrdner:   alle Dateien , sofern sie im globalPath neuer als im awinPath sind kopieren

                For Each srcFile In My.Computer.FileSystem.GetFiles(dirItem)

                    ' Name des lokalen Files zusammensetzen
                    Dim hstr() As String
                    hstr = srcFile.Split(New Char() {CChar("\")})

                    ' Name des destinationDirectories zusammen setzen
                    destdir = awinPath & requirementsOrdner

                    destFile = destdir & "\" & hstr(hstr.Length - 1)

                    ' Test ob globales File neuer als lokales
                    Dim srcDate As Date = My.Computer.FileSystem.GetFileInfo(srcFile).LastWriteTime
                    Dim destDate As Date = My.Computer.FileSystem.GetFileInfo(destFile).LastWriteTime
                    Dim ddiff As Long = DateDiff(DateInterval.Second, _
                                                 My.Computer.FileSystem.GetFileInfo(srcFile).LastWriteTime, _
                                                 My.Computer.FileSystem.GetFileInfo(destFile).LastWriteTime)

                    ' Wenn globales neuer als lokales, dann von globalPath nach awinPath kopieren
                    If ddiff < 0 Then
                        ' Kopieren der Datei, mit Overwrite erzwingen
                        My.Computer.FileSystem.CopyFile(srcFile, destFile, True)
                        ' Debug Mode? 
                        If awinSettings.visboDebug Then
                            Call MsgBox("kopiert von global nach local:" & hstr(hstr.Length - 1))
                        End If
                    End If

                Next srcFile


                ' Unterdirectories von requirementsOrdner:      alle Dateien dieser  werden von globalPath nACH awinPath kopiert, sofern neueres Änderungsdatum

                For Each dirItem In My.Computer.FileSystem.GetDirectories(globalPath & requirementsOrdner)

                    For Each srcFile In My.Computer.FileSystem.GetFiles(dirItem)

                        ' Name des lokalen Files zusammensetzen
                        Dim hstr() As String
                        hstr = srcFile.Split(New Char() {CChar("\")})
                        ' Name des destinationDirectories zusammen setzen
                        Dim dirstr() As String
                        dirstr = dirItem.Split(New Char() {CChar("\")})
                        destdir = awinPath & requirementsOrdner & dirstr(dirstr.Length - 1)

                        destFile = destdir & "\" & hstr(hstr.Length - 1)

                        ' Test ob globales File neuer als lokales
                        Dim srcDate As Date = My.Computer.FileSystem.GetFileInfo(srcFile).LastWriteTime
                        Dim destDate As Date = My.Computer.FileSystem.GetFileInfo(destFile).LastWriteTime
                        Dim ddiff As Long = DateDiff(DateInterval.Second, _
                                                     My.Computer.FileSystem.GetFileInfo(srcFile).LastWriteTime, _
                                                     My.Computer.FileSystem.GetFileInfo(destFile).LastWriteTime)

                        ' Wenn globales neuer als lokales, dann von globalPath nach awinPath kopieren
                        If ddiff < 0 Then
                            ' Kopieren der Datei, mit Overwrite erzwingen
                            My.Computer.FileSystem.CopyFile(srcFile, destFile, True)

                            ' Debug Mode? 
                            If awinSettings.visboDebug Then
                                Call MsgBox("kopiert von global nach local:" & hstr(hstr.Length - 2) & "/" & hstr(hstr.Length - 1))
                            End If
                        End If

                    Next srcFile

                Next dirItem

            End If


            ' ''End If


        Catch ex As Exception

        End Try
    End Sub

    Public Sub XMLExportLicences(ByVal lic As clsLicences, ByVal nameLicfile As String)


        Dim xmlfilename As String = awinPath & nameLicfile

        Try

            Dim serializer = New DataContractSerializer(GetType(clsLicences))

            ' ''Dim file As New FileStream(xmlfilename, FileMode.Create)
            ' ''serializer.WriteObject(file, lic)
            ' ''file.Close()

            Dim settings As New XmlWriterSettings()
            settings.Indent = True
            settings.IndentChars = (ControlChars.Tab)
            settings.OmitXmlDeclaration = True

            Dim writer As XmlWriter = XmlWriter.Create(xmlfilename, settings)
            serializer.WriteObject(writer, lic)
            writer.Flush()
            writer.Close()
        Catch ex As Exception

            Call MsgBox("Beim Schreiben der XML-Datei '" & xmlfilename & "' ist ein Fehler aufgetreten !")

        End Try

    End Sub

    Public Function XMLImportLicences(ByVal licfile As String) As clsLicences

        Dim lic As New clsLicences

        Dim serializer = New DataContractSerializer(GetType(clsLicences))
        Dim xmlfilename As String = awinPath & licfile
        Try

            ' XML-Datei Öffnen
            ' A FileStream is needed to read the XML document.
            Dim file As New FileStream(xmlfilename, FileMode.Open)
            lic = serializer.ReadObject(file)
            file.Close()

            XMLImportLicences = lic

        Catch ex As Exception
            'Call MsgBox("Beim Lesen der XML-Datei '" & xmlfilename & "' ist ein Fehler aufgetreten !")
            Throw New ArgumentException("Beim Lesen der XML-Datei '" & xmlfilename & "' ist ein Fehler aufgetreten !")
            XMLImportLicences = Nothing
        End Try

    End Function

    Public Function XMLImportReportMsg(ByVal repMsgfile As String, ByVal language As String) As clsReportMessages

        Dim reportMessages As New clsReportMessages

        Dim serializer = New DataContractSerializer(GetType(clsReportMessages))
        Dim xmlfilename As String = awinPath & requirementsOrdner & repMsgfile & "_" & language & ".xml"
        Try

            ' XML-Datei Öffnen
            ' A FileStream is needed to read the XML document.
            Dim file As New FileStream(xmlfilename, FileMode.Open)
            reportMessages = serializer.ReadObject(file)
            file.Close()

            XMLImportReportMsg = reportMessages

        Catch ex As Exception

            Call MsgBox("Beim Lesen der XML-Datei '" & xmlfilename & "' ist ein Fehler aufgetreten !")
            XMLImportReportMsg = Nothing
        End Try

    End Function

    Public Sub XMLExportReportMsg(ByVal reportMsg As clsReportMessages, ByVal repMsgfile As String, ByVal language As String)



        Dim xmlfilename As String = awinPath & requirementsOrdner & repMsgfile & "_" & language & ".xml"
        Try
            Dim serializer = New DataContractSerializer(GetType(clsReportMessages))

            ' ''Dim file As New FileStream(xmlfilename, FileMode.Create)
            ' ''serializer.WriteObject(file, lic)
            ' ''file.Close()

            Dim settings As New XmlWriterSettings()
            settings.Indent = True
            settings.IndentChars = (ControlChars.Tab)
            settings.OmitXmlDeclaration = True

            Dim writer As XmlWriter = XmlWriter.Create(xmlfilename, settings)
            serializer.WriteObject(writer, reportMsg)
            writer.Flush()
            writer.Close()

        Catch ex As Exception

            Call MsgBox("Beim Schreiben der XML-Datei '" & xmlfilename & "' ist ein Fehler aufgetreten !")

        End Try
    End Sub
    ''' <summary>
    ''' importiert die ProjectboardConfig.xml
    ''' </summary>
    ''' <param name="cfgXMLfilename"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function XMLImportConfig(ByVal cfgXMLfilename As String) As configuration

        ' XML-Datei Öffnen
        ' A FileStream is needed to read the XML document.
        Dim fs As New FileStream(cfgXMLfilename, FileMode.Open)

        ' Declare an object variable of the type to be deserialized.
        Dim cfgs As New configuration           ' Class configuration erzeugt aus Projectboard.dll.config
        Try


            ' Create an instance of the XmlSerializer class;
            ' specify the type of object to be deserialized.
            Dim deserializer As New XmlSerializer(GetType(configuration))


            ' If the XML document has been altered with unknown
            ' nodes or attributes, handle them with the
            ' UnknownNode and UnknownAttribute events.
            AddHandler deserializer.UnknownNode, AddressOf deserializer_UnknownNode
            AddHandler deserializer.UnknownAttribute, AddressOf deserializer_UnknownAttribute


            ' Einlesen des kompletten XML-Dokument im die Klasse rxf
            ' Use the Deserialize method to restore the object's state with
            ' data from the XML document. 
            cfgs = CType(deserializer.Deserialize(fs), configuration)

            XMLImportConfig = cfgs

        Catch ex As Exception
            XMLImportConfig = Nothing
            Call MsgBox("Lesen der " & cfgXMLfilename & " fehlgeschlagen")
        End Try

        ' ProjectboardConfig.xml-Datei schließen
        fs.Close()

    End Function


    ''' <summary>
    ''' schreibt eine Datei mit den monatlichen Zuordnungen Rollenbedarfe / Kosten 
    ''' Diese Datei kann editiert werden , dann wieder importiert werden 
    ''' in Abhängigkeit vom Typ wird geschrieben: 
    ''' 0: alles
    ''' 1: nur Vergangenheit, von bestimmt den Start , Heute-1 das Ende 
    ''' 2: nur die Zukunft, Heute bestimmt den Start, bis  das Ende  
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub writeProjektBedarfeXLSX(ByVal von As Integer, ByVal bis As Integer, ByVal type As Integer)


        appInstance.EnableEvents = False

        Dim newWB As Excel.Workbook
        Dim rng As Excel.Range
        Dim ersteZeile As Excel.Range
        ' hier muss jetzt das entsprechende File aufgemacht werden ...
        ' das File 
        Try

            newWB = appInstance.Workbooks.Add()

        Catch ex As Exception
            Call MsgBox("Excel Datei konnte nicht erzeugt werden ... Abbruch ")
            appInstance.EnableEvents = True
            Exit Sub
        End Try

        ' jetzt schreiben der ersten Zeile 
        Dim zeile As Integer = 1
        Dim spalte As Integer = 1

        With newWB.ActiveSheet

            ersteZeile = CType(.Range(.cells(1, 1), .cells(1, 6 + bis - von)), Excel.Range)

            CType(.Cells(1, 1), Excel.Range).Value = "Projekt-Name"
            CType(.Cells(1, 2), Excel.Range).Value = "Varianten-Name"
            CType(.Cells(1, 3), Excel.Range).Value = "Phasen-Name"
            CType(.Cells(1, 4), Excel.Range).Value = "Ressourcen-Name"
            CType(.Cells(1, 5), Excel.Range).Value = "Kostenart-Name"


            ' jetzt wird die Zeile 1 geschrieben 
            CType(.Cells(1, 6), Global.Microsoft.Office.Interop.Excel.Range).Value = StartofCalendar.AddMonths(von - 1)
            CType(.Cells(1, 7), Global.Microsoft.Office.Interop.Excel.Range).Value = StartofCalendar.AddMonths(von)
            rng = .Range(.Cells(1, 6), .Cells(1, 7))

            '' Deutsches Format:
            'rng.NumberFormat = "[$-407]mmm yy;@"
            ' Englisches Format:
            rng.NumberFormat = "[$-409]mmm yy;@"

            Dim destinationRange As Excel.Range = .Range(.Cells(1, 6), .Cells(1, 6 + bis - von))
            With destinationRange
                .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                .VerticalAlignment = Excel.XlVAlign.xlVAlignBottom
                '' Deutsches Format: 
                'rng.NumberFormat = "[$-407]mmm yy;@"
                ' Englische Format:
                .NumberFormat = "[$-409]mmm yy;@"
                .WrapText = False
                .Orientation = 90
                .AddIndent = False
                .IndentLevel = 0
                .ReadingOrder = Excel.Constants.xlContext
                .MergeCells = False
            End With

            rng.AutoFill(Destination:=destinationRange, Type:=Excel.XlAutoFillType.xlFillMonths)

        End With



        zeile = 2

        Dim tmpName As String = ""
        Dim tmpValues() As Double
        Dim schnittmenge() As Double
        Dim usedRoles As Collection
        Dim usedCosts As Collection
        Dim pStart As Integer, pEnde As Integer

        Dim editRange As Excel.Range


        For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

            pStart = getColumnOfDate(kvp.Value.startDate)
            pEnde = getColumnOfDate(kvp.Value.endeDate)

            usedRoles = kvp.Value.getRoleNames
            usedCosts = kvp.Value.getCostNames

            For r = 1 To usedRoles.Count
                tmpName = usedRoles.Item(r)
                tmpValues = kvp.Value.getRessourcenBedarf(tmpName)
                schnittmenge = calcArrayIntersection(von, bis, pStart, pEnde, tmpValues)

                If schnittmenge.Sum <> tmpValues.Sum Then
                    Dim a As Integer = 99
                End If

                ' Schreiben der Projekt-Informationen 
                With newWB.ActiveSheet
                    CType(.cells(zeile, 1), Excel.Range).Value = kvp.Value.name
                    CType(.cells(zeile, 2), Excel.Range).Value = kvp.Value.variantName
                    CType(.cells(zeile, 3), Excel.Range).Value = "."
                End With


                With newWB.ActiveSheet
                    CType(.cells(zeile, 4), Excel.Range).Value = tmpName
                    editRange = CType(.range(.cells(zeile, 6), .cells(zeile, 6 + bis - von)), Excel.Range)
                End With

                editRange.Value = schnittmenge
                zeile = zeile + 1

            Next



            For k As Integer = 1 To usedCosts.Count

                tmpName = usedCosts.Item(k)
                tmpValues = kvp.Value.getKostenBedarf(tmpName)
                schnittmenge = calcArrayIntersection(von, bis, pStart, pEnde, tmpValues)

                If schnittmenge.Sum <> tmpValues.Sum Then
                    Dim a As Integer = 99
                End If

                ' Schreiben der Projekt-Informationen 
                With newWB.ActiveSheet
                    CType(.cells(zeile, 1), Excel.Range).Value = kvp.Value.name
                    CType(.cells(zeile, 2), Excel.Range).Value = kvp.Value.variantName
                    CType(.cells(zeile, 3), Excel.Range).Value = "."
                End With

                With newWB.ActiveSheet
                    CType(.cells(zeile, 5), Excel.Range).Value = tmpName
                    editRange = CType(.range(.cells(zeile, 6), .cells(zeile, 6 + bis - von)), Excel.Range)
                End With

                editRange.Value = schnittmenge
                zeile = zeile + 1
            Next

        Next


        ' jetzt den Bereich markieren bzw. schützen 
        Dim startProtectedArea As Integer
        Dim endProtectedArea As Integer
        Dim protectedRange As Excel.Range = Nothing
        Dim wbName As String

        Select Case type
            Case 0
                startProtectedArea = 0
                endProtectedArea = 0
                wbName = "all"
            Case 1

                startProtectedArea = getColumnOfDate(Date.Now)
                endProtectedArea = bis
                wbName = "past"
            Case 2
                startProtectedArea = von
                endProtectedArea = getColumnOfDate(Date.Now)
                wbName = "future"
            Case Else
                Call MsgBox("Typ nicht erkannt, muss Werte 0, 1 oder 2 haben: ist aber" & type)
                appInstance.EnableEvents = True
                Exit Sub
        End Select

        Dim generalRange As Excel.Range = CType(newWB.ActiveSheet.Range(newWB.ActiveSheet.cells(1, 1), _
                                                newWB.ActiveSheet.cells(zeile - 1, 5)),  _
                                                Excel.Range)
        Dim valueRange As Excel.Range = CType(newWB.ActiveSheet.Range(newWB.ActiveSheet.cells(1, 6), _
                                                newWB.ActiveSheet.cells(zeile - 1, 6 + bis - von + 1)),  _
                                                Excel.Range)

        With generalRange
            .Columns.AutoFit()
        End With


        With ersteZeile
            .Interior.Color = awinSettings.AmpelGruen
        End With


        If type <> 0 Then

            With newWB.ActiveSheet
                protectedRange = CType(.Range(.cells(1, startProtectedArea), _
                                                               .cells(zeile - 1, endProtectedArea)),  _
                                                                Excel.Range)

            End With
            protectedRange.Interior.Color = awinSettings.AmpelNichtBewertet
        End If



        Dim expFName As String = exportOrdnerNames(PTImpExp.visbo) & "\EditNeeds_" & _
            Date.Now.ToString.Replace(":", ".") & ".xlsx"

        Try
            appInstance.ActiveWorkbook.SaveAs(Filename:=expFName, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges)
        Catch ex As Exception

        End Try

        Try
            appInstance.ActiveWorkbook.Close(SaveChanges:=False)
        Catch ex As Exception

        End Try

        appInstance.EnableEvents = True

        Call MsgBox("ok, Datei exportiert")

    End Sub

    ''' <summary>
    ''' schreibt eine Datei mit den monatlichen Zuordnungen Projekt/Phase - Rollenbedarfe / Kosten 
    ''' Diese Datei kann editiert werden , dann wieder importiert werden 
    ''' in Abhängigkeit vom Typ wird geschrieben: 
    ''' 0: alles
    ''' 1: nur Vergangenheit, von bestimmt den Start , Heute-1 das Ende 
    ''' 2: nur die Zukunft, Heute bestimmt den Start, bis  das Ende  
    ''' </summary>
    ''' <param name="von"></param>
    ''' <param name="bis"></param>
    ''' <param name="type"></param>
    ''' <remarks></remarks>
    Public Sub writeProjektPhasenBedarfeXLSX(ByVal von As Integer, ByVal bis As Integer, ByVal type As Integer)


        appInstance.EnableEvents = False

        Dim newWB As Excel.Workbook
        Dim ersteZeile As Excel.Range
        Dim ressCostColumn As Integer

        Dim expFName As String = exportOrdnerNames(PTImpExp.massenEdit) & "\EditNeeds_" & _
        Date.Now.ToString.Replace(":", ".") & ".xlsx"

        ' hier muss jetzt das entsprechende File aufgemacht werden ...
        ' das File 
        Try

            newWB = appInstance.Workbooks.Add()

            CType(newWB.Worksheets.Item(1), Excel.Worksheet).Name = "VISBO"
            If newWB.Worksheets.Count < 2 Then
                newWB.Worksheets.Add(After:=newWB.Worksheets("VISBO"))
                CType(appInstance.ActiveSheet, Excel.Worksheet).Name = "tmp"

            Else
                CType(newWB.Worksheets.Item(2), Excel.Worksheet).Name = "tmp"
            End If

            newWB.SaveAs(expFName)

        Catch ex As Exception
            Call MsgBox("Excel Datei konnte nicht erzeugt werden ... Abbruch ")
            appInstance.EnableEvents = True
            Exit Sub
        End Try

        ' jetzt schreiben der ersten Zeile 
        Dim zeile As Integer = 1
        Dim spalte As Integer = 1

        'Dim startSpalteDaten As Integer = 8
        Dim startSpalteDaten As Integer = 8
        Dim roleCostNames As Excel.Range = Nothing
        Dim roleCostInput As Excel.Range = Nothing

        Dim tmpName As String = ""

        ' hier werden jetzt erst mal die Ressourcen und Kostenarten geschrieben 
        With CType(newWB.Worksheets("tmp"), Excel.Worksheet)

            Dim sortedRCListe As New SortedList(Of String, String)
            For iz As Integer = 1 To RoleDefinitions.Count
                tmpName = RoleDefinitions.getRoledef(iz).name
                If Not sortedRCListe.ContainsKey(tmpName) Then
                    sortedRCListe.Add(tmpName, tmpName)
                End If
            Next
            For iz As Integer = 1 To sortedRCListe.Count
                tmpName = sortedRCListe.ElementAt(iz - 1).Value
                CType(.Cells(iz, 1), Excel.Range).Value = tmpName
            Next

            Dim offz As Integer = sortedRCListe.Count
            sortedRCListe.Clear()

            For iz As Integer = 1 To CostDefinitions.Count - 1
                tmpName = CostDefinitions.getCostdef(iz).name
                If Not sortedRCListe.ContainsKey(tmpName) Then
                    sortedRCListe.Add(tmpName, tmpName)
                End If
            Next

            For iz As Integer = 1 To sortedRCListe.Count
                tmpName = sortedRCListe.ElementAt(iz - 1).Value
                CType(.Cells(iz + offz, 1), Excel.Range).Value = tmpName
            Next

            offz = offz + sortedRCListe.Count
            roleCostNames = CType(.Range(.Cells(1, 1), .Cells(offz, 1)), Excel.Range)
            newWB.Names.Add(Name:="RollenKostenNamen", RefersToR1C1:=roleCostNames)
            CType(newWB.Worksheets("tmp"), Excel.Worksheet).Visible = False    ' Worksheet "tmp" ausblenden

        End With



        With CType(newWB.Worksheets("VISBO"), Excel.Worksheet)

            ersteZeile = CType(.Range(.Cells(1, 1), .Cells(1, 6 + bis - von)), Excel.Range)

            CType(.Cells(1, 1), Excel.Range).Value = "Business-Unit"
            CType(.Cells(1, 2), Excel.Range).Value = "Projekt-Name"
            CType(.Cells(1, 3), Excel.Range).Value = "Varianten-Name"
            CType(.Cells(1, 4), Excel.Range).Value = "Phasen-Name"
            CType(.Cells(1, 5), Excel.Range).Value = "Ress./Kostenart-Name"
            CType(.Cells(1, 6), Excel.Range).Value = "Summe"
            CType(.Cells(1, 7), Excel.Range).Value = "Proz."
            'CType(.Cells(1, 7), Excel.Range).Value = "Kostenart-Name"

            ' jetzt wird die Spalten-Nummer festgelegt, wo die Ressourcen/ Kosten später eingetragen werden
            ressCostColumn = 5
            ' jetzt wird die Zeile 1 geschrieben 
            Dim startMonat As Date = StartofCalendar.AddMonths(von - 1)

            ' jetzt wird der Name hinzugefügt
            Dim tmpRange1 As Excel.Range = CType(.Cells(1, startSpalteDaten), Global.Microsoft.Office.Interop.Excel.Range)
            Dim tmpRange2 As Excel.Range = CType(.Cells(1, startSpalteDaten + 2 * (bis - von)), Global.Microsoft.Office.Interop.Excel.Range)
            newWB.Names.Add(Name:="StartData", RefersToR1C1:=tmpRange1)
            newWB.Names.Add(Name:="EndData", RefersToR1C1:=tmpRange2)

            ' jetzt werden die Überschriften des Datenbereichs geschrieben 
            For m As Integer = 0 To bis - von
                With CType(.Cells(1, startSpalteDaten + 2 * m), Global.Microsoft.Office.Interop.Excel.Range)
                    .Value = startMonat.AddMonths(m)
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .VerticalAlignment = Excel.XlVAlign.xlVAlignBottom
                    .NumberFormat = "[$-409]mmm yy;@"
                    .WrapText = False
                    .Orientation = 90
                    .AddIndent = False
                    .IndentLevel = 0
                    .ReadingOrder = Excel.Constants.xlContext
                End With

                With CType(.Cells(1, startSpalteDaten + 2 * m + 1), Global.Microsoft.Office.Interop.Excel.Range)
                    .Value = ""
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                    .Orientation = 0
                    .AddIndent = False
                    .IndentLevel = 0
                    .ReadingOrder = Excel.Constants.xlContext
                End With

            Next

            ' bevor die Prozentualen Anteile ergänzt wurden ... 
            ' ''CType(.Cells(1, startSpalteDaten), Global.Microsoft.Office.Interop.Excel.Range).Value = StartofCalendar.AddMonths(von - 1)
            ' ''CType(.Cells(1, startSpalteDaten + 1), Global.Microsoft.Office.Interop.Excel.Range).Value = StartofCalendar.AddMonths(von)
            ' ''rng = .Range(.Cells(1, startSpalteDaten), .Cells(1, startSpalteDaten + 1))

            ' '' '' Deutsches Format:
            '' ''rng.NumberFormat = "[$-407]mmm yy;@"
            '' '' Englisches Format:
            ' ''rng.NumberFormat = "[$-409]mmm yy;@"

            ' ''Dim destinationRange As Excel.Range = .Range(.Cells(1, 6), .Cells(1, 6 + bis - von))
            ' ''With destinationRange
            ' ''    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            ' ''    .VerticalAlignment = Excel.XlVAlign.xlVAlignBottom
            ' ''    '' Deutsches Format: 
            ' ''    'rng.NumberFormat = "[$-407]mmm yy;@"
            ' ''    ' Englische Format:
            ' ''    .NumberFormat = "[$-409]mmm yy;@"
            ' ''    .WrapText = False
            ' ''    .Orientation = 90
            ' ''    .AddIndent = False
            ' ''    .IndentLevel = 0
            ' ''    .ReadingOrder = Excel.Constants.xlContext
            ' ''    .MergeCells = False
            ' ''End With

            ' ''rng.AutoFill(Destination:=destinationRange, Type:=Excel.XlAutoFillType.xlFillMonths)

        End With

        zeile = 2

        Dim schnittmenge() As Double
        Dim zeilenWerte() As Double
        Dim zeilensumme As Double
        Dim pStart As Integer, pEnde As Integer

        Dim editRange As Excel.Range


        ' zu Beginn werden die rollen-spezifischen Auslastungskennzahlen ermittelt, die sich über alle aktuell 
        ' betrachteten Projekte ergeben; 
        ' es werden sowohl die Gesamt-Auslastungs Werte im Zeitraum betrachtet als auch der einzelne monats-spezifische Wert   
        ' dazu wird ein Array angelegt mit der Dimension (anzahlRollen-1, bis-von+1) 
        Dim auslastungsArray(,) As Double

        Try
            auslastungsArray = visboZustaende.getUpDatedAuslastungsArray(Nothing, von, bis, awinSettings.mePrzAuslastung)
            'auslastungsArray = ShowProjekte.getAuslastungsArray(von, bis)
        Catch ex As Exception
            ReDim auslastungsArray(RoleDefinitions.Count - 1, bis - von + 1)
        End Try




        For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

            pStart = getColumnOfDate(kvp.Value.startDate)
            pEnde = getColumnOfDate(kvp.Value.endeDate)

            For p = 1 To kvp.Value.CountPhases

                Dim cphase As clsPhase = kvp.Value.getPhase(p)
                Dim phaseNameID As String = cphase.nameID
                Dim phaseName As String = cphase.name
                Dim chckNameID As String = calcHryElemKey(phaseName, False)

                If phaseWithinTimeFrame(pStart, cphase.relStart, cphase.relEnde, von, bis) Then
                    ' nur wenn die Phase überhaupt im betrachteten Zeitraum liegt, muss das berücksichtigt werden 

                    ' jetzt müssen die Zellen, die zur Phase gehören , entsperrt werden  ...
                    Dim ixZeitraum As Integer
                    Dim ix As Integer, breite As Integer

                    Dim atLeastOne As Boolean = False

                    Call awinIntersectZeitraum(pStart + cphase.relStart - 1, pStart + cphase.relEnde - 1, ixZeitraum, ix, breite)


                    For r = 1 To cphase.countRoles


                        Dim role As clsRolle = cphase.getRole(r)
                        Dim roleName As String = role.name
                        Dim roleUID As Integer = RoleDefinitions.getRoledef(roleName).UID
                        Dim xValues() As Double = role.Xwerte

                        schnittmenge = calcArrayIntersection(von, bis, pStart + cphase.relStart - 1, pStart + cphase.relEnde - 1, xValues)
                        zeilensumme = schnittmenge.Sum

                        ReDim zeilenWerte(2 * (bis - von + 1) - 1)

                        ' Schreiben der Projekt-Informationen 
                        With CType(newWB.Worksheets("VISBO"), Excel.Worksheet)
                            CType(.Cells(zeile, 1), Excel.Range).Value = kvp.Value.businessUnit
                            CType(.Cells(zeile, 2), Excel.Range).Value = kvp.Value.name
                            CType(.Cells(zeile, 3), Excel.Range).Value = kvp.Value.variantName
                            CType(.Cells(zeile, 4), Excel.Range).Value = cphase.name

                            Dim cellComment As Excel.Comment = CType(.Cells(zeile, 4), Excel.Range).Comment
                            If Not IsNothing(cellComment) Then
                                CType(.Cells(zeile, 4), Excel.Range).Comment.Delete()
                            End If
                            If chckNameID = phaseNameID Then
                                ' nichts weiter tun ... 
                                ' denn dann kann die PhaseNameID aus der PhaseName konstruiert werden
                                ' wenn es eine laufende Nummer 2, 3 etc ist, dann muss explizit die PhaseNameID in den Kommentarbereich geschreiben werden 
                            Else
                                CType(.Cells(zeile, 4), Excel.Range).AddComment(Text:=cphase.nameID)
                                CType(.Cells(zeile, 4), Excel.Range).Comment.Visible = False
                            End If

                            CType(.Cells(zeile, 5), Excel.Range).Value = roleName
                            CType(.Cells(zeile, 6), Excel.Range).Value = zeilensumme.ToString("0")
                            CType(.Cells(zeile, 7), Excel.Range).Value = auslastungsArray(roleUID - 1, 0).ToString("0%")
                            editRange = CType(.Range(.Cells(zeile, startSpalteDaten), .Cells(zeile, startSpalteDaten + 2 * (bis - von + 1) - 1)), Excel.Range)
                        End With

                        ' zusammenmischen von Schnittmenge und Prozentual-Werte 
                        For mis As Integer = 0 To bis - von
                            zeilenWerte(2 * mis) = schnittmenge(mis)
                            ' in auslastungsarray(r, 0) steht die Gesamt-Auslastung
                            zeilenWerte(2 * mis + 1) = auslastungsArray(roleUID - 1, mis + 1)
                        Next

                        'editRange.Value = schnittmenge
                        editRange.Value = zeilenWerte
                        atLeastOne = True
                        ' die Zellen entsperren, die editiert werden dürfen ...

                        With CType(newWB.Worksheets("VISBO"), Excel.Worksheet)

                            For l = 0 To bis - von

                                If l >= ixZeitraum And l <= ixZeitraum + breite - 1 Then
                                    'CType(.Cells(zeile, 2 * l + startSpalteDaten), Excel.Range).Locked = False
                                    CType(.Range(.Cells(zeile, 2 * l + startSpalteDaten), _
                                                 .Cells(zeile, 2 * l + 1 + startSpalteDaten)), Excel.Range).Interior.Color = awinSettings.AmpelNichtBewertet
                                Else
                                    CType(.Cells(zeile, 2 * l + startSpalteDaten), Excel.Range).Value = ""
                                End If

                            Next

                            ' vorheriger Code
                            ''For l As Integer = ixZeitraum To ixZeitraum + breite - 1
                            ''    CType(.cell(zeile, l + 6), Excel.Range).Locked = False
                            ''    CType(.cell(zeile, l + 6), Excel.Range).Interior.Color = awinSettings.AmpelNichtBewertet
                            ''Next
                        End With


                        ''With newWB.ActiveSheet
                        ''    For l As Integer = ixZeitraum To ixZeitraum + breite - 1
                        ''        CType(.cells(zeile, l + 6), Excel.Range).Locked = False
                        ''    Next
                        ''End With

                        zeile = zeile + 1

                    Next r

                    For c = 1 To cphase.countCosts
                        Dim cost As clsKostenart = cphase.getCost(c)
                        Dim xValues() As Double = cost.Xwerte
                        Dim costName As String = cost.name
                        schnittmenge = calcArrayIntersection(von, bis, pStart + cphase.relStart - 1, pStart + cphase.relEnde - 1, xValues)
                        zeilensumme = schnittmenge.Sum

                        ReDim zeilenWerte(2 * (bis - von + 1) - 1)

                        ' Schreiben der Projekt-Informationen 
                        With CType(newWB.Worksheets("VISBO"), Excel.Worksheet)
                            CType(.Cells(zeile, 1), Excel.Range).Value = kvp.Value.businessUnit
                            CType(.Cells(zeile, 2), Excel.Range).Value = kvp.Value.name
                            CType(.Cells(zeile, 3), Excel.Range).Value = kvp.Value.variantName
                            CType(.Cells(zeile, 4), Excel.Range).Value = cphase.name

                            Dim cellComment As Excel.Comment = CType(.Cells(zeile, 4), Excel.Range).Comment
                            If Not IsNothing(cellComment) Then
                                CType(.Cells(zeile, 4), Excel.Range).Comment.Delete()
                            End If
                            If chckNameID = phaseNameID Then
                                ' nichts weiter tun ... 
                                ' denn dann kann die PhaseNameID aus der PhaseName konstruiert werden
                                ' wenn es eine laufende Nummer 2, 3 etc ist, dann muss explizit die PhaseNameID in den Kommentarbereich geschreiben werden 
                            Else
                                CType(.Cells(zeile, 4), Excel.Range).AddComment(Text:=cphase.nameID)
                                CType(.Cells(zeile, 4), Excel.Range).Comment.Visible = False
                            End If

                            CType(.Cells(zeile, 5), Excel.Range).Value = costName
                            CType(.Cells(zeile, 6), Excel.Range).Value = zeilensumme.ToString("0")
                            editRange = CType(.Range(.Cells(zeile, startSpalteDaten), .Cells(zeile, startSpalteDaten + 2 * (bis - von + 1) - 1)), Excel.Range)
                        End With

                        ' zusammenmischen von Schnittmenge und Prozentual-Werte 
                        For mis As Integer = 0 To bis - von
                            zeilenWerte(2 * mis) = schnittmenge(mis)
                            ' in auslastungsarray(r, 0) steht die Gesamt-Auslastung, spielt aber kein Kostenarten keine Rolle 
                            zeilenWerte(2 * mis + 1) = 0
                        Next

                        'editRange.Value = schnittmenge
                        editRange.Value = zeilenWerte
                        atLeastOne = True
                        ' die Zellen entsperren, die editiert werden dürfen ...

                        ' die Zellen entsperren, die editiert werden dürfen ...

                        With CType(newWB.Worksheets("VISBO"), Excel.Worksheet)

                            For l = 0 To bis - von

                                If l >= ixZeitraum And l <= ixZeitraum + breite - 1 Then
                                    'CType(.Cells(zeile, 2 * l + startSpalteDaten), Excel.Range).Locked = False
                                    CType(.Range(.Cells(zeile, 2 * l + startSpalteDaten), _
                                                 .Cells(zeile, 2 * l + 1 + startSpalteDaten)), Excel.Range).Interior.Color = awinSettings.AmpelNichtBewertet
                                    CType(.Cells(zeile, 2 * l + 1 + startSpalteDaten), Excel.Range).Value = ""
                                Else
                                    CType(.Cells(zeile, 2 * l + startSpalteDaten), Excel.Range).Value = ""
                                    CType(.Cells(zeile, 2 * l + 1 + startSpalteDaten), Excel.Range).Value = ""
                                End If

                            Next

                        End With

                        zeile = zeile + 1

                    Next c

                    If Not atLeastOne Then
                        ' jetzt sollte eine leere Projekt-Phasen-Information geschrieben werden, quasi ein Platzhalter
                        ' in diesem Platzhalter kann dann später die Ressourcen Information aufgenommen werden  
                        ' Schreiben der Projekt-Informationen 
                        With CType(newWB.Worksheets("VISBO"), Excel.Worksheet)
                            CType(.Cells(zeile, 1), Excel.Range).Value = kvp.Value.businessUnit
                            CType(.Cells(zeile, 2), Excel.Range).Value = kvp.Value.name
                            CType(.Cells(zeile, 3), Excel.Range).Value = kvp.Value.variantName
                            CType(.Cells(zeile, 4), Excel.Range).Value = cphase.name

                            Dim cellComment As Excel.Comment = CType(.Cells(zeile, 4), Excel.Range).Comment
                            If Not IsNothing(cellComment) Then
                                CType(.Cells(zeile, 4), Excel.Range).Comment.Delete()
                            End If
                            If chckNameID = phaseNameID Then
                                ' nichts weiter tun ... 
                                ' denn dann kann die PhaseNameID aus der PhaseName konstruiert werden
                                ' wenn es eine laufende Nummer 2, 3 etc ist, dann muss explizit die PhaseNameID in den Kommentarbereich geschreiben werden 
                            Else
                                CType(.Cells(zeile, 4), Excel.Range).AddComment(Text:=cphase.nameID)
                                CType(.Cells(zeile, 4), Excel.Range).Comment.Visible = False
                            End If

                            CType(.Cells(zeile, 5), Excel.Range).Value = ""
                            CType(.Cells(zeile, 6), Excel.Range).Value = ""
                            CType(.Cells(zeile, 7), Excel.Range).Value = ""
                            editRange = CType(.Range(.Cells(zeile, startSpalteDaten), .Cells(zeile, startSpalteDaten + 2 * (bis - von))), Excel.Range)
                        End With

                        ' die Zellen entsperren, die editiert werden dürfen ...
                        With CType(newWB.Worksheets("VISBO"), Excel.Worksheet)

                            For l = 0 To bis - von

                                If l >= ixZeitraum And l <= ixZeitraum + breite - 1 Then
                                    'CType(.Cells(zeile, 2 * l + startSpalteDaten), Excel.Range).Locked = False
                                    CType(.Range(.Cells(zeile, 2 * l + startSpalteDaten), _
                                                 .Cells(zeile, 2 * l + 1 + startSpalteDaten)), Excel.Range).Interior.Color = awinSettings.AmpelNichtBewertet
                                Else
                                    CType(.Cells(zeile, 2 * l + startSpalteDaten), Excel.Range).Value = ""
                                End If

                            Next

                        End With

                        zeile = zeile + 1

                    End If

                End If



            Next p



        Next


        ' jetzt die Größe der Spalten anpassen 
        Dim infoBlock As Excel.Range
        With CType(newWB.Worksheets("VISBO"), Excel.Worksheet)
            infoBlock = CType(.Range(.Columns(1), .Columns(startSpalteDaten - 1)), Excel.Range)
            infoBlock.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            infoBlock.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            infoBlock.AutoFit()
        End With

        Dim tmpRange As Excel.Range
        With CType(newWB.Worksheets("VISBO"), Excel.Worksheet)

            Dim isPrz As Boolean = False
            For mis As Integer = 0 To 2 * (bis - von + 1) - 1
                tmpRange = CType(.Range(.Cells(2, startSpalteDaten + mis), .Cells(zeile, startSpalteDaten + mis)), Excel.Range)
                If isPrz Then
                    tmpRange.Columns.ColumnWidth = 3.1
                    tmpRange.Font.Size = 6
                    tmpRange.NumberFormat = "0%"
                    tmpRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                Else
                    tmpRange.Columns.ColumnWidth = 5
                    tmpRange.Font.Size = 10
                    tmpRange.NumberFormat = "0"
                    tmpRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End If
                isPrz = Not isPrz
            Next

        End With



        ' jetzt wird der RoleCostInput Bereich festgelegt 
        With CType(newWB.Worksheets("VISBO"), Excel.Worksheet)
            Dim maxRows As Integer = .Rows.Count
            roleCostInput = CType(.Range(.Cells(2, ressCostColumn), .Cells(maxRows, ressCostColumn)), Excel.Range)
        End With

        With roleCostInput
            .Validation.Delete()
            .Validation.Add(Type:=XlDVType.xlValidateList, AlertStyle:=XlDVAlertStyle.xlValidAlertStop, _
                                           Formula1:="=RollenKostenNamen")
        End With



        Try
            ' jetzt die Autofilter aktivieren ... 
            If Not CType(newWB.Worksheets("VISBO"), Excel.Worksheet).AutoFilterMode = True Then
                'CType(CType(newWB.Worksheets("VISBO"), Excel.Worksheet).Cells(1, 1), Excel.Range).Select()
                CType(newWB.Worksheets("VISBO"), Excel.Worksheet).Cells(1, 1).AutoFilter()
            End If

            ' ExcelFile abspeichern und schließen
            newWB.Close(SaveChanges:=True)
        Catch ex As Exception
            Throw New ArgumentException("Fehler beim Filtersetzen und Speichern" & ex.Message)
        End Try

        appInstance.EnableEvents = True

        Call MsgBox("ok, Datei exportiert")

    End Sub

    ''' <summary>
    ''' schreibt eine Datei, die zur Priorisierung verwendet werdne kann 
    ''' Diese Datei kann editiert werden , dann wieder importiert werden 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub writeProjektsForSequencing()


        appInstance.EnableEvents = False

        Dim newWB As Excel.Workbook

        Dim expFName As String = exportOrdnerNames(PTImpExp.scenariodefs) & "\" & currentConstellationName & "_Prio.xlsx"

        ' hier muss jetzt das entsprechende File aufgemacht werden ...
        ' das File 
        Try

            newWB = appInstance.Workbooks.Add()
            CType(newWB.Worksheets.Item(1), Excel.Worksheet).Name = "VISBO"
            newWB.SaveAs(Filename:=expFName, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges)

        Catch ex As Exception
            Call MsgBox("Excel Datei konnte nicht erzeugt werden ... Abbruch ")
            appInstance.EnableEvents = True
            Exit Sub
        End Try

        ' jetzt schreiben der ersten Zeile 
        Dim zeile As Integer = 1
        Dim spalte As Integer = 1




        With CType(newWB.Worksheets("VISBO"), Excel.Worksheet)

            If awinSettings.englishLanguage Then
                CType(.Cells(1, 1), Excel.Range).Value = "Project-Name"
                CType(.Cells(1, 2), Excel.Range).Value = "Variant-Name"
                CType(.Cells(1, 3), Excel.Range).Value = "Business-Unit"
                CType(.Cells(1, 4), Excel.Range).Value = "Budget"
                CType(.Cells(1, 5), Excel.Range).Value = "Project-Start"
                CType(.Cells(1, 6), Excel.Range).Value = "Project-End"
                CType(.Cells(1, 7), Excel.Range).Value = "Sum Personnel-Cost [T€]"
                CType(.Cells(1, 8), Excel.Range).Value = "Sum Other Cost [T€]"
                CType(.Cells(1, 9), Excel.Range).Value = "Profit/Loss"
                CType(.Cells(1, 10), Excel.Range).Value = "Strategye"
                CType(.Cells(1, 11), Excel.Range).Value = "Risk"
            Else

                CType(.Cells(1, 1), Excel.Range).Value = "Projekt-Name"
                CType(.Cells(1, 2), Excel.Range).Value = "Varianten-Name"
                CType(.Cells(1, 3), Excel.Range).Value = "Business-Unit"
                CType(.Cells(1, 4), Excel.Range).Value = "Budget"
                CType(.Cells(1, 5), Excel.Range).Value = "Projekt-Start"
                CType(.Cells(1, 6), Excel.Range).Value = "Projekt-Ende"
                CType(.Cells(1, 7), Excel.Range).Value = "Summe Personalkosten [T€]"
                CType(.Cells(1, 8), Excel.Range).Value = "Summe sonst. Kosten [T€]"
                CType(.Cells(1, 9), Excel.Range).Value = "Profit/Loss"
                CType(.Cells(1, 10), Excel.Range).Value = "Strategie"
                CType(.Cells(1, 11), Excel.Range).Value = "Risiko"

                spalte = 12
                For Each cstField As KeyValuePair(Of Integer, clsCustomFieldDefinition) In customFieldDefinitions.liste
                    .Cells(zeile, spalte).value = cstField.Value.name
                    spalte = spalte + 1
                Next
            End If


        End With

        zeile = 2


        For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

            Dim budget As Double, pk As Double, ok As Double, rk As Double, pl As Double

            Call kvp.Value.calculateRoundedKPI(budget, pk, ok, rk, pl)

            With CType(newWB.Worksheets("VISBO"), Excel.Worksheet)
                CType(.Cells(zeile, 1), Excel.Range).Value = kvp.Value.name
                CType(.Cells(zeile, 2), Excel.Range).Value = kvp.Value.variantName
                CType(.Cells(zeile, 3), Excel.Range).Value = kvp.Value.businessUnit
                CType(.Cells(zeile, 4), Excel.Range).Value = budget
                CType(.Cells(zeile, 5), Excel.Range).Value = kvp.Value.startDate
                CType(.Cells(zeile, 6), Excel.Range).Value = kvp.Value.endeDate
                CType(.Cells(zeile, 7), Excel.Range).Value = pk
                CType(.Cells(zeile, 8), Excel.Range).Value = ok
                CType(.Cells(zeile, 9), Excel.Range).Value = pl
                CType(.Cells(zeile, 10), Excel.Range).Value = kvp.Value.StrategicFit
                CType(.Cells(zeile, 11), Excel.Range).Value = kvp.Value.Risiko

                spalte = 12
                For Each cstField As KeyValuePair(Of Integer, clsCustomFieldDefinition) In customFieldDefinitions.liste

                    Dim qualifier As String = cstField.Value.name
                    Dim ausgabe As String = ""
                    If cstField.Value.type = ptCustomFields.Str Then
                        ausgabe = kvp.Value.getCustomSField(qualifier)
                    ElseIf cstField.Value.type = ptCustomFields.Dbl Then
                        ausgabe = kvp.Value.getCustomDField(qualifier).ToString
                    ElseIf cstField.Value.type = ptCustomFields.bool Then
                        ausgabe = kvp.Value.getCustomBField(qualifier).ToString
                    End If

                    If IsNothing(ausgabe) Then
                        ausgabe = ""
                    End If

                    CType(.Cells(zeile, spalte), Excel.Range).Value = ausgabe
                    spalte = spalte + 1
                Next

            End With
            zeile = zeile + 1
        Next


        Try
            ' jetzt die Autofilter aktivieren ... 
            If Not CType(newWB.Worksheets("VISBO"), Excel.Worksheet).AutoFilterMode = True Then
                'CType(CType(newWB.Worksheets("VISBO"), Excel.Worksheet).Cells(1, 1), Excel.Range).Select()
                CType(newWB.Worksheets("VISBO"), Excel.Worksheet).Cells(1, 1).AutoFilter()
            End If

            ' ExcelFile abspeichern und schließen
            newWB.Close(SaveChanges:=True)
        Catch ex As Exception
            Throw New ArgumentException("Fehler beim Filtersetzen und Speichern" & ex.Message)
        End Try

        appInstance.EnableEvents = True

        Call MsgBox("ok, Datei exportiert")

    End Sub


    ''' <summary>
    ''' erstellt für alles, alleRollen, alleKosten und die einzelnen Sammel-Rollen die ValidationStrings
    ''' die dann im Mass-Edit verwendet werden können 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function createMassEditRcValidations() As SortedList(Of String, String)
        Dim validationStrings As New SortedList(Of String, String)
        Dim validationName As String

        ' Aufbau Alles
        validationName = "alles"
        Dim sortedRCListe As New SortedList(Of String, String)
        Dim rcDefinition As String = ""
        Dim tmpName As String

        For iz As Integer = 1 To RoleDefinitions.Count
            tmpName = RoleDefinitions.getRoledef(iz).name
            If Not sortedRCListe.ContainsKey(tmpName) Then
                sortedRCListe.Add(tmpName, tmpName)
            End If
        Next

        For iz As Integer = 1 To sortedRCListe.Count
            If rcDefinition.Length = 0 Then
                rcDefinition = sortedRCListe.ElementAt(iz - 1).Value
            Else
                rcDefinition = rcDefinition & ";" & sortedRCListe.ElementAt(iz - 1).Value
            End If
        Next

        sortedRCListe.Clear()

        For iz As Integer = 1 To CostDefinitions.Count - 1
            tmpName = CostDefinitions.getCostdef(iz).name
            If Not sortedRCListe.ContainsKey(tmpName) Then
                sortedRCListe.Add(tmpName, tmpName)
            End If
        Next

        For iz As Integer = 1 To sortedRCListe.Count
            If rcDefinition.Length = 0 Then
                rcDefinition = sortedRCListe.ElementAt(iz - 1).Value
            Else
                rcDefinition = rcDefinition & ";" & sortedRCListe.ElementAt(iz - 1).Value
            End If
        Next

        If Not validationStrings.ContainsKey(validationName) Then
            validationStrings.Add(validationName, rcDefinition)
        End If


        '
        ' jetzt kommen alleRollen 

        validationName = "alleRollen"
        sortedRCListe = New SortedList(Of String, String)
        rcDefinition = ""

        For iz As Integer = 1 To RoleDefinitions.Count
            tmpName = RoleDefinitions.getRoledef(iz).name
            If Not sortedRCListe.ContainsKey(tmpName) Then
                sortedRCListe.Add(tmpName, tmpName)
            End If
        Next

        For iz As Integer = 1 To sortedRCListe.Count
            If rcDefinition.Length = 0 Then
                rcDefinition = sortedRCListe.ElementAt(iz - 1).Value
            Else
                rcDefinition = rcDefinition & ";" & sortedRCListe.ElementAt(iz - 1).Value
            End If
        Next

        If Not validationStrings.ContainsKey(validationName) Then
            validationStrings.Add(validationName, rcDefinition)
        End If

        ' Ende alleRollen
        '

        '
        ' jetzt kommen alleKosten 

        validationName = "alleKosten"
        sortedRCListe = New SortedList(Of String, String)
        rcDefinition = ""

        For iz As Integer = 1 To CostDefinitions.Count - 1
            tmpName = CostDefinitions.getCostdef(iz).name
            If Not sortedRCListe.ContainsKey(tmpName) Then
                sortedRCListe.Add(tmpName, tmpName)
            End If
        Next

        For iz As Integer = 1 To sortedRCListe.Count
            If rcDefinition.Length = 0 Then
                rcDefinition = sortedRCListe.ElementAt(iz - 1).Value
            Else
                rcDefinition = rcDefinition & ";" & sortedRCListe.ElementAt(iz - 1).Value
            End If
        Next

        If Not validationStrings.ContainsKey(validationName) Then
            validationStrings.Add(validationName, rcDefinition)
        End If

        ' Ende alleKosten
        '

        '
        ' jetzt kommen die einzelnen Sammelrollen, unter Angabe ihres Namens 

        Dim sammelrollenNamen As Collection = RoleDefinitions.getSummaryRoles

        For iz As Integer = 1 To sammelrollenNamen.Count
            tmpName = CStr(sammelrollenNamen.Item(iz))
            If Not sortedRCListe.ContainsKey(tmpName) Then
                sortedRCListe.Add(tmpName, tmpName)
            End If
        Next

        For Each validationName In sammelrollenNamen

            sortedRCListe = New SortedList(Of String, String)
            rcDefinition = ""

            For iz As Integer = 1 To sammelrollenNamen.Count
                tmpName = CStr(sammelrollenNamen.Item(iz))
                If Not sortedRCListe.ContainsKey(tmpName) Then
                    sortedRCListe.Add(tmpName, tmpName)
                End If
            Next

            Dim subRoleNames As Collection = RoleDefinitions.getSubRoleNamesOf(validationName, PTcbr.all)

            For iz As Integer = 1 To subRoleNames.Count
                tmpName = CStr(subRoleNames.Item(iz))
                If Not sortedRCListe.ContainsKey(tmpName) Then
                    sortedRCListe.Add(tmpName, tmpName)
                End If
            Next

            For iz As Integer = 1 To sortedRCListe.Count
                If rcDefinition.Length = 0 Then
                    rcDefinition = sortedRCListe.ElementAt(iz - 1).Value
                Else
                    rcDefinition = rcDefinition & ";" & sortedRCListe.ElementAt(iz - 1).Value
                End If
            Next

            ' jetzt den Validation String hinzufügen 
            If Not validationStrings.ContainsKey(validationName) Then
                validationStrings.Add(validationName, rcDefinition)
            End If

        Next

        createMassEditRcValidations = validationStrings
    End Function

    ''' <summary>
    ''' schreibt die Daten der in einer todoListe übergebenen Projekt-Namen in ein extra Tabellenblatt 
    ''' die Info-Daten werden in einer Range mit Name informationColumns zusammengefasst   
    ''' </summary>
    ''' <param name="von"></param>
    ''' <param name="bis"></param>
    ''' <remarks></remarks>
    Public Sub writeOnlineMassEditRessCost(ByVal todoListe As Collection, _
                                           ByVal von As Integer, ByVal bis As Integer)

        Dim mahleRange As Excel.Range

        If todoListe.Count = 0 Then
            If awinSettings.englishLanguage Then
                Call MsgBox("no projects for mass-edit available ..")
            Else
                Call MsgBox("keine Projekte für den Massen-Edit vorhanden ..")
            End If

            Exit Sub
        End If

        Try

            appInstance.EnableEvents = False

            ' jetzt die selectedProjekte Liste zurücksetzen ... ohne die currentConstellation zu verändern ...
            selectedProjekte.Clear(False)

            Dim currentWS As Excel.Worksheet
            Dim currentWB As Excel.Workbook
            Dim ersteZeile As Excel.Range
            Dim ressCostColumn As Integer
            Dim tmpName As String

            ' jetzt werden die Validation-Strings für alles, alleRollen, alleKosten und die einzelnen SammelRollen aufgebaut 
            Dim validationStrings As SortedList(Of String, String) = createMassEditRcValidations()
            Dim anzahlRollen As Integer = RoleDefinitions.Count
            Dim rcValidation() As String
            ' in rcValidation(0) steht der Name "alleKosten" für den Validation-String für alle Kosten
            ' in rcValidation(i) steht der Name des Validation-String für Rolle mit UID i 
            ReDim rcValidation(anzahlRollen + 1)

            rcValidation(0) = "alleKosten"
            rcValidation(anzahlRollen + 1) = "alles"

            For i = 1 To anzahlRollen
                Dim tmprole As clsRollenDefinition = RoleDefinitions.getRoledef(i)
                If tmprole.isCombinedRole Then
                    rcValidation(i) = tmprole.name
                Else
                    Dim parentName As String = RoleDefinitions.getParentRoleOf(tmprole.name)
                    If parentName = "" Then
                        rcValidation(i) = "alleRollen"
                    Else
                        rcValidation(i) = parentName
                    End If
                End If
            Next

            ' hier muss jetzt das entsprechende File aufgemacht werden ...
            ' das File 
            Try
                currentWB = CType(appInstance.Workbooks.Item(myProjektTafel), Excel.Workbook)
                currentWS = CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.meRC)), Excel.Worksheet)

                Try
                    ' off setzen des AutoFilter Modus ... 
                    If CType(currentWS, Excel.Worksheet).AutoFilterMode = True Then
                        'CType(CType(currentWS, Excel.Worksheet).Cells(1, 1), Excel.Range).Select()
                        CType(currentWS, Excel.Worksheet).Cells(1, 1).AutoFilter()
                    End If
                Catch ex As Exception

                End Try

                ' braucht man eigentlich nicht mehr, aber sicher ist sicher ...
                Try
                    currentWS.UsedRange.Clear()
                Catch ex As Exception

                End Try


            Catch ex As Exception
                Call MsgBox("es gibt Probleme mit dem Mass-Edit Worksheet ...")
                appInstance.EnableEvents = True
                Exit Sub
            End Try


            ' jetzt schreiben der ersten Zeile 
            Dim zeile As Integer = 1
            Dim spalte As Integer = 1

            'Dim startSpalteDaten As Integer = 8
            Dim startSpalteDaten As Integer = 8
            'Dim roleCostNames As Excel.Range = Nothing
            Dim roleCostInput As Excel.Range = Nothing

            tmpName = ""

            With CType(currentWS, Excel.Worksheet)

                If .ProtectContents Then
                    .Unprotect(Password:="x")
                End If

                ersteZeile = CType(.Range(.Cells(1, 1), .Cells(1, 6 + bis - von)), Excel.Range)

                If awinSettings.englishLanguage Then
                    CType(.Cells(1, 1), Excel.Range).Value = "Business-Unit"
                    CType(.Cells(1, 2), Excel.Range).Value = "Project-Name"
                    CType(.Cells(1, 3), Excel.Range).Value = "Variant-Name"
                    CType(.Cells(1, 4), Excel.Range).Value = "Phase-Name"
                    CType(.Cells(1, 5), Excel.Range).Value = "Res./Cost-Name"
                    CType(.Cells(1, 6), Excel.Range).Value = "Sum"

                    If awinSettings.mePrzAuslastung Then
                        CType(.Cells(1, 7), Excel.Range).Value = "Percent."
                    Else
                        CType(.Cells(1, 7), Excel.Range).Value = "Avail."
                    End If
                Else
                    CType(.Cells(1, 1), Excel.Range).Value = "Business-Unit"
                    CType(.Cells(1, 2), Excel.Range).Value = "Projekt-Name"
                    CType(.Cells(1, 3), Excel.Range).Value = "Varianten-Name"
                    CType(.Cells(1, 4), Excel.Range).Value = "Phasen-Name"
                    CType(.Cells(1, 5), Excel.Range).Value = "Ress./Kostenart-Name"
                    CType(.Cells(1, 6), Excel.Range).Value = "Summe"

                    If awinSettings.mePrzAuslastung Then
                        CType(.Cells(1, 7), Excel.Range).Value = "Proz."
                    Else
                        CType(.Cells(1, 7), Excel.Range).Value = "Frei"
                    End If
                End If



                ' jetzt wird die Spalten-Nummer festgelegt, wo die Ressourcen/ Kosten später eingetragen werden
                ressCostColumn = 5
                ' jetzt wird die Zeile 1 geschrieben 
                Dim startMonat As Date = StartofCalendar.AddMonths(von - 1)

                ' jetzt wird der Mahle Range definiert ...
                mahleRange = CType(.Columns(startSpalteDaten - 1), Global.Microsoft.Office.Interop.Excel.Range).EntireColumn
                For tmpi As Integer = 0 To bis - von
                    mahleRange = appInstance.Union(mahleRange, CType(.Columns(startSpalteDaten + 2 * tmpi + 1), Global.Microsoft.Office.Interop.Excel.Range).EntireColumn)
                Next

                ' jetzt wird der Name hinzugefügt
                Dim tmpRange1 As Excel.Range = CType(.Cells(1, startSpalteDaten), Global.Microsoft.Office.Interop.Excel.Range)
                Dim tmpRange2 As Excel.Range = CType(.Cells(1, startSpalteDaten + 2 * (bis - von)), Global.Microsoft.Office.Interop.Excel.Range)
                Dim tmpRange3 As Excel.Range = CType(.Cells(1, 5), Global.Microsoft.Office.Interop.Excel.Range)

                Try
                    If Not IsNothing(CType(currentWB.Names.Item("MahleInfo"), Excel.Name)) Then
                        currentWB.Names.Item("MahleInfo").Delete()
                    End If
                Catch ex As Exception

                End Try

                Try
                    If Not IsNothing(CType(currentWB.Names.Item("StartData"), Excel.Name)) Then
                        currentWB.Names.Item("StartData").Delete()
                    End If
                Catch ex As Exception

                End Try

                Try
                    If Not IsNothing(CType(currentWB.Names.Item("EndData"), Excel.Name)) Then
                        currentWB.Names.Item("EndData").Delete()
                    End If
                Catch ex As Exception

                End Try

                Try
                    If Not IsNothing(CType(currentWB.Names.Item("RoleCost"), Excel.Name)) Then
                        currentWB.Names.Item("RoleCost").Delete()
                    End If
                Catch ex As Exception

                End Try

                currentWB.Names.Add(Name:="MahleInfo", RefersToR1C1:=mahleRange)
                currentWB.Names.Add(Name:="StartData", RefersToR1C1:=tmpRange1)
                currentWB.Names.Add(Name:="EndData", RefersToR1C1:=tmpRange2)
                currentWB.Names.Add(Name:="RoleCost", RefersToR1C1:=tmpRange3)

                ' jetzt werden die Überschriften des Datenbereichs geschrieben 
                For m As Integer = 0 To bis - von
                    With CType(.Cells(1, startSpalteDaten + 2 * m), Global.Microsoft.Office.Interop.Excel.Range)
                        .Value = startMonat.AddMonths(m)
                        .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                        .VerticalAlignment = Excel.XlVAlign.xlVAlignBottom
                        .NumberFormat = "[$-409]mmm yy;@"
                        .WrapText = False
                        .Orientation = 90
                        .ShrinkToFit = False
                        .AddIndent = False
                        .IndentLevel = 0
                        .ReadingOrder = Excel.Constants.xlContext
                    End With

                    With CType(.Cells(1, startSpalteDaten + 2 * m + 1), Global.Microsoft.Office.Interop.Excel.Range)
                        .Value = ""
                        .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                        .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                        .Orientation = 0
                        .ShrinkToFit = False
                        .AddIndent = False
                        .IndentLevel = 0
                        .ReadingOrder = Excel.Constants.xlContext
                    End With

                Next


            End With


            zeile = 2

            Dim schnittmenge() As Double
            Dim zeilenWerte() As Double
            Dim zeilensumme As Double
            Dim pStart As Integer, pEnde As Integer

            Dim editRange As Excel.Range


            ' zu Beginn werden die rollen-spezifischen Auslastungskennzahlen ermittelt, die sich über alle aktuell 
            ' betrachteten Projekte ergeben; 
            ' es werden sowohl die Gesamt-Auslastungs Werte im Zeitraum betrachtet als auch der einzelne monats-spezifische Wert   
            ' dazu wird ein Array angelegt mit der Dimension (anzahlRollen-1, bis-von+1) 
            Dim auslastungsArray(,) As Double

            ' tk, 18.5.17 damit kann die Gesamt- und Monatliche Auslastungs-Info ausgeblendet werden 
            If awinSettings.meExtendedColumnsView Then
                Try
                    auslastungsArray = visboZustaende.getUpDatedAuslastungsArray(Nothing, von, bis, awinSettings.mePrzAuslastung)
                    'auslastungsArray = ShowProjekte.getAuslastungsArray(von, bis)
                Catch ex As Exception
                    ReDim auslastungsArray(RoleDefinitions.Count - 1, bis - von + 1)
                End Try
            Else
                ReDim auslastungsArray(RoleDefinitions.Count - 1, bis - von + 1)
            End If


            Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)


            For Each projektName As String In todoListe

                Dim hproj As clsProjekt = Nothing
                If ShowProjekte.contains(projektName) Then
                    hproj = ShowProjekte.getProject(projektName)
                End If

                If Not IsNothing(hproj) Then

                    ' ist das Projekt geschützt ? 
                    ' wenn nein, dann temporär schützen 
                    Dim protectionText As String = ""
                    Dim wpItem As clsWriteProtectionItem
                    Dim isProtectedbyOthers As Boolean = Not tryToprotectProjectforMe(hproj.name, hproj.variantName)

                    If isProtectedbyOthers Then

                        ' nicht erfolgreich, weil durch anderen geschützt ... 
                        ' oder aber noch gar nicht in Datenbank: aber das ist noch nicht berücksichtigt  
                        wpItem = request.getWriteProtection(hproj.name, hproj.variantName)
                        writeProtections.upsert(wpItem)

                        protectionText = writeProtections.getProtectionText(calcProjektKey(hproj.name, hproj.variantName))

                    End If


                    pStart = getColumnOfDate(hproj.startDate)
                    pEnde = getColumnOfDate(hproj.endeDate)
                    Dim defaultEmptyValidation As String = validationStrings(rcValidation(anzahlRollen + 1)) ' alle Rollen und Kostenarten 

                    For p = 1 To hproj.CountPhases

                        Dim cphase As clsPhase = hproj.getPhase(p)
                        Dim phaseNameID As String = cphase.nameID
                        Dim phaseName As String = cphase.name
                        Dim chckNameID As String = calcHryElemKey(phaseName, False)

                        Dim indentlevel As Integer = hproj.hierarchy.getIndentLevel(phaseNameID)

                        If phaseWithinTimeFrame(pStart, cphase.relStart, cphase.relEnde, von, bis) Then
                            ' nur wenn die Phase überhaupt im betrachteten Zeitraum liegt, muss das berücksichtigt werden 

                            ' jetzt müssen die Zellen, die zur Phase gehören , geschrieben werden ...
                            Dim ixZeitraum As Integer
                            Dim ix As Integer, breite As Integer

                            Dim atLeastOne As Boolean = False

                            Call awinIntersectZeitraum(pStart + cphase.relStart - 1, pStart + cphase.relEnde - 1, ixZeitraum, ix, breite)


                            For r = 1 To cphase.countRoles

                                Dim role As clsRolle = cphase.getRole(r)
                                Dim roleName As String = role.name
                                Dim roleUID As Integer = RoleDefinitions.getRoledef(roleName).UID
                                Dim isSammelRolle As Boolean = RoleDefinitions.getRoledef(roleName).isCombinedRole
                                Dim xValues() As Double = role.Xwerte

                                If p = 1 And cphase.countRoles = 1 And r = 1 Then
                                    ' bestimme die defaultValidation für leere Zeilen : siehe if not atleastOne ...
                                    defaultEmptyValidation = validationStrings.Item(rcValidation(roleUID)) & ";" & _
                                        validationStrings.Item(rcValidation(0))
                                End If

                                schnittmenge = calcArrayIntersection(von, bis, pStart + cphase.relStart - 1, pStart + cphase.relEnde - 1, xValues)
                                zeilensumme = schnittmenge.Sum

                                ReDim zeilenWerte(2 * (bis - von + 1) - 1)

                                ' Schreiben der Projekt-Informationen 
                                With CType(currentWS, Excel.Worksheet)
                                    Dim cellComment As Excel.Comment

                                    ' Business Unit schreiben 
                                    CType(.Cells(zeile, 1), Excel.Range).Value = hproj.businessUnit

                                    ' Name schreiben
                                    CType(.Cells(zeile, 2), Excel.Range).Value = hproj.name
                                    ' wenn es protected ist, entsprechend markieren 
                                    If isProtectedbyOthers Then
                                        'CType(.Cells(zeile, 2), Excel.Range).Interior.Color = awinSettings.protectedByOtherColor
                                        CType(.Cells(zeile, 2), Excel.Range).Font.Color = awinSettings.protectedByOtherColor
                                        ' Kommentar einfügen 
                                        cellComment = CType(.Cells(zeile, 2), Excel.Range).Comment
                                        If Not IsNothing(cellComment) Then
                                            CType(.Cells(zeile, 2), Excel.Range).Comment.Delete()
                                        End If
                                        CType(.Cells(zeile, 2), Excel.Range).AddComment(Text:=protectionText)
                                        CType(.Cells(zeile, 2), Excel.Range).Comment.Visible = False
                                    End If

                                    CType(.Cells(zeile, 3), Excel.Range).Value = hproj.variantName
                                    CType(.Cells(zeile, 4), Excel.Range).Value = cphase.name

                                    ' Den Indent schreiben 
                                    CType(.Cells(zeile, 4), Excel.Range).IndentLevel = indentlevel

                                    cellComment = CType(.Cells(zeile, 4), Excel.Range).Comment
                                    If Not IsNothing(cellComment) Then
                                        CType(.Cells(zeile, 4), Excel.Range).Comment.Delete()
                                    End If
                                    If chckNameID = phaseNameID Then
                                        ' nichts weiter tun ... 
                                        ' denn dann kann die PhaseNameID aus der PhaseName konstruiert werden
                                        ' wenn es eine laufende Nummer 2, 3 etc ist, dann muss explizit die PhaseNameID in den Kommentarbereich geschreiben werden 
                                    Else
                                        CType(.Cells(zeile, 4), Excel.Range).AddComment(Text:=cphase.nameID)
                                        CType(.Cells(zeile, 4), Excel.Range).Comment.Visible = False
                                    End If

                                    With CType(.Cells(zeile, 5), Excel.Range)
                                        .Value = roleName
                                        If isProtectedbyOthers Then
                                        Else
                                            .Locked = False
                                            .Interior.Color = awinSettings.AmpelNichtBewertet
                                            Try

                                                If Not IsNothing(.Validation) Then
                                                    .Validation.Delete()
                                                End If

                                                ' jetzt wird die ValidationList aufgebaut 

                                                .Validation.Add(Type:=XlDVType.xlValidateList, AlertStyle:=XlDVAlertStyle.xlValidAlertStop, _
                                                                           Formula1:=validationStrings.Item(rcValidation(roleUID)))
                                            Catch ex As Exception

                                            End Try
                                        End If

                                    End With

                                    CType(.Cells(zeile, 6), Excel.Range).Value = zeilensumme.ToString("0")
                                    If awinSettings.allowSumEditing Then
                                        With CType(.Cells(zeile, 6), Excel.Range)

                                            If isProtectedbyOthers Then
                                            Else
                                                .Locked = False
                                                .Interior.Color = awinSettings.AmpelNichtBewertet
                                                Try
                                                    If Not IsNothing(.Validation) Then
                                                        .Validation.Delete()
                                                    End If
                                                    ' jetzt wird die ValidationList aufgebaut 
                                                    .Validation.Add(Type:=XlDVType.xlValidateDecimal, _
                                                                    AlertStyle:=XlDVAlertStyle.xlValidAlertStop, _
                                                                    Operator:=XlFormatConditionOperator.xlGreaterEqual, _
                                                                    Formula1:="0")
                                                Catch ex As Exception

                                                End Try
                                            End If

                                        End With
                                    End If


                                    If awinSettings.mePrzAuslastung Then
                                        CType(.Cells(zeile, 7), Excel.Range).Value = auslastungsArray(roleUID - 1, 0).ToString("0%")
                                    Else
                                        CType(.Cells(zeile, 7), Excel.Range).Value = auslastungsArray(roleUID - 1, 0).ToString("#,##0")
                                    End If

                                    editRange = CType(.Range(.Cells(zeile, startSpalteDaten), .Cells(zeile, startSpalteDaten + 2 * (bis - von + 1) - 1)), Excel.Range)
                                End With

                                ' zusammenmischen von Schnittmenge und Prozentual-Werte 
                                For mis As Integer = 0 To bis - von
                                    zeilenWerte(2 * mis) = schnittmenge(mis)
                                    ' in auslastungsarray(r, 0) steht die Gesamt-Auslastung
                                    If awinSettings.meExtendedColumnsView Then
                                        zeilenWerte(2 * mis + 1) = auslastungsArray(roleUID - 1, mis + 1)
                                    End If
                                Next

                                'editRange.Value = schnittmenge
                                editRange.Value = zeilenWerte
                                atLeastOne = True
                                ' die Zellen entsperren, die editiert werden dürfen ...

                                With CType(currentWS, Excel.Worksheet)
                                    For l = 0 To bis - von

                                        If l >= ixZeitraum And l <= ixZeitraum + breite - 1 Then

                                            With CType(.Cells(zeile, 2 * l + startSpalteDaten), Excel.Range)

                                                If isProtectedbyOthers Then
                                                Else
                                                    .Locked = False
                                                    Try
                                                        If Not IsNothing(.Validation) Then
                                                            .Validation.Delete()
                                                        End If
                                                    Catch ex As Exception

                                                    End Try

                                                    Try
                                                        .Validation.Add(Type:=XlDVType.xlValidateDecimal, _
                                                                    AlertStyle:=XlDVAlertStyle.xlValidAlertStop, _
                                                                    Operator:=XlFormatConditionOperator.xlGreaterEqual, _
                                                                    Formula1:="0")
                                                    Catch ex As Exception
                                                        
                                                    End Try
                                                End If


                                            End With
                                            ' erlaubter Eingabebereich grau markieren, aber nur wenn nicht protected 
                                            If isProtectedbyOthers Then
                                            Else
                                                CType(.Range(.Cells(zeile, 2 * l + startSpalteDaten), _
                                                         .Cells(zeile, 2 * l + 1 + startSpalteDaten)), Excel.Range).Interior.Color = awinSettings.AmpelNichtBewertet
                                            End If


                                            'CType(.Cells(zeile, 2 * l + startSpalteDaten), Excel.Range).Interior.Color = awinSettings.AmpelNichtBewertet
                                        Else
                                            CType(.Cells(zeile, 2 * l + startSpalteDaten), Excel.Range).Value = ""
                                            CType(.Cells(zeile, 2 * l + startSpalteDaten + 1), Excel.Range).Value = ""
                                        End If

                                    Next
                                End With


                                zeile = zeile + 1

                            Next r

                            For c = 1 To cphase.countCosts
                                Dim cost As clsKostenart = cphase.getCost(c)
                                Dim xValues() As Double = cost.Xwerte
                                Dim costName As String = cost.name
                                schnittmenge = calcArrayIntersection(von, bis, pStart + cphase.relStart - 1, pStart + cphase.relEnde - 1, xValues)
                                zeilensumme = schnittmenge.Sum

                                ReDim zeilenWerte(2 * (bis - von + 1) - 1)

                                ' Schreiben der Projekt-Informationen 
                                With CType(currentWS, Excel.Worksheet)
                                    Dim cellComment As Excel.Comment

                                    CType(.Cells(zeile, 1), Excel.Range).Value = hproj.businessUnit
                                    CType(.Cells(zeile, 2), Excel.Range).Value = hproj.name
                                    If isProtectedbyOthers Then
                                        'CType(.Cells(zeile, 2), Excel.Range).Interior.Color = awinSettings.protectedByOtherColor
                                        CType(.Cells(zeile, 2), Excel.Range).Font.Color = awinSettings.protectedByOtherColor
                                        ' Kommentar einfügen 
                                        cellComment = CType(.Cells(zeile, 2), Excel.Range).Comment
                                        If Not IsNothing(cellComment) Then
                                            CType(.Cells(zeile, 2), Excel.Range).Comment.Delete()
                                        End If
                                        CType(.Cells(zeile, 2), Excel.Range).AddComment(Text:=protectionText)
                                        CType(.Cells(zeile, 2), Excel.Range).Comment.Visible = False
                                    End If


                                    CType(.Cells(zeile, 3), Excel.Range).Value = hproj.variantName
                                    CType(.Cells(zeile, 4), Excel.Range).Value = cphase.name

                                    ' Den Indent schreiben 
                                    CType(.Cells(zeile, 4), Excel.Range).IndentLevel = indentlevel

                                    cellComment = CType(.Cells(zeile, 4), Excel.Range).Comment
                                    If Not IsNothing(cellComment) Then
                                        CType(.Cells(zeile, 4), Excel.Range).Comment.Delete()
                                    End If
                                    If chckNameID = phaseNameID Then
                                        ' nichts weiter tun ... 
                                        ' denn dann kann die PhaseNameID aus der PhaseName konstruiert werden
                                        ' wenn es eine laufende Nummer 2, 3 etc ist, dann muss explizit die PhaseNameID in den Kommentarbereich geschreiben werden 
                                    Else
                                        CType(.Cells(zeile, 4), Excel.Range).AddComment(Text:=cphase.nameID)
                                        CType(.Cells(zeile, 4), Excel.Range).Comment.Visible = False
                                    End If

                                    With CType(.Cells(zeile, 5), Excel.Range)
                                        .Value = costName
                                        If isProtectedbyOthers Then
                                        Else
                                            .Locked = False
                                            .Interior.Color = awinSettings.AmpelNichtBewertet
                                            Try
                                                If Not IsNothing(.Validation) Then
                                                    .Validation.Delete()
                                                End If
                                                ' jetzt wird die ValidationList aufgebaut 
                                                'Dim tmpVal As String = validationStrings.Item(rcValidation(0))
                                                .Validation.Add(Type:=XlDVType.xlValidateList, AlertStyle:=XlDVAlertStyle.xlValidAlertStop, _
                                                                               Formula1:=validationStrings.Item(rcValidation(0)))
                                            Catch ex As Exception

                                            End Try
                                        End If


                                    End With

                                    CType(.Cells(zeile, 6), Excel.Range).Value = zeilensumme.ToString("0")
                                    If awinSettings.allowSumEditing Then

                                        With CType(.Cells(zeile, 6), Excel.Range)
                                            If isProtectedbyOthers Then
                                            Else
                                                .Locked = False
                                                .Interior.Color = awinSettings.AmpelNichtBewertet
                                                Try
                                                    If Not IsNothing(.Validation) Then
                                                        .Validation.Delete()
                                                    End If
                                                    ' jetzt wird die ValidationList aufgebaut 
                                                    .Validation.Add(Type:=XlDVType.xlValidateDecimal, _
                                                                    AlertStyle:=XlDVAlertStyle.xlValidAlertStop, _
                                                                    Operator:=XlFormatConditionOperator.xlGreaterEqual, _
                                                                    Formula1:="0")
                                                Catch ex As Exception

                                                End Try
                                            End If


                                        End With
                                    End If

                                    editRange = CType(.Range(.Cells(zeile, startSpalteDaten), .Cells(zeile, startSpalteDaten + 2 * (bis - von + 1) - 1)), Excel.Range)
                                End With

                                ' zusammenmischen von Schnittmenge und Prozentual-Werte 
                                For mis As Integer = 0 To bis - von
                                    zeilenWerte(2 * mis) = schnittmenge(mis)
                                    ' in auslastungsarray(r, 0) steht die Gesamt-Auslastung, spielt aber kein Kostenarten keine Rolle 
                                    ' tk, 18.5 wird ja schon durch Redim erledigt ...
                                    'zeilenWerte(2 * mis + 1) = 0
                                Next

                                'editRange.Value = schnittmenge
                                editRange.Value = zeilenWerte
                                atLeastOne = True

                                ' die Zellen entsperren, die editiert werden dürfen ...

                                With CType(currentWS, Excel.Worksheet)

                                    For l = 0 To bis - von

                                        If l >= ixZeitraum And l <= ixZeitraum + breite - 1 Then

                                            With CType(.Cells(zeile, 2 * l + startSpalteDaten), Excel.Range)
                                                If isProtectedbyOthers Then
                                                Else
                                                    .Locked = False
                                                    Try
                                                        If Not IsNothing(.Validation) Then
                                                            .Validation.Delete()
                                                        End If
                                                        .Validation.Add(Type:=XlDVType.xlValidateDecimal, _
                                                                    AlertStyle:=XlDVAlertStyle.xlValidAlertStop, _
                                                                    Operator:=XlFormatConditionOperator.xlGreaterEqual, _
                                                                    Formula1:="0")
                                                    Catch ex As Exception

                                                    End Try

                                                End If

                                            End With

                                            CType(.Cells(zeile, 2 * l + 1 + startSpalteDaten), Excel.Range).Value = ""

                                            ' nur die Zelle grau markieren , um in der Logik konsistent zu sein 
                                            If isProtectedbyOthers Then
                                            Else
                                                CType(.Range(.Cells(zeile, 2 * l + startSpalteDaten), _
                                                         .Cells(zeile, 2 * l + 1 + startSpalteDaten)), Excel.Range).Interior.Color = awinSettings.AmpelNichtBewertet
                                            End If

                                        Else
                                            CType(.Cells(zeile, 2 * l + startSpalteDaten), Excel.Range).Value = ""
                                            CType(.Cells(zeile, 2 * l + 1 + startSpalteDaten), Excel.Range).Value = ""
                                        End If

                                    Next

                                End With

                                zeile = zeile + 1

                            Next c

                            If Not atLeastOne Then
                                ' in diesem Fall sollte eine leere Projekt-Phasen-Information geschrieben werden, quasi ein Platzhalter
                                ' in diesem Platzhalter kann dann später die Ressourcen Information aufgenommen werden  
                                ' Schreiben der Projekt-Informationen 
                                Dim currentValidation As String = rcValidation(anzahlRollen + 1)

                                ' bestimmen, ob rootPhase nur eine Rolle hat, dann soll die Validation aus der Validation dieser Rolle plus allen Kostenarten gebildet werden ... 
                                Try
                                    If cphase.nameID <> rootPhaseName Then

                                    End If
                                Catch ex As Exception

                                End Try

                                With CType(currentWS, Excel.Worksheet)
                                    Dim cellComment As Excel.Comment

                                    CType(.Cells(zeile, 1), Excel.Range).Value = hproj.businessUnit
                                    CType(.Cells(zeile, 2), Excel.Range).Value = hproj.name
                                    If isProtectedbyOthers Then
                                        'CType(.Cells(zeile, 2), Excel.Range).Interior.Color = awinSettings.protectedByOtherColor
                                        CType(.Cells(zeile, 2), Excel.Range).Font.Color = awinSettings.protectedByOtherColor
                                        ' Kommentar einfügen 
                                        cellComment = CType(.Cells(zeile, 2), Excel.Range).Comment
                                        If Not IsNothing(cellComment) Then
                                            CType(.Cells(zeile, 2), Excel.Range).Comment.Delete()
                                        End If
                                        CType(.Cells(zeile, 2), Excel.Range).AddComment(Text:=protectionText)
                                        CType(.Cells(zeile, 2), Excel.Range).Comment.Visible = False
                                    End If

                                    CType(.Cells(zeile, 3), Excel.Range).Value = hproj.variantName
                                    CType(.Cells(zeile, 4), Excel.Range).Value = cphase.name

                                    ' Den Indent schreiben 
                                    CType(.Cells(zeile, 4), Excel.Range).IndentLevel = indentlevel

                                    cellComment = CType(.Cells(zeile, 4), Excel.Range).Comment

                                    If Not IsNothing(cellComment) Then
                                        CType(.Cells(zeile, 4), Excel.Range).Comment.Delete()
                                    End If
                                    If chckNameID = phaseNameID Then
                                        ' nichts weiter tun ... 
                                        ' denn dann kann die PhaseNameID aus der PhaseName konstruiert werden
                                        ' wenn es eine laufende Nummer 2, 3 etc ist, dann muss explizit die PhaseNameID in den Kommentarbereich geschreiben werden 
                                    Else
                                        CType(.Cells(zeile, 4), Excel.Range).AddComment(Text:=cphase.nameID)
                                        CType(.Cells(zeile, 4), Excel.Range).Comment.Visible = False
                                    End If

                                    With CType(.Cells(zeile, 5), Excel.Range)
                                        .Value = ""
                                        If isProtectedbyOthers Then
                                        Else
                                            .Locked = False
                                            .Interior.Color = awinSettings.AmpelNichtBewertet
                                            Try
                                                If Not IsNothing(.Validation) Then
                                                    .Validation.Delete()
                                                End If
                                                ' jetzt wird die ValidationList aufgebaut 
                                                .Validation.Add(Type:=XlDVType.xlValidateList, AlertStyle:=XlDVAlertStyle.xlValidAlertStop, _
                                                                               Formula1:=defaultEmptyValidation)
                                            Catch ex As Exception
                                                Dim a As Integer = 0
                                            End Try
                                        End If


                                    End With

                                    If awinSettings.allowSumEditing Then
                                        With CType(.Cells(zeile, 6), Excel.Range)
                                            .Value = ""
                                            If isProtectedbyOthers Then
                                            Else
                                                .Locked = False
                                                .Interior.Color = awinSettings.AmpelNichtBewertet
                                                Try
                                                    If Not IsNothing(.Validation) Then
                                                        .Validation.Delete()
                                                    End If
                                                    ' jetzt wird die ValidationList aufgebaut 
                                                    .Validation.Add(Type:=XlDVType.xlValidateDecimal, _
                                                                    AlertStyle:=XlDVAlertStyle.xlValidAlertStop, _
                                                                    Operator:=XlFormatConditionOperator.xlGreaterEqual, _
                                                                    Formula1:="0")
                                                Catch ex As Exception

                                                End Try
                                            End If


                                        End With

                                    Else
                                        CType(.Cells(zeile, 6), Excel.Range).Value = ""
                                    End If


                                    CType(.Cells(zeile, 7), Excel.Range).Value = ""
                                    editRange = CType(.Range(.Cells(zeile, startSpalteDaten), .Cells(zeile, startSpalteDaten + 2 * (bis - von))), Excel.Range)
                                End With

                                ' die Zellen farblich markieren, die editiert werden können ...
                                With CType(currentWS, Excel.Worksheet)

                                    For l = 0 To bis - von

                                        If l >= ixZeitraum And l <= ixZeitraum + breite - 1 Then

                                            With CType(.Cells(zeile, 2 * l + startSpalteDaten), Excel.Range)
                                                If isProtectedbyOthers Then
                                                Else
                                                    .Locked = False

                                                    If Not IsNothing(.Validation) Then
                                                        .Validation.Delete()
                                                    End If

                                                    Try
                                                        .Validation.Add(Type:=XlDVType.xlValidateDecimal, _
                                                                    AlertStyle:=XlDVAlertStyle.xlValidAlertStop, _
                                                                    Operator:=XlFormatConditionOperator.xlGreaterEqual, _
                                                                    Formula1:="0")
                                                    Catch ex As Exception

                                                    End Try

                                                End If

                                            End With

                                            CType(.Cells(zeile, 2 * l + 1 + startSpalteDaten), Excel.Range).Value = ""

                                            If isProtectedbyOthers Then
                                            Else
                                                CType(.Range(.Cells(zeile, 2 * l + startSpalteDaten), _
                                                         .Cells(zeile, 2 * l + 1 + startSpalteDaten)), Excel.Range).Interior.Color = awinSettings.AmpelNichtBewertet
                                            End If

                                        Else
                                            CType(.Cells(zeile, 2 * l + startSpalteDaten), Excel.Range).Value = ""
                                            CType(.Cells(zeile, 2 * l + 1 + startSpalteDaten), Excel.Range).Value = ""
                                        End If

                                    Next

                                End With

                                zeile = zeile + 1

                            End If

                        End If



                    Next p


                End If



            Next


            ' tk 7.12.16 kommt immer auf Fehler, weil nur 1 Zeile und eine Auswahl von Spalten .... 
            '' jetzt die erste Zeile so groß wie nötig machen 
            'Try
            '    ersteZeile.AutoFit()
            'Catch ex As Exception

            'End Try

            ' jetzt die Größe der Spalten für BU, pName, vName, Phasen-Name, RC-Name anpassen 
            Dim infoBlock As Excel.Range
            With CType(currentWS, Excel.Worksheet)
                infoBlock = CType(.Range(.Columns(1), .Columns(startSpalteDaten - 3)), Excel.Range)
                infoBlock.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
                infoBlock.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                infoBlock.AutoFit()
            End With

            ' Summe und Frei / Proz. 
            With CType(currentWS, Excel.Worksheet)
                infoBlock = CType(.Range(.Columns(startSpalteDaten - 2), .Columns(startSpalteDaten - 1)), Excel.Range)
                infoBlock.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
                infoBlock.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                infoBlock.AutoFit()
            End With

            Dim tmpRange As Excel.Range
            With CType(currentWS, Excel.Worksheet)

                Dim isPrz As Boolean = False
                For mis As Integer = 0 To 2 * (bis - von + 1) - 1
                    tmpRange = CType(.Range(.Cells(2, startSpalteDaten + mis), .Cells(zeile, startSpalteDaten + mis)), Excel.Range)
                    If isPrz Then
                        tmpRange.Columns.ColumnWidth = 4
                        tmpRange.Font.Size = 8
                        If awinSettings.mePrzAuslastung Then
                            tmpRange.NumberFormat = "0%"
                        Else
                            tmpRange.NumberFormat = "0"
                        End If

                        tmpRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
                    Else
                        tmpRange.Columns.ColumnWidth = 5
                        tmpRange.Font.Size = 10
                        tmpRange.NumberFormat = "0"
                        tmpRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
                    End If
                    isPrz = Not isPrz
                Next

            End With

            ' jetzt wird ggf der MahleRange ausgeblendet ... 
            If Not awinSettings.meExtendedColumnsView Then
                Try
                    'mahleRange.Locked = False
                    mahleRange.EntireColumn.Hidden = True
                Catch ex As Exception

                End Try

            End If

            appInstance.EnableEvents = True

        Catch ex As Exception
            Dim a As Integer = 0
        End Try

        


    End Sub

    ''' <summary>
    ''' versucht das Projekt für mich zu schützen 
    ''' gibt false zurück , wenn das Projekt durch andere geschützt ist 
    ''' </summary>
    ''' <param name="pname"></param>
    ''' <param name="vName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function tryToprotectProjectforMe(ByVal pName As String, ByVal vName As String) As Boolean

        Dim wpItem As clsWriteProtectionItem
        Dim isProtectedbyOthers As Boolean
        Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)

        If request.projectNameAlreadyExists(pName, vName, Date.Now) Then

            ' es existiert in der Datenbank ...
            If request.checkChgPermission(pName, vName, dbUsername) Then

                isProtectedbyOthers = False
                ' jetzt prüfen, ob es Null ist, von mir permanent/nicht permanent geschützt wurde .. 
                wpItem = request.getWriteProtection(pName, vName)

                Dim notYetDone As Boolean = False

                If IsNothing(wpItem) Then
                    ' wpitem kann NULL sein
                    notYetDone = True

                ElseIf wpItem.permanent Then
                    notYetDone = False
                    ' meinen permanenten Schutz einbauen 
                    writeProtections.upsert(wpItem)

                Else
                    notYetDone = True
                End If

                If notYetDone Then
                    wpItem = New clsWriteProtectionItem(calcProjektKey(pName, vName), _
                                                              ptWriteProtectionType.project, _
                                                              dbUsername, _
                                                              False, _
                                                              True)

                    If request.setWriteProtection(wpItem) Then
                        ' erfolgreich ...
                        writeProtections.upsert(wpItem)
                    Else
                        ' in diesem Fall wurde es in der Zwischenzeit von jdn anders geschützt  
                        isProtectedbyOthers = True
                    End If

                End If

            Else
                isProtectedbyOthers = True
            End If
        Else
            ' das Projekt existiert bisher nur in der Session des Nutzers 
            isProtectedbyOthers = False
        End If


        tryToprotectProjectforMe = Not isProtectedbyOthers

    End Function

    ''' <summary>
    ''' aktualisiert in Tabelle2 die Auslastungs-Values 
    ''' Voraussetzung: der Auslastungs-Array in visbozustaende ist aktualisiert 
    ''' </summary>
    ''' <param name="roleNames">eine sortierte Collection mit den Namen der Rollen, die aktualisiert werden sollen
    ''' Kostenarten brauchen nicht aktualisiert zu werden, da die keine Kapa / Grenze kennen </param>
    ''' <remarks></remarks>
    Public Sub updateMassEditAuslastungsValues(ByVal von As Integer, ByVal bis As Integer, _
                                               Optional ByVal roleNames As Collection = Nothing)
        Dim treatAllRoles As Boolean
        Dim roleID As Integer

        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False

        If CType(appInstance.ActiveSheet, Excel.Worksheet).Name = arrWsNames(ptTables.meRC) Then
            ' nur dann befindet sich das Programm im MassEdit Sheet 

            'Dim meWS As Excel.Worksheet = CType(appInstance.ActiveSheet, Excel.Worksheet)
            Dim meWS As Excel.Worksheet = CType(CType(appInstance.Workbooks(myProjektTafel), Excel.Workbook) _
            .Worksheets(arrWsNames(ptTables.meRC)), Excel.Worksheet)

            If IsNothing(roleNames) Then
                treatAllRoles = True
            Else
                treatAllRoles = False
            End If

            Dim columnStartData As Integer = visboZustaende.meColSD
            Dim columnEndData As Integer = visboZustaende.meColED
            Dim columnRC As Integer = visboZustaende.meColRC

            Dim auslastungsArray(,) As Double = visboZustaende.getUpDatedAuslastungsArray(roleNames, von, bis, awinSettings.mePrzAuslastung)

            Dim tstZeilenanzahl = meWS.UsedRange.Rows.Count

            ' jetzt in der Überschrift den Text anpassen 
            If awinSettings.mePrzAuslastung Then
                CType(meWS.Cells(1, 7), Excel.Range).Value = "Proz."
            Else
                CType(meWS.Cells(1, 7), Excel.Range).Value = "Frei"
            End If

            ' jetzt muss einfach jede Zeile im Mass-Edit Sheet durchgegangen werden 
            For zeile As Integer = 2 To visboZustaende.meMaxZeile
                Dim curRoleName As String = CStr(meWS.Cells(zeile, columnRC).value)
                If Not IsNothing(curRoleName) Then
                    If curRoleName.Trim.Length > 0 Then
                        If RoleDefinitions.containsName(curRoleName) Then
                            Dim updateNecessary As Boolean = False
                            If treatAllRoles Then
                                updateNecessary = True
                            ElseIf roleNames.Contains(curRoleName) Then
                                updateNecessary = True
                            End If
                            If updateNecessary Then
                                ' nur in diesem Fall muss was gemacht werden ... 
                                roleID = RoleDefinitions.getRoledef(curRoleName).UID

                                For mis As Integer = 0 To bis - von + 1
                                    Dim tmpCol As Integer = columnStartData - 1 + 2 * mis
                                    With meWS.Cells(zeile, tmpCol)
                                        If ((tmpCol = columnStartData - 1) Or _
                                            (meWS.Cells(zeile, tmpCol - 1).locked = False)) Then
                                            .value = auslastungsArray(roleID - 1, mis)
                                            If awinSettings.mePrzAuslastung Then
                                                .NumberFormat = "0%"
                                            Else
                                                .NumberFormat = "0"
                                            End If
                                        End If
                                    End With
                                Next

                            End If
                        End If
                    End If
                End If
            Next

        Else
            Call MsgBox("Mass-Edit Sheet nicht aktiv ...")
        End If

        appInstance.EnableEvents = formerEE
        'appInstance.EnableEvents = True

    End Sub


    ''' <summary>
    ''' aktualisiert die Summen-Werte im Massen-Edit Sheet der Ressourcen-/Kostenzuordnungen  
    ''' </summary>
    ''' <param name="pname"></param>
    ''' <param name="von"></param>
    ''' <param name="bis"></param>
    ''' <param name="roleCostNames"></param>
    ''' <remarks></remarks>
    Public Sub updateMassEditSummenValues(ByVal pname As String, ByVal phaseNameID As String, _
                                              ByVal von As Integer, ByVal bis As Integer, _
                                              ByVal roleCostNames As Collection)


        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False

        If CType(appInstance.ActiveSheet, Excel.Worksheet).Name = arrWsNames(ptTables.meRC) Then
            ' nur dann befindet sich das Programm im MassEdit Sheet 

            'Dim meWS As Excel.Worksheet = CType(appInstance.ActiveSheet, Excel.Worksheet)
            Dim meWS As Excel.Worksheet = CType(CType(appInstance.Workbooks(myProjektTafel), Excel.Workbook) _
            .Worksheets(arrWsNames(ptTables.meRC)), Excel.Worksheet)

            If IsNothing(roleCostNames) Then
                ' nichts tun 
            ElseIf roleCostNames.Count = 0 Then
                ' nichts tun 
            Else
                ' Update Lauf der Summen 
                Dim columnSummen As Integer = visboZustaende.meColRC + 1
                Dim columnRC As Integer = visboZustaende.meColRC

                ' jetzt muss einfach jede Zeile im Mass-Edit Sheet durchgegangen werden 
                For zeile As Integer = 2 To visboZustaende.meMaxZeile

                    Dim curpName As String = CStr(meWS.Cells(zeile, 2).value)
                    Dim curphaseName As String = CStr(meWS.Cells(zeile, 4).value)
                    Dim curphaseNameID As String = calcHryElemKey(curphaseName, False)
                    Dim curComment As Excel.Comment = CType(meWS.Cells(zeile, 4), Excel.Range).Comment
                    If Not IsNothing(curComment) Then
                        curphaseNameID = curComment.Text
                    End If

                    ' es soll auf jeden Fall auch die Rootphase geupdated werden ..., da ja die evtl auch als secondbest geändert wurde ...
                    If curpName = pname And ((curphaseNameID = phaseNameID) Or (curphaseNameID = rootPhaseName)) Then

                        Dim curRCName As String = CStr(meWS.Cells(zeile, columnRC).value)

                        If Not IsNothing(curRCName) Then
                            If curRCName.Trim.Length > 0 Then
                                If roleCostNames.Contains(curRCName) Then
                                    Dim tmpSum As Double = 0.0
                                    ' jetzt muss die Summe aktualisiert werden 
                                    Dim hproj As clsProjekt = ShowProjekte.getProject(pname)
                                    If Not IsNothing(hproj) Then
                                        Dim cphase As clsPhase = hproj.getPhaseByID(curphaseNameID)

                                        If Not IsNothing(cphase) Then

                                            Dim xWerte() As Double
                                            Dim ixZeitraum As Integer
                                            Dim ix As Integer
                                            Dim anzLoops As Integer

                                            ' diese MEthode definiert, wo der Zeitraum sich mit den Werte überlappt ... 
                                            ' Anzloops sind die Anzahl Überlappungen 
                                            Call awinIntersectZeitraum(getColumnOfDate(cphase.getStartDate), getColumnOfDate(cphase.getEndDate), _
                                                               ixZeitraum, ix, anzLoops)

                                            If RoleDefinitions.containsName(curRCName) Then

                                                Dim tmpRole As clsRolle = cphase.getRole(curRCName)

                                                If Not IsNothing(tmpRole) Then
                                                    xWerte = tmpRole.Xwerte

                                                    ' jetzt werden die Werte summiert ...
                                                    Try
                                                        For al As Integer = 1 To anzLoops
                                                            tmpSum = tmpSum + xWerte(ix + al - 1)
                                                        Next
                                                    Catch ex As Exception
                                                        Call MsgBox("Fehler bei Summenbildung ...")
                                                        tmpSum = 0
                                                    End Try


                                                Else
                                                    ' Summe löschen
                                                End If

                                            ElseIf CostDefinitions.containsName(curRCName) Then

                                                Dim tmpCost As clsKostenart = cphase.getCost(curRCName)

                                                If Not IsNothing(tmpCost) Then
                                                    xWerte = tmpCost.Xwerte

                                                    ' jetzt werden die Werte summiert ...
                                                    Try
                                                        For al As Integer = 1 To anzLoops
                                                            tmpSum = tmpSum + xWerte(ix + al - 1)
                                                        Next
                                                    Catch ex As Exception
                                                        Call MsgBox("Fehler bei Summenbildung ...")
                                                        tmpSum = 0
                                                    End Try

                                                Else
                                                    ' Summe löschen
                                                End If
                                            Else
                                                ' Summe löschen 
                                            End If

                                        Else
                                            ' Summe löschen  
                                        End If
                                    Else
                                        ' Summe löschen 
                                    End If

                                    ' jetzt den Wert in die Zelle schreiben
                                    If tmpSum > 0 Then
                                        CType(meWS.Cells(zeile, columnSummen), Excel.Range).Value = tmpSum.ToString("#,##0")
                                    Else
                                        CType(meWS.Cells(zeile, columnSummen), Excel.Range).Value = ""
                                    End If

                                End If
                            End If
                        End If

                    End If

                Next

            End If


        Else
            Call MsgBox("Mass-Edit Sheet nicht aktiv ...")
        End If

        appInstance.EnableEvents = formerEE


    End Sub

    ''' <summary>
    ''' aktualisiert für das angegebene Projekt die Validation Strings aller leeren / empty RoleCost Felder gemäß dem übergebenen 
    ''' dient dazu, um die Validation an der rootPhaseName Setzung zu orientieren 
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <param name="validationString"></param>
    ''' <remarks></remarks>
    Public Sub updateEmptyRcCellValidations(ByVal pName As String, ByVal validationString As String)

        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False

        If CType(appInstance.ActiveSheet, Excel.Worksheet).Name = arrWsNames(ptTables.meRC) Then
            ' nur dann befindet sich das Programm im MassEdit Sheet 

            Dim meWS As Excel.Worksheet = CType(CType(appInstance.Workbooks(myProjektTafel), Excel.Workbook) _
            .Worksheets(arrWsNames(ptTables.meRC)), Excel.Worksheet)

            If IsNothing(pName) Or IsNothing(validationString) Then
                ' nichts tun 
            ElseIf pName.Trim.Length = 0 Or validationString.Trim.Length = 0 Then
                ' nichts tun 
            Else
                ' Update der Validations der leeren RoleCost Zuordnungen  
                Dim columnRC As Integer = visboZustaende.meColRC

                ' jetzt muss einfach jede Zeile im Mass-Edit Sheet durchgegangen werden 
                For zeile As Integer = 2 To visboZustaende.meMaxZeile

                    Dim curpName As String = CStr(meWS.Cells(zeile, 2).value)
                    Dim curphaseName As String = CStr(meWS.Cells(zeile, 4).value)
                    Dim needsUpdate As Boolean = False

                    ' es soll auf jeden Fall auch die Rootphase geupdated werden ..., da ja die evtl auch als secondbest geändert wurde ...
                    If curpName = pName And curphaseName <> "." Then

                        Dim curRCName As String = CStr(meWS.Cells(zeile, columnRC).value)

                        If IsNothing(curRCName) Then
                            needsUpdate = True
                        ElseIf curRCName.Trim.Length = 0 Then
                            needsUpdate = True
                        End If

                    End If

                    If needsUpdate Then
                        Try
                            With CType(meWS.Cells(zeile, columnRC), Excel.Range)
                                If Not IsNothing(.Validation) Then
                                    .Validation.Delete()
                                End If
                                ' jetzt wird die ValidationList aufgebaut 

                                .Validation.Add(Type:=XlDVType.xlValidateList, AlertStyle:=XlDVAlertStyle.xlValidAlertStop, _
                                                               Formula1:=validationString)

                            End With
                        Catch ex As Exception

                        End Try

                    End If

                Next

            End If


        Else
            Call MsgBox("Mass-Edit Sheet nicht aktiv ...")
        End If

        appInstance.EnableEvents = formerEE

    End Sub

    ''' <summary>
    ''' prüft ob in dem aktiven Massen-Edit Sheet die übergebene Kombination nocheinmal vorkommt ... 
    ''' wenn nein: Rückgabe true
    ''' wenn ja: Rückgabe false
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <param name="phaseNameID"></param>
    ''' <param name="rcName"></param>
    ''' <param name="zeile"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function noDuplicatesInSheet(ByVal pName As String, ByVal phaseNameID As String, ByVal rcName As String, _
                                             ByVal zeile As Integer) As Boolean
        Dim found As Boolean = False
        Dim curZeile As Integer = 2

        Dim chckName As String
        Dim chckPhNameID As String
        Dim chckRCName As String

        Dim meWS As Excel.Worksheet = CType(CType(appInstance.Workbooks(myProjektTafel), Excel.Workbook) _
            .Worksheets(arrWsNames(ptTables.meRC)), Excel.Worksheet)

        With meWS
            chckName = CStr(meWS.Cells(curZeile, 2).value)

            Dim phaseName As String = CStr(meWS.Cells(curZeile, 4).value)
            chckPhNameID = calcHryElemKey(phaseName, False)
            Dim curComment As Excel.Comment = CType(meWS.Cells(curZeile, 4), Excel.Range).Comment
            If Not IsNothing(curComment) Then
                chckPhNameID = curComment.Text
            End If

            chckRCName = CStr(meWS.Cells(curZeile, 5).value)

        End With
        ' aus der Funktionalität zeile löschen wird rcName auch mit Nothing aufgerufen ... 
        Do While Not found And curZeile <= visboZustaende.meMaxZeile


            If chckName = pName And _
                phaseNameID = chckPhNameID And _
                zeile <> curZeile Then

                If IsNothing(rcName) Then
                    found = True
                ElseIf rcName = chckRCName Then
                    found = True
                End If

            End If

            If Not found Then

                curZeile = curZeile + 1

                With meWS
                    chckName = CStr(meWS.Cells(curZeile, 2).value)

                    Dim phaseName As String = CStr(meWS.Cells(curZeile, 4).value)
                    chckPhNameID = calcHryElemKey(phaseName, False)
                    Dim curComment As Excel.Comment = CType(meWS.Cells(curZeile, 4), Excel.Range).Comment
                    If Not IsNothing(curComment) Then
                        chckPhNameID = curComment.Text
                    End If

                    chckRCName = CStr(meWS.Cells(curZeile, 5).value)

                End With

            End If

        Loop

        noDuplicatesInSheet = Not found

    End Function

    ''' <summary>
    ''' gibt eine Zeile zurück, die zu dem angegebenen Projekt, der Phase und dem rcName eine Sammelrolle zurückgibt
    ''' Wenn mehrere mögliche Sammelrollen existieren, dann wird die erste auftretende zurückgegeben
    ''' 0, wenn in diesem Projekt zu dieser Rolle keine Sammelrolle definiert ist  
    ''' ausserdem wird erst in der gleichen Phase gesucht, oder aber in der RootPhase
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <param name="phaseNameID"></param>
    ''' <param name="rcName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function findeSammelRollenZeile(ByVal pName As String, ByVal phaseNameID As String, ByVal rcName As String) As Integer
        Dim found As Boolean = False
        Dim curZeile As Integer = 2

        Dim chckName As String
        Dim chckPhNameID As String
        Dim chckRCName As String
        Dim bestName As String = ""
        Dim secondBestzeile As Integer = 0

        Dim tmpRole As clsRollenDefinition = RoleDefinitions.getRoledef(rcName)

        If Not IsNothing(tmpRole) Then

            Dim istSammelRolle As Boolean = tmpRole.isCombinedRole

            If istSammelRolle Then
                curZeile = 0
            Else
                ' nur für echte Rollen durchführen ...
                Dim potentialParentRoles As Collection = RoleDefinitions.getSummaryRoles(rcName)

                If potentialParentRoles.Count = 0 Then
                    curZeile = 0

                Else
                    '
                    ' auf die Suche gehen ... 
                    Dim meWS As Excel.Worksheet = CType(CType(appInstance.Workbooks(myProjektTafel), Excel.Workbook) _
                    .Worksheets(arrWsNames(ptTables.meRC)), Excel.Worksheet)

                    With meWS
                        chckName = CStr(meWS.Cells(curZeile, 2).value)
                        If IsNothing(chckName) Then
                            chckName = ""
                        End If

                        Dim phaseName As String = CStr(meWS.Cells(curZeile, 4).value)
                        If IsNothing(phaseName) Then
                            phaseName = ""
                        End If

                        chckPhNameID = calcHryElemKey(phaseName, False)
                        Dim curComment As Excel.Comment = CType(meWS.Cells(curZeile, 4), Excel.Range).Comment
                        If Not IsNothing(curComment) Then
                            chckPhNameID = curComment.Text
                        End If

                        chckRCName = CStr(meWS.Cells(curZeile, 5).value)
                        If IsNothing(chckRCName) Then
                            chckRCName = ""
                        End If

                    End With
                    ' 
                    ' jetzt wird erst geprüft, ob es eine Sammelrolle in der gleichen Phase gibt 
                    ' dann wird geprüft , ob es eine Sammelrolle in der rootphase gibt 



                    Do While Not found And curZeile <= visboZustaende.meMaxZeile


                        If ((chckName = pName) And _
                            ((phaseNameID = chckPhNameID) Or (rootPhaseName = chckPhNameID))) Then

                            If potentialParentRoles.Contains(chckRCName) Then
                                ' nimm jetzt einfach mal den ersten, der auftritt  ... 
                                If phaseNameID = chckPhNameID Then
                                    found = True
                                Else
                                    secondBestzeile = curZeile
                                    ' noch weitersuchen, ob nicht noch das found-Kriterium greift ... 
                                End If

                            End If

                        End If

                        If Not found Then

                            curZeile = curZeile + 1

                            With meWS
                                chckName = CStr(meWS.Cells(curZeile, 2).value)
                                If IsNothing(chckName) Then
                                    chckName = ""
                                End If

                                Dim phaseName As String = CStr(meWS.Cells(curZeile, 4).value)
                                If IsNothing(phaseName) Then
                                    phaseName = ""
                                End If

                                chckPhNameID = calcHryElemKey(phaseName, False)
                                Dim curComment As Excel.Comment = CType(meWS.Cells(curZeile, 4), Excel.Range).Comment
                                If Not IsNothing(curComment) Then
                                    chckPhNameID = curComment.Text
                                End If

                                chckRCName = CStr(meWS.Cells(curZeile, 5).value)
                                If IsNothing(chckRCName) Then
                                    chckRCName = ""
                                End If

                            End With

                        End If

                    Loop

                End If

            End If

        End If




        If found Then
            findeSammelRollenZeile = curZeile
        ElseIf secondBestzeile > 0 Then
            findeSammelRollenZeile = secondBestzeile
        Else
            findeSammelRollenZeile = 0
        End If

    End Function
    ''' <summary>
    ''' liest, falls vorhanden aus ProjectboardConfig.xml die Settings
    ''' wenn nicht vorhanden, gibt false zurück 
    ''' </summary>
    ''' <param name="path"></param>
    ''' <returns>ob erfolgreich oder nicht </returns>
    ''' <remarks></remarks>
    Public Function readawinSettings(ByVal path As String) As Boolean


        Dim cfgs As New configuration
        Dim cfgFile As String = path & "\ProjectboardConfig.xml"

        Dim erg As Boolean = My.Computer.FileSystem.FileExists(cfgFile)

        Try

            cfgs = XMLImportConfig(cfgFile)

            If Not IsNothing(cfgs) Then

                Dim anzahlSettings As Integer = cfgs.applicationSettings.ExcelWorkbook1MySettings.Length

                For i = 0 To anzahlSettings - 1

                    Select Case cfgs.applicationSettings.ExcelWorkbook1MySettings(i).name
                        Case "mongoDBURL"
                            awinSettings.databaseURL = cfgs.applicationSettings.ExcelWorkbook1MySettings(i).value
                        Case "mongoDBname"
                            awinSettings.databaseName = cfgs.applicationSettings.ExcelWorkbook1MySettings(i).value
                        Case "globalPath"
                            awinSettings.globalPath = cfgs.applicationSettings.ExcelWorkbook1MySettings(i).value
                        Case "awinPath"
                            awinSettings.awinPath = cfgs.applicationSettings.ExcelWorkbook1MySettings(i).value
                        Case "VISBOTaskClass"
                            awinSettings.visboTaskClass = cfgs.applicationSettings.ExcelWorkbook1MySettings(i).value
                        Case "VISBOAbbreviation"
                            awinSettings.visboAbbreviation = cfgs.applicationSettings.ExcelWorkbook1MySettings(i).value
                        Case "VISBOAmpel"
                            awinSettings.visboAmpel = cfgs.applicationSettings.ExcelWorkbook1MySettings(i).value
                        Case "VISBODebug"
                            awinSettings.visboDebug = CType(cfgs.applicationSettings.ExcelWorkbook1MySettings(i).value, Boolean)

                    End Select
                Next

                readawinSettings = True

            Else

                readawinSettings = False

            End If

        Catch ex As Exception

            readawinSettings = False

        End Try

    End Function


    ''' <summary>
    ''' Erstellen eines Powerpoint-Reports auf Grund von einem ReportProfil, TimeRange, DB Zugriff, und ausgewählte EinzelProjekte oder Konstellationen
    ''' </summary>
    ''' <param name="projekte">Projektname oder Konstellationsname</param>
    ''' <param name="variante">Variante eines Projektes</param>
    ''' <param name="profilname">Name des Reportprofils</param>
    ''' <param name="vonDate">von Zeit</param>
    ''' <param name="bisDate">bis Zeit</param>
    ''' <param name="reportname">Name des Report (wie abgespeichert werden soll)</param>
    ''' <param name="dbUsername">DB User</param>
    ''' <param name="dbPassword">DB pwd</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function reportErstellen(ByVal projekte As String, ByVal variante As String, ByVal profilname As String, ByVal timestamp As Date, _
                                        ByVal vonDate As Date, ByVal bisDate As Date, ByVal reportname As String, ByVal append As Boolean, _
                                        ByVal dbUsername As String, ByVal dbPassword As String) As Boolean


        Dim currentPresentationName As String = ""
        Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPassword)
        Dim reportProfil As clsReportAll = XMLImportReportProfil(profilname)
        Dim zeilenhoehe As Double = 0.0     ' zeilenhöhe muss für alle Projekte gleich sein, daher mit übergeben
        Dim legendFontSize As Single = 0.0  ' FontSize der Legenden der Schriftgröße des Projektnamens angepasst

        Dim selectedPhases As New Collection
        Dim selectedMilestones As New Collection
        Dim selectedRoles As New Collection
        Dim selectedCosts As New Collection
        Dim selectedBUs As New Collection
        Dim selectedTypes As New Collection

        reportErstellen = False

        selectedPhases = copySortedListtoColl(reportProfil.Phases)
        selectedMilestones = copySortedListtoColl(reportProfil.Milestones)
        selectedRoles = copySortedListtoColl(reportProfil.Roles)
        selectedCosts = copySortedListtoColl(reportProfil.Costs)
        selectedBUs = copySortedListtoColl(reportProfil.BUs)
        selectedTypes = copySortedListtoColl(reportProfil.Typs)

        With awinSettings

            .mppExtendedMode = reportProfil.ExtendedMode
            .mppOnePage = reportProfil.OnePage
            .mppShowAllIfOne = reportProfil.AllIfOne
            .mppShowAmpel = reportProfil.Ampeln
            .mppShowLegend = reportProfil.Legend
            .mppShowMsDate = reportProfil.MSDate
            .mppShowMsName = reportProfil.MSName
            .mppShowPhDate = reportProfil.PhDate
            .mppShowPhName = reportProfil.PhName
            .mppShowProjectLine = reportProfil.ProjectLine
            .mppSortiertDauer = reportProfil.SortedDauer
            .mppVertikalesRaster = reportProfil.VLinien
            .mppFullyContained = reportProfil.FullyContained
            .mppShowHorizontals = reportProfil.ShowHorizontals
            .mppUseAbbreviation = reportProfil.UseAbbreviation
            .mppUseOriginalNames = reportProfil.UseOriginalNames
            .mppKwInMilestone = reportProfil.KwInMilestone
            .mppProjectsWithNoMPmayPass = reportProfil.projectsWithNoMPmayPass

        End With

        If Not (IsNothing(vonDate) Or vonDate = Date.MinValue) Then
            showRangeLeft = getColumnOfDate(vonDate)
        Else
            showRangeLeft = 0
        End If
        If Not (IsNothing(bisDate) Or bisDate = Date.MinValue) Then
            showRangeRight = getColumnOfDate(bisDate)
        Else
            showRangeRight = 0
        End If


        Try
            If Not reportProfil.isMpp Then

                Try


                    Dim vorlagendateiname As String = awinPath & RepProjectVorOrdner & "\" & reportProfil.PPTTemplate
                    If My.Computer.FileSystem.FileExists(vorlagendateiname) Then

                        'Das gewählte Projekt reporten

                        Dim hproj As New clsProjekt
                        hproj = request.retrieveOneProjectfromDB(projekte, variante, timestamp)

                        If Not IsNothing(hproj) Then

                            Dim key As String = calcProjektKey(hproj)

                            ' diese Liste wird benötigt, damit in zeichneMultiprojektSicht die Routine bestimmeProjekteAndMinMaxDates funktioniert
                            If Not AlleProjekte.Containskey(calcProjektKey(hproj)) Then
                                AlleProjekte.Add(hproj)
                            End If


                            If Not ShowProjekte.contains(hproj.name) Then  ' akt. Projekt nicht in ShowProjekte
                                ShowProjekte.Add(hproj)

                            Else   ' es ist eventuell nicht die richtige Variante enthalten
                                If ShowProjekte.getProject(hproj.name).variantName <> hproj.variantName Then
                                    ShowProjekte.Remove(hproj.name)
                                    ShowProjekte.Add(hproj)
                                End If
                            End If

                            Call createPPTSlidesFromProject(hproj, vorlagendateiname, _
                                                        selectedPhases, selectedMilestones, _
                                                        selectedRoles, selectedCosts, _
                                                        selectedBUs, selectedTypes, True, _
                                                        True, zeilenhoehe, legendFontSize, _
                                                        Nothing, Nothing)


                            Dim pptApp As Microsoft.Office.Interop.PowerPoint.Application = Nothing
                            Try
                                ' prüft, ob bereits Powerpoint geöffnet ist 
                                pptApp = CType(GetObject(, "PowerPoint.Application"), Microsoft.Office.Interop.PowerPoint.Application)
                            Catch ex As Exception
                                Try
                                    pptApp = CType(CreateObject("PowerPoint.Application"), Microsoft.Office.Interop.PowerPoint.Application)

                                Catch ex1 As Exception
                                    Call MsgBox("Powerpoint konnte nicht gestartet werden ..." & ex1.Message)
                                    reportErstellen = False
                                    Exit Function
                                End Try

                            End Try
                            ' aktive Präsentation unter angegebenem Namen "reportname" abspeichern
                            Dim currentPraesi As Microsoft.Office.Interop.PowerPoint.Presentation = pptApp.ActivePresentation

                            If reportname = "" Then
                                Dim aktDate As String = Date.Now.ToString
                                reportname = aktDate & "Report.pptx"
                                Call logfileSchreiben("EinzelprojektReport mit ' " & projekte & "/" & variante & "/" & _
                                                      profilname & "/ ... wurde in " & reportname & "ersatzweise gespeichert", "reportErstellen", anzFehler)
                            Else
                                reportname = reportname & ".pptx"
                            End If

                            If My.Computer.FileSystem.FileExists(reportOrdnerName & reportname) And append Then

                                ' die Seiten 2 - ende der vorhandenen Powerpoint-Datei müssen in das currentPraesi eingefügt werden
                                Dim oldPraesi As Microsoft.Office.Interop.PowerPoint.Presentation = pptApp.Presentations.Open(reportOrdnerName & reportname)
                                Dim anzoldSlides As Integer = oldPraesi.Slides.Count
                                oldPraesi.Close()

                                currentPraesi.Slides.InsertFromFile(FileName:=reportOrdnerName & reportname, Index:=1, SlideStart:=2, SlideEnd:=anzoldSlides)
                                currentPraesi.SaveAs(reportOrdnerName & reportname)
                                currentPraesi.Close()
                            Else
                                'If My.Computer.FileSystem.FileExists(reportOrdnerName & reportname & ".pptx") Then
                                '    My.Computer.FileSystem.DeleteFile(reportOrdnerName & reportname & ".pptx")
                                'End If
                                currentPraesi.SaveAs(reportOrdnerName & reportname)
                                currentPraesi.Close()


                            End If

                            reportErstellen = True
                        Else

                            Call logfileSchreiben("reportErstellen", "Projekt '" & projekte & "' existiert nicht in DB!", anzFehler)

                        End If
                    Else
                        Call logfileSchreiben("reportErstellen", "Vorlagendatei " & vorlagendateiname & " existiert nicht!", anzFehler)
                    End If

                Catch ex As Exception

                End Try

            Else    ' isMPP

                Try

                    If Not (showRangeLeft > 0 And showRangeRight > showRangeLeft) Then

                        showRangeLeft = getColumnOfDate(reportProfil.VonDate)
                        showRangeRight = getColumnOfDate(reportProfil.BisDate)

                    End If

                    Dim hproj As New clsProjekt
                    Dim constellations As New clsConstellations
                    constellations = request.retrieveConstellationsFromDB()
                    If Not IsNothing(constellations) Then

                        Dim curconstellation As clsConstellation = constellations.getConstellation(projekte)

                        If Not IsNothing(curconstellation) Then

                            For Each kvp As KeyValuePair(Of String, clsConstellationItem) In curconstellation.Liste

                                hproj = request.retrieveOneProjectfromDB(kvp.Value.projectName, kvp.Value.variantName, timestamp)

                                If Not IsNothing(hproj) Then

                                    Dim key As String = calcProjektKey(hproj)

                                    ' diese Liste wird benötigt, damit in zeichneMultiprojektSicht die Routine bestimmeProjekteAndMinMaxDates funktioniert
                                    If Not AlleProjekte.Containskey(calcProjektKey(hproj)) Then
                                        AlleProjekte.Add(hproj)
                                    End If

                                    If kvp.Value.show Then

                                        If Not ShowProjekte.contains(hproj.name) Then
                                            ShowProjekte.Add(hproj)
                                        Else
                                            If ShowProjekte.getProject(hproj.name).variantName <> kvp.Value.variantName Then
                                                ShowProjekte.Remove(kvp.Value.projectName)
                                                ShowProjekte.Add(hproj)
                                            End If
                                        End If

                                    End If
                                Else

                                    Call logfileSchreiben("reportErstellen", "Projekt '" & kvp.Value.projectName & " mit TimeStamp '" & timestamp.ToString & "' existiert nicht in DB!", anzFehler)

                                End If  ' if hproj existiert
                            Next


                            Dim vorlagendateiname As String = awinPath & RepPortfolioVorOrdner & "\" & reportProfil.PPTTemplate
                            If My.Computer.FileSystem.FileExists(vorlagendateiname) Then

                                Call createPPTSlidesFromConstellation(vorlagendateiname, _
                                                                      selectedPhases, selectedMilestones, _
                                                                      selectedRoles, selectedCosts, _
                                                                      selectedBUs, selectedTypes, True, _
                                                                      Nothing, Nothing)

                                Dim pptApp As Microsoft.Office.Interop.PowerPoint.Application = Nothing
                                Try
                                    ' prüft, ob bereits Powerpoint geöffnet ist 
                                    pptApp = CType(GetObject(, "PowerPoint.Application"), Microsoft.Office.Interop.PowerPoint.Application)
                                Catch ex As Exception
                                    Try
                                        pptApp = CType(CreateObject("PowerPoint.Application"), Microsoft.Office.Interop.PowerPoint.Application)

                                    Catch ex1 As Exception
                                        Call MsgBox("Powerpoint konnte nicht gestartet werden ..." & ex1.Message)
                                        reportErstellen = False
                                        Exit Function
                                    End Try

                                End Try

                                ' aktive Präsentation unter angegebenem Namen "reportname" abspeichern
                                Dim currentPraesi As Microsoft.Office.Interop.PowerPoint.Presentation = pptApp.ActivePresentation

                                If reportname = "" Then
                                    Dim aktDate As String = Date.Now.ToString
                                    reportname = aktDate & "MP Report.pptx"
                                    Call logfileSchreiben("MulitprojektReport mit ' " & projekte & "/" & _
                                                          profilname & "/ ... wurde in " & reportname & "ersatzweise gespeichert", "reportErstellen", anzFehler)
                                Else
                                    reportname = reportname & ".pptx"
                                End If

                                If My.Computer.FileSystem.FileExists(reportOrdnerName & reportname) And append Then

                                    ' die Seiten 2 - ende der vorhandenen Powerpoint-Datei müssen in das currentPraesi eingefügt werden
                                    Dim oldPraesi As Microsoft.Office.Interop.PowerPoint.Presentation = pptApp.Presentations.Open(reportOrdnerName & reportname)
                                    Dim anzoldSlides As Integer = oldPraesi.Slides.Count
                                    oldPraesi.Close()

                                    currentPraesi.Slides.InsertFromFile(FileName:=reportOrdnerName & reportname, Index:=1, SlideStart:=2, SlideEnd:=anzoldSlides)
                                    currentPraesi.SaveAs(reportOrdnerName & reportname)
                                    currentPraesi.Close()
                                Else
                                    'If My.Computer.FileSystem.FileExists(reportOrdnerName & reportname & ".pptx") Then
                                    '    My.Computer.FileSystem.DeleteFile(reportOrdnerName & reportname & ".pptx")
                                    'End If
                                    currentPraesi.SaveAs(reportOrdnerName & reportname)
                                    currentPraesi.Close()


                                End If

                                reportErstellen = True

                            End If
                        Else
                            Call logfileSchreiben("reportErstellen", "angegebene Constellation nicht in der DB", anzFehler)

                        End If

                    Else
                        Call logfileSchreiben("reportErstellen", "keine Constellations in der DB vorhanden", anzFehler)
                    End If

                Catch ex As Exception

                End Try

            End If



        Catch ex As Exception

            Call MsgBox("Fehler: " & vbLf & ex.Message)

            reportErstellen = False
        End Try

        'pptApp.Quit()

    End Function


    ''' <summary>
    ''' behandelt die Missing Definitions, nimmt ggf in 
    ''' </summary>
    ''' <param name="definitionName"></param>
    ''' <param name="isVorlage"></param>
    ''' <param name="isMilestone"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function isMissingDefinitionOK(ByVal definitionName As String, ByVal isVorlage As Boolean, ByVal isMilestone As Boolean) As Boolean

        Dim checkResult As Boolean = True

        If isMilestone And Not MilestoneDefinitions.Contains(definitionName) Then
            ' Behandlung Meilenstein Definition, aber nur wenn nicht enthalten ... 

            Dim hMilestoneDef As New clsMeilensteinDefinition

            With hMilestoneDef
                .name = definitionName
                .belongsTo = ""
                .shortName = ""
                .darstellungsKlasse = ""
                .UID = MilestoneDefinitions.Count + 1
            End With

            If (isVorlage And awinSettings.alwaysAcceptTemplateNames) Or _
                awinSettings.addMissingPhaseMilestoneDef Then
                ' in die Milestone-Definitions aufnehmen 
                Try
                    If Not MilestoneDefinitions.Contains(hMilestoneDef.name) Then
                        MilestoneDefinitions.Add(hMilestoneDef)
                    End If

                Catch ex As Exception
                End Try

            Else


                ' in die Missing Milestone-Definitions aufnehmen 
                Try
                    ' das Element aufnehmen, in Abhängigkeit vom Setting 
                    If awinSettings.importUnknownNames Then
                        checkResult = True
                    Else
                        checkResult = False
                    End If

                    If Not missingMilestoneDefinitions.Contains(hMilestoneDef.name) Then
                        missingMilestoneDefinitions.Add(hMilestoneDef)
                    End If

                Catch ex As Exception
                End Try
            End If

        ElseIf Not isMilestone And Not (PhaseDefinitions.Contains(definitionName)) Then

            ' Behandlung Phasen 
            Dim hphaseDef As clsPhasenDefinition
            hphaseDef = New clsPhasenDefinition

            hphaseDef.darstellungsKlasse = ""
            hphaseDef.shortName = ""
            hphaseDef.name = definitionName
            hphaseDef.UID = PhaseDefinitions.Count + 1



            If (isVorlage And awinSettings.alwaysAcceptTemplateNames) Or _
                awinSettings.addMissingPhaseMilestoneDef Then
                ' in die Phase-Definitions aufnehmen 
                checkResult = True
                Try
                    If Not PhaseDefinitions.Contains(hphaseDef.name) Then
                        PhaseDefinitions.Add(hphaseDef)
                    End If
                Catch ex As Exception
                    checkResult = False
                End Try
            Else
                ' in Abhängigkeit vom Setting die Elemente aufnehmen oder nicht 
                Try
                    If awinSettings.importUnknownNames Then
                        checkResult = True
                    Else
                        checkResult = False
                    End If

                    If Not missingPhaseDefinitions.Contains(hphaseDef.name) Then
                        missingPhaseDefinitions.Add(hphaseDef)
                    End If

                Catch ex As Exception
                    checkResult = False
                End Try


            End If
        End If

        isMissingDefinitionOK = checkResult

    End Function


    ''' <summary>
    ''' setzt die Projekt-Historie für das angegebene Projekt
    ''' wenn nicht existiert, wird Projekt-Historie auf Nothing gesetzt 
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <param name="vName"></param>
    ''' <remarks></remarks>
    Public Sub setProjektHistorie(ByVal pName As String, ByVal vName As String)

        Dim holeHistory As Boolean = True
        Dim vglProj As clsProjekt = Nothing

        If Not projekthistorie Is Nothing Then
            If projekthistorie.Count > 0 Then
                vglProj = projekthistorie.First
            End If
        End If

        If Not noDB Then

            If Not IsNothing(vglProj) Then
                If vglProj.name = pName And vglProj.variantName = vName Then
                    holeHistory = False
                End If
            End If
            

            If holeHistory Then
                Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
                If Request.pingMongoDb() Then
                    Try
                        If request.projectNameAlreadyExists(pName, vName, Date.Now) Then
                            projekthistorie.liste = request.retrieveProjectHistoryFromDB(projectname:=pName, variantName:=vName, _
                                                                        storedEarliest:=Date.MinValue, storedLatest:=Date.Now)
                        Else
                            projekthistorie.clear()
                        End If
                        
                    Catch ex As Exception
                        projekthistorie.clear()
                    End Try
                Else
                    If awinSettings.englishLanguage Then
                        Call MsgBox("Database Connection failed ...")
                    Else
                        Call MsgBox("Datenbank-Verbindung ist unterbrochen!")
                    End If

                End If
            End If

        End If
    End Sub

    ''' <summary>
    ''' aktualisiert mit dem selektierten Projekt die evtl angezeigten Projekt-Info Charts
    ''' replaceProj = false, wenn die Skalierung nicht angepasst werden soll; also z.Bsp bei Aufruf aus Time-Machine 
    ''' </summary>
    ''' <param name="hproj">das selektierte Projekt</param>
    ''' <remarks></remarks>
    Public Sub aktualisiereCharts(ByVal hproj As clsProjekt, ByVal replaceProj As Boolean)
        Dim chtobj As Excel.ChartObject

        Dim vglName As String = hproj.name.Trim
        Dim founddiagram As New clsDiagramm
        ' ''Dim IDkennung As String

        Dim currentWsName As String
        If visboZustaende.projectBoardMode = ptModus.graficboard Then
            currentWsName = arrWsNames(ptTables.mptPrCharts)
        Else
            currentWsName = arrWsNames(ptTables.meCharts)
        End If

        ' aktualisieren der Window Caption ...
        Try
            If visboWindowExists(PTwindows.mptpr) Then
                Dim tmpmsg As String = hproj.getShapeText & " (" & hproj.timeStamp.ToString & ")"
                projectboardWindows(PTwindows.mptpr).Caption = bestimmeWindowCaption(PTwindows.mptpr, tmpmsg)
            End If
        Catch ex As Exception

        End Try


        If Not (hproj Is Nothing) Then

            With CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(currentWsName), Excel.Worksheet)
                Dim tmpArray() As String
                Dim anzDiagrams As Integer
                anzDiagrams = CType(.ChartObjects, Excel.ChartObjects).Count

                If anzDiagrams > 0 Then
                    For i = 1 To anzDiagrams
                        chtobj = CType(.ChartObjects(i), Excel.ChartObject)
                        If chtobj.Name <> "" Then
                            tmpArray = chtobj.Name.Split(New Char() {CType("#", Char)}, 5)
                            ' chtobj name ist aufgebaut: pr#PTprdk.kennung#pName#Auswahl
                            If tmpArray(0) = "pr" Then

                                Dim chartTyp As String = ""
                                Dim typID As Integer = -1
                                Dim auswahl As Integer = -1
                                Dim chartPname As String = ""
                                Call getChartKennungen(chtobj.Name, chartTyp, typID, auswahl, chartPname)

                                If replaceProj Or (chartPname.Trim = vglName) Then
                                    Select Case typID


                                        ' replaceProj sorgt in den nachfolgenden Sequenzen dafür, daß das Chart im Falle eines Aufrufes aus der 
                                        ' Time-Machine (replaceProj = false) nicht in der Skalierung angepasst wird; das geschieht initial beim Laden der Time-Machine
                                        ' wenn es aus dem Selektieren von Projekten aus aufgerufen wird, dann wird die optimal passende Skalierung schon jedesmal berechnet 

                                        Case PTprdk.Phasen
                                            ' Update Phasen Diagramm

                                            If CInt(tmpArray(3)) = PThis.current Then
                                                ' nur dann muss aktualisiert werden ...
                                                Call updatePhasesBalken(hproj, chtobj, auswahl, replaceProj)
                                            End If


                                        Case PTprdk.PersonalBalken

                                            Call updateRessBalkenOfProject(hproj, chtobj, auswahl, replaceProj)


                                        Case PTprdk.PersonalPie


                                            ' Update Pie-Diagramm
                                            Call updateRessPieOfProject(hproj, chtobj, auswahl)


                                        Case PTprdk.KostenBalken


                                            Call updateCostBalkenOfProject(hproj, chtobj, auswahl, replaceProj)


                                        Case PTprdk.KostenPie


                                            Call updateCostPieOfProject(hproj, chtobj, auswahl)


                                        Case PTprdk.StrategieRisiko

                                            Call updateProjectPfDiagram(hproj, chtobj, auswahl)

                                        Case PTprdk.FitRisikoVol

                                            Call updateProjectPfDiagram(hproj, chtobj, auswahl)

                                        Case PTprdk.ComplexRisiko

                                            Call updateProjectPfDiagram(hproj, chtobj, auswahl)

                                        Case PTprdk.Ergebnis
                                            ' Update Ergebnis Diagramm
                                            Call updateProjektErgebnisCharakteristik2(hproj, chtobj, auswahl, replaceProj)

                                        Case PTprdk.SollIstGesamtkosten

                                            Call setProjektHistorie(hproj.name, hproj.variantName)
                                            Call updateSollIstOfProject(hproj, chtobj, Date.Now, auswahl, "", True, False)

                                        Case PTprdk.SollIstPersonalkosten

                                            Call setProjektHistorie(hproj.name, hproj.variantName)
                                            Call updateSollIstOfProject(hproj, chtobj, Date.Now, auswahl, "", True, False)

                                        Case PTprdk.SollIstSonstKosten

                                            Call setProjektHistorie(hproj.name, hproj.variantName)
                                            Call updateSollIstOfProject(hproj, chtobj, Date.Now, auswahl, "", True, False)

                                        Case Else


                                    End Select

                                End If

                            End If

                        End If

                    Next
                End If

            End With

        End If

    End Sub

End Module
