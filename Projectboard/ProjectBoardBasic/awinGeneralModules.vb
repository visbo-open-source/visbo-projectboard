
'Option Explicit On
'Option Strict On

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
Imports Microsoft.VisualBasic
Imports ProjectBoardBasic
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
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub awinWritePhaseMilestoneDefinitions(Optional ByVal writeMappings As Boolean = False)

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
        ' (4) ist Registerblatt Einstellungen 
        Dim wsName4 As Excel.Worksheet = CType(appInstance.Worksheets(arrWsNames(4)), _
                                                Global.Microsoft.Office.Interop.Excel.Worksheet)

        phaseDefs = wsName4.Range("awin_Phasen_Definition")

        ' diese Range sollte auf alle Fälle mindestens eine Zeile haben 
        Dim anzZeilen As Integer = phaseDefs.Rows.Count
        lastrow = CType(phaseDefs.Rows(anzZeilen), Excel.Range)
        firstrow = CType(phaseDefs.Rows(1), Excel.Range)


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

        ' jetzt können erst diue PhaseDefinitions, dann die MilestoneDefiitions geschrieben werden 



        ' hier muss erst mal geprüft werden, ob Zeilen eingefügt oder gelöscht werden müssen 
        ' anzZeilen muss immer um 2 größer sein als die Anzahl der Definitionen ; 
        ' die erste und letzte Zeile des Bereichs sind leer  
        Dim anzDefinitions As Integer = PhaseDefinitions.Count
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

            With PhaseDefinitions.getPhaseDef(ix)
                phName = .name
                shortName = .shortName
                darstellungsKlasse = .darstellungsKlasse
            End With

            CType(firstrow.Cells(ix, 1), Excel.Range).Offset(1, 0).Value = phName.ToString
            CType(firstrow.Cells(ix, 1), Excel.Range).Offset(1, 5).Value = shortName
            CType(firstrow.Cells(ix, 1), Excel.Range).Offset(1, 6).Value = darstellungsKlasse


        Next ix




        '
        ' jetzt werden die Meilensteine geschrieben 
        '

        ' erste , letzte Zeile des Meilenstein Ranges setzen 
        ' diese Range sollte auf alle Fälle mindestens eine Zeile haben 

        milestoneDefs = wsName4.Range("awin_Meilenstein_Definition")
        anzZeilen = milestoneDefs.Rows.Count
        lastrow = CType(milestoneDefs.Rows(anzZeilen), Excel.Range)
        firstrow = CType(milestoneDefs.Rows(1), Excel.Range)


        ' hier muss erst mal geprüft werden, ob Zeilen eingefügt oder gelöscht werden müssen 
        anzDefinitions = MilestoneDefinitions.Count
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

            With MilestoneDefinitions.getMilestoneDef(ix)
                msName = .name
                shortName = .shortName
                darstellungsKlasse = .darstellungsKlasse
            End With

            CType(firstrow.Cells(ix, 1), Excel.Range).Offset(1, 0).Value = msName.ToString
            CType(firstrow.Cells(ix, 1), Excel.Range).Offset(1, 5).Value = shortName
            CType(firstrow.Cells(ix, 1), Excel.Range).Offset(1, 6).Value = darstellungsKlasse


        Next ix



        '
        ' Ende der Behandlung der Phasen-/Meilenstein Behandlung 

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
    ''' liest das Customization File aus und initialisiert die globalen Variablen entsprechend
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub awinsetTypen()

        Dim i As Integer
        Dim xlsCustomization As Excel.Workbook = Nothing

        ReDim importOrdnerNames(6)
        ReDim exportOrdnerNames(4)


        ' hier werden die Ordner Namen für den Import wie Export festgelegt ... 
        'awinPath = appInstance.ActiveWorkbook.Path & "\"

        '' ''If (Dir(awinSettings.globalPath, vbDirectory) <> "") Then
        '' ''    globalPath = awinSettings.globalPath
        '' ''Else

        '' ''    Throw New ArgumentException("Globaler Requirementsordner " & awinSettings.globalPath & " existiert nicht")

        '' ''End If

        '' ''If (Dir(awinSettings.awinPath, vbDirectory) <> "") Then
        '' ''    awinPath = awinSettings.awinPath
        '' ''Else

        '' ''    Throw New ArgumentException("Lokaler Requirementsordner " & awinSettings.awinPath & " existiert nicht")

        '' ''End If


        globalPath = awinSettings.globalPath
        awinPath = awinSettings.awinPath


        If awinPath = "" And globalPath <> "" Then
            awinPath = globalPath
        ElseIf globalPath = "" And awinPath <> "" Then
            globalPath = awinPath
        ElseIf globalPath = "" And awinPath = "" Then
            Throw New ArgumentException("Globaler Ordner " & awinSettings.globalPath & " und Lokaler Ordner " & awinSettings.awinPath & " wurden nicht angegeben")
        End If

        If (Dir(globalPath, vbDirectory) = "") Then
            If (Dir(awinPath, vbDirectory) = "") Then
                Throw New ArgumentException("Requirementsordner " & awinSettings.globalPath & " existiert nicht")
            Else
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

        exportOrdnerNames(PTImpExp.visbo) = awinPath & "Export\VISBO Steckbriefe"
        exportOrdnerNames(PTImpExp.rplan) = awinPath & "Export\RPLAN-Excel"
        exportOrdnerNames(PTImpExp.msproject) = awinPath & "Export\MSProject"
        exportOrdnerNames(PTImpExp.simpleScen) = awinPath & "Export\einfache Szenarien"
        exportOrdnerNames(PTImpExp.modulScen) = awinPath & "Export\modulare Szenarien"


        If globalPath <> awinPath Then
            Call synchronizeGlobalToLocalFolder()
        End If


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

        ' bestimmen der maximalen Breite und Höhe 
        Dim formerSU As Boolean = appInstance.ScreenUpdating
        appInstance.ScreenUpdating = False


        ' um dahinter temporär die Darstellungsklassen kopieren zu können  
        Dim projectBoardSheet As Excel.Worksheet = CType(appInstance.ActiveSheet, _
                                                Global.Microsoft.Office.Interop.Excel.Worksheet)



        With appInstance.ActiveWindow


            If .WindowState = Excel.XlWindowState.xlMaximized Then
                maxScreenHeight = .Height
                maxScreenWidth = .Width
            Else
                Dim formerState As Excel.XlWindowState = .WindowState
                .WindowState = Excel.XlWindowState.xlMaximized
                maxScreenHeight = .Height
                maxScreenWidth = .Width
                .WindowState = formerState
            End If


        End With

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

        ' With listOfWorkSheets(arrWsNames(4))

        ' Logfile öffnen und ggf. initialisieren
        Call logfileOpen()


        ' Auslesen des Window Namens 
        Dim accountToken As IntPtr = WindowsIdentity.GetCurrent().Token
        Dim myUser As New WindowsIdentity(accountToken)
        myWindowsName = myUser.Name
        Call logfileSchreiben("Windows-User: ", myWindowsName, anzFehler)


        ' hier muss jetzt das Customization File aufgemacht werden ...
        Try
            xlsCustomization = appInstance.Workbooks.Open(awinPath & customizationFile)
            myCustomizationFile = appInstance.ActiveWorkbook.Name
        Catch ex As Exception
            appInstance.ScreenUpdating = formerSU
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



        ' ur: 23.01.2015: Abfragen der Login-Informationen
        loginErfolgreich = loginProzedur()


        If Not loginErfolgreich Then
            ' Customization-File wird geschlossen
            xlsCustomization.Close(SaveChanges:=False)
            Call logfileSchreiben("LOGIN fehlerhaft", "", -1)
            Call logfileSchliessen()
            appInstance.Quit()
            Exit Sub
        Else




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

                ' jetzt muss die Seite mit den Appearance-Shapes kopiert werden 
                appInstance.EnableEvents = False
                CType(appInstance.Worksheets(arrWsNames(7)), _
                Global.Microsoft.Office.Interop.Excel.Worksheet).Copy(After:=projectBoardSheet)

                ' hier wird die Datei Projekt Tafel Customizations als aktives workbook wieder geschlossen ....
                appInstance.Workbooks(myCustomizationFile).Close(SaveChanges:=False) ' ur: 6.5.2014 savechanges hinzugefügt
                appInstance.EnableEvents = True


                ' jetzt muss die apperanceDefinitions wieder neu aufgebaut werden 
                appearanceDefinitions.Clear()
                wsName7810 = CType(appInstance.Worksheets(arrWsNames(7)), _
                                                        Global.Microsoft.Office.Interop.Excel.Worksheet)
                Call aufbauenAppearanceDefinitions(wsName7810)


                ' jetzt werden die ggf vorhandenen detaillierten Ressourcen Kapazitäten ausgelesen 
                Call readRessourcenDetails()


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

                ' jetzt werden aus der Datenbank die Konstellationen und Dependencies gelesen 
                Call readInitConstellations()

            Catch ex As Exception
                appInstance.ScreenUpdating = formerSU
                appInstance.EnableEvents = True
                Throw New ArgumentException(ex.Message)
            End Try

            ' Logfile wird geschlossen
            Call logfileSchliessen()


        End If  ' von "if Login erfolgt"


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
    ''' setzt Kalenderleiste und Spaltenbreite sowie -Höhe 
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub prepareProjektTafel()


        ' bestimmen der Spaltenbreite und Spaltenhöhe ...
        Dim testCase As String = appInstance.ActiveWorkbook.Name
        If testCase <> myProjektTafel Then
            CType(appInstance.Workbooks(myProjektTafel), Excel.Workbook).Activate()
        End If

        Dim wsName3 As Excel.Worksheet = CType(appInstance.Worksheets(arrWsNames(3)), _
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
                    rng.NumberFormat = "mmm-yy"

                    Dim destinationRange As Excel.Range = .Range(.Cells(1, 1), .Cells(1, 480))
                    With destinationRange
                        .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                        .VerticalAlignment = Excel.XlVAlign.xlVAlignBottom
                        .NumberFormat = "mmm-yy"
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
    Public Sub awinChangeTimeSpan(ByVal von As Integer, ByVal bis As Integer)

        'Dim k As Integer

        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim formerSU As Boolean = appInstance.ScreenUpdating
        Dim noTimeFrame As Boolean = False

        appInstance.EnableEvents = False



        If von < 1 Then
            von = 1
        End If

        If bis < von + 5 Then
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

                    Try
                        nameSopTyp = tmpStr(0).Trim
                        pName = nameSopTyp
                        Try
                            nameBU = tmpStr(1)
                            tmpStr = nameBU.Split(New Char() {CChar(" ")}, 3)
                            nameBU = tmpStr(0)
                        Catch ex1 As Exception
                            nameBU = ""
                        End Try


                    Catch ex As Exception
                        Throw New Exception("Name, SOP, Typ kann nicht bestimmt werden " & vbLf & nameSopTyp)
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
                    ImportProjekte.Add(calcProjektKey(hproj), hproj)
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

        ' hier wird eingetragen, welches vordefinierte Flag das customized Field VISBO repräsentiert
        Dim visboflag As MSProject.PjField = Nothing

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

                Try
                    visboflag = CType(prj.FieldNameToFieldConstant("VISBO", MSProject.PjFieldType.pjTask), MSProject.PjField)
                Catch ex As Exception
                    visboflag = 0
                End Try

               
                If modus = "BHTC" Then
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
                hproj.name = hhstr(0)
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

                        ' nachsehen, ob msTask.Name in PhaseDefinitions definiert ist
                        If Not PhaseDefinitions.Contains(msTask.Name) Then
                            Dim newPhaseDef As New clsPhasenDefinition
                            newPhaseDef.name = msTask.Name
                            newPhaseDef.shortName = msTask.Name
                            newPhaseDef.UID = PhaseDefinitions.Count + 1
                            'PhaseDefinitions.Add(newPhaseDef)
                            If Not missingPhaseDefinitions.Contains(newPhaseDef.name) Then
                                missingPhaseDefinitions.Add(newPhaseDef)
                            End If

                        End If

                        With cphase

                            If Not istElemID(msTask.Name) Then
                                .nameID = hproj.hierarchy.findUniqueElemKey(msTask.Name, False)
                            End If

                            If visboflag <> 0 Then          ' VISBO-Flag ist definiert

                                ' Liste, ob Task in Projekt für die Projekt-Tafel aufgenommen werden soll, oder nicht
                                visboFlagListe.Add(.nameID, msTask.GetField(visboflag) = "Yes")

                            End If

                            ' Änderung 28.11.13: jetzt wird die Phasen Länge exakt bestimmt , über startoffset in Tagen und dauerinDays als Länge
                            Dim cphaseStartOffset As Long
                            Dim dauerIndays As Long
                            cphaseStartOffset = DateDiff(DateInterval.Day, hproj.startDate, CDate(msTask.Start))
                            dauerIndays = calcDauerIndays(CDate(msTask.Start), CDate(msTask.Finish))
                            .changeStartandDauer(cphaseStartOffset, dauerIndays)
                            .Offset = 0

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

                                            Try
                                                k = CInt(CostDefinitions.getCostdef(ass.ResourceName).UID)
                                            Catch ex As Exception
                                                ' Kostenart existiert noch nicht
                                                ' wird hier neu aufgenommen
                                                Dim newCostDef As New clsKostenartDefinition
                                                newCostDef.name = ass.ResourceName
                                                newCostDef.farbe = RGB(120, 120, 120)   ' Farbe: grau
                                                newCostDef.UID = CostDefinitions.Count + 1
                                                If Not missingCostDefinitions.Contains(newCostDef.name) Then
                                                    missingCostDefinitions.Add(newCostDef)
                                                End If

                                                CostDefinitions.Add(newCostDef)

                                                k = CInt(missingCostDefinitions.getCostdef(ass.ResourceName).UID)
                                            End Try

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

                                            Try
                                                r = CInt(RoleDefinitions.getRoledef(ass.ResourceName).UID)
                                            Catch ex As Exception
                                                ' Rolle existiert noch nicht
                                                ' wird hier neu aufgenommen

                                                Dim newRoleDef As New clsRollenDefinition
                                                newRoleDef.name = ass.ResourceName
                                                newRoleDef.farbe = RGB(120, 120, 120)
                                                newRoleDef.Startkapa = 200000

                                                ' OvertimeRate in Tagessatz umrechnen
                                                Dim hoverstr() As String = Split(CStr(ass.Resource.OvertimeRate), "/", -1)
                                                hoverstr = Split(hoverstr(0), "$", -1)
                                                newRoleDef.tagessatzExtern = CType(hoverstr(1), Double) * msproj.HoursPerDay

                                                ' StandardRate in Tagessatz umrechnen
                                                Dim hstdstr() As String = Split(CStr(ass.Resource.StandardRate), "/", -1)
                                                hstdstr = Split(hstdstr(0), "$", -1)
                                                newRoleDef.tagessatzIntern = CType(hstdstr(1), Double) * msproj.HoursPerDay

                                                newRoleDef.UID = RoleDefinitions.Count + 1
                                                If Not missingRoleDefinitions.Contains(newRoleDef.name) Then
                                                    missingRoleDefinitions.Add(newRoleDef)
                                                End If

                                                RoleDefinitions.Add(newRoleDef)

                                                r = CInt(missingRoleDefinitions.getRoledef(ass.ResourceName).UID)
                                            End Try

                                            Dim work As Double = CType(ass.Work, Double)
                                            'Dim duration As Double = CType(ass.Duration, Double)
                                            Dim unit As Double = CType(ass.Units, Double)

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
                                                    ' Xwerte in Anzahl Tage; in MSProject alle Werte in anz. Minuten
                                                    Xwerte(m - 1) = CType(work / anzmonth / 60 / 8, Double)

                                                Catch ex As Exception
                                                    Xwerte(m - 1) = 0.0
                                                End Try

                                            Next m

                                            crole = New clsRolle(anzmonth - 1)
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

                            hproj.AddPhase(cphase, parentID:=hrchynode.parentNodeKey)

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

                        End If

                        msPhase = hproj.getPhaseByID(parentID)

                        Dim cmilestone As New clsMeilenstein(msPhase)

                        ' prüfen, ob MeilensteinDefinition bereits vorhanden
                        If Not MilestoneDefinitions.Contains(msTask.Name) Then
                            Dim msDef As New clsMeilensteinDefinition
                            msDef.belongsTo = msPhase.name
                            msDef.name = msTask.Name
                            msDef.schwellWert = 0
                            msDef.shortName = ""
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
                            cmilestone.addBewertung(msBewertung)


                            If visboflag <> 0 Then        ' Ist VISBO-flag definiert?

                                ' Liste, ob Meilenstein in Projekt für die Projekt-Tafel aufgenommen werden soll, oder nicht
                                visboFlagListe.Add(cmilestone.nameID, msTask.GetField(visboflag) = "Yes")

                            End If

                            Try
                                With msPhase
                                    .addMilestone(cmilestone)
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

                AlleProjekte.Add(key, hproj)

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
    ''' Prüft, ob eine Phase (elemID) aus dem Projekt hproj gelöscht werden kann, 
    ''' da weder sie selbst betrachtet werden soll, noch all ihre Kinder
    ''' </summary>
    ''' <param name="elemID"></param>
    ''' <param name="hproj"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property isRemovable(ByVal elemID As String, ByVal hproj As clsProjekt, ByVal liste As SortedList(Of String, Boolean)) As Boolean
        Get
            Dim ind As Integer = 1
            Dim hrchynode As clsHierarchyNode = Nothing

            isRemovable = True

            Try

                hrchynode = hproj.hierarchy.nodeItem(elemID)
                If hrchynode.childCount = 0 Then
                    isRemovable = isRemovable And True
                End If
                If hrchynode.childCount > 0 Then
                    For ind = 1 To hrchynode.childCount
                        Dim nodeID As String = hrchynode.getChild(ind)
                        isRemovable = isRemovable And liste.ContainsKey(nodeID)
                        isRemovable = isRemovable And isRemovable(hrchynode.getChild(ind), hproj, liste)
                    Next

                End If

            Catch ex As Exception
                Call MsgBox("Fehler bei der Prüfung, ob das Element elemID= " & elemID & " entfernt werden kann")
                Throw New ArgumentException("Fehler bei der Prüfung, ob das Element elemID entfernt werden kann")
            End Try

        End Get

    End Property

    ''' <summary>
    ''' erzeugt die Projekte, die in der Batch-Datei angegeben sind
    ''' </summary>
    ''' <param name="myCollection"></param>
    ''' <remarks></remarks>
    Public Sub awinImportProjektInventur(ByRef myCollection As Collection)
        Dim zeile As Integer, spalte As Integer
        Dim pName As String = ""
        Dim vorlageName As String = ""
        Dim start As Date, inputStart As Date
        Dim startElem As String = ""
        Dim endElem As String = ""
        Dim ende As Date, inputEnde As Date
        Dim budget As Double
        Dim dauer As Integer = 0
        Dim sfit As Double, risk As Double
        Dim volume As Double, complexity As Double
        Dim description As String = ""
        Dim businessUnit As String = ""
        Dim lastRow As Integer
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

        Dim dauerFaktor As Double = 1.0
        Dim refProj As New clsProjekt

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

        Dim suchstr(13) As String
        suchstr(ptInventurSpalten.Name) = "Name"
        suchstr(ptInventurSpalten.Vorlage) = "Vorlage"
        suchstr(ptInventurSpalten.Start) = "Start-Datum"
        suchstr(ptInventurSpalten.Ende) = "Ende-Datum"
        suchstr(ptInventurSpalten.startElement) = "Bezug Start"
        suchstr(ptInventurSpalten.endElement) = "Bezug Ende"
        suchstr(ptInventurSpalten.Dauer) = "Dauer [Tage]"
        suchstr(ptInventurSpalten.Budget) = "Budget [T€]"
        suchstr(ptInventurSpalten.Risiko) = "Risiko"
        suchstr(ptInventurSpalten.Strategie) = "Strategie"
        suchstr(ptInventurSpalten.Volumen) = "Volumen"
        suchstr(ptInventurSpalten.Komplexitaet) = "Komplexität"
        suchstr(ptInventurSpalten.Businessunit) = "Business Unit"
        suchstr(ptInventurSpalten.Beschreibung) = "Beschreibung"


        Dim inputColumns(11) As Integer



        Try
            Dim activeWSListe As Excel.Worksheet = CType(appInstance.ActiveWorkbook.Worksheets("Liste"), _
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet)
            With activeWSListe

                firstZeile = CType(.Rows(1), Excel.Range)

                ' jetzt werden die Spalten bestimmt 
                Try
                    For i As Integer = 0 To 13
                        inputColumns(i) = firstZeile.Find(What:=suchstr(i)).Column
                    Next
                Catch ex As Exception

                End Try


                lastRow = CType(.Cells(2000, 1), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlUp).Row

                While zeile <= lastRow
                    ok = False
                    Dim sMilestone As clsMeilenstein = Nothing
                    Dim eMilestone As clsMeilenstein = Nothing

                    pName = CStr(CType(.Cells(zeile, spalte), Global.Microsoft.Office.Interop.Excel.Range).Value)
                    vorlageName = CStr(CType(.Cells(zeile, spalte + 1), Global.Microsoft.Office.Interop.Excel.Range).Value)

                    If Projektvorlagen.Liste.ContainsKey(vorlageName) Then

                        vproj = Projektvorlagen.getProject(vorlageName)
                        refProj = New clsProjekt
                        vproj.copyTo(refProj)
                        refProj.startDate = Date.Now

                        Try

                            start = CDate(CType(.Cells(zeile, spalte + 2), Global.Microsoft.Office.Interop.Excel.Range).Value)
                            ende = CDate(CType(.Cells(zeile, spalte + 3), Global.Microsoft.Office.Interop.Excel.Range).Value)
                            startElem = CStr(CType(.Cells(zeile, spalte + 4), Global.Microsoft.Office.Interop.Excel.Range).Value)
                            endElem = CStr(CType(.Cells(zeile, spalte + 5), Global.Microsoft.Office.Interop.Excel.Range).Value)
                            dauer = CInt(CType(.Cells(zeile, spalte + 6), Global.Microsoft.Office.Interop.Excel.Range).Value)
                            budget = CDbl(CType(.Cells(zeile, spalte + 7), Global.Microsoft.Office.Interop.Excel.Range).Value)
                            risk = CDbl(CType(.Cells(zeile, spalte + 8), Global.Microsoft.Office.Interop.Excel.Range).Value)
                            sfit = CDbl(CType(.Cells(zeile, spalte + 9), Global.Microsoft.Office.Interop.Excel.Range).Value)
                            volume = CDbl(CType(.Cells(zeile, spalte + 10), Global.Microsoft.Office.Interop.Excel.Range).Value)
                            complexity = CDbl(CType(.Cells(zeile, spalte + 11), Global.Microsoft.Office.Interop.Excel.Range).Value)
                            businessUnit = CStr(CType(.Cells(zeile, spalte + 12), Global.Microsoft.Office.Interop.Excel.Range).Value)
                            description = CStr(CType(.Cells(zeile, spalte + 13), Global.Microsoft.Office.Interop.Excel.Range).Value)
                            'vglName = pName.Trim & "#" & ""
                            vglName = calcProjektKey(pName.Trim, scenarioName)
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
                            CType(.Cells(zeile, spalte + 1), Global.Microsoft.Office.Interop.Excel.Range).Value = ".?."
                            ok = False
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

                        End Try


                    Else
                        CType(.Cells(zeile, spalte + 1), Global.Microsoft.Office.Interop.Excel.Range).Value = ".?."
                        ok = False
                    End If

                    ' jetzt die Aktion durchführen, wenn alles ok 
                    If ok Then
                        If AlleProjekte.Containskey(vglName) Then
                            ' nichts tun ...
                            Call MsgBox("Projekt aus Batch-Liste existiert bereits - keine Neuanlage")
                        Else
                            'Projekt anlegen ,Verschiebung um 
                            hproj = New clsProjekt(start, start.AddMonths(-1), start.AddMonths(1))

                            Dim variantName As String
                            If scenarioName = "Init" Then
                                variantName = ""
                            Else
                                variantName = scenarioName
                            End If
                            Call erstelleInventurProjekt(hproj, pName, vorlageName, variantName, _
                                                         start, ende, budget, zeile, sfit, risk, _
                                                         volume, complexity, businessUnit, description)

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

                            If Not hproj Is Nothing Then
                                Try
                                    ImportProjekte.Add(calcProjektKey(hproj), hproj)
                                    myCollection.Add(calcProjektKey(hproj))
                                Catch ex As Exception

                                End Try

                            End If

                        End If
                    End If



                    zeile = zeile + 1

                End While





            End With
        Catch ex As Exception
            Throw New Exception("Fehler in Szenario-Datei")
        End Try

        ' jetzt noch ein Szenario anlegen, wenn ImportProjekte was enthält 
        If ImportProjekte.Count > 0 Then
            Call storeSessionConstellation(ShowProjekte, scenarioName, ImportProjekte)
        End If

    End Sub

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
                            'Projekt anlegen ,Verschiebung um 
                            hproj = New clsProjekt(start, start.AddMonths(-1), start.AddMonths(1))

                            Call erstelleInventurProjekt(hproj, pName, vorlagenName, scenarioName, _
                                                         start, ende, budget, zeile, sfit, risk, _
                                                         volume, complexity, businessUnit, description)
                            projectStartDate = start
                            projectEndDate = ende

                        End If
                    End If

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
                            ImportProjekte.Add(calcProjektKey(hproj), hproj)
                            myCollection.Add(calcProjektKey(hproj))
                        Catch ex As Exception

                        End Try

                    End If


                    zeile = zeile + 1

                End While





            End With
        Catch ex As Exception
            Throw New Exception("Fehler in Datei Module Import ...")
        End Try


        ' jetzt noch ein Szenario anlegen, wenn ImportProjekte was enthält 
        If ImportProjekte.Count > 0 Then
            Call storeSessionConstellation(ShowProjekte, scenarioName, ImportProjekte)
        End If

        currentConstellation = scenarioName

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
                        Call splitHryFullnameTo2(.referenceName, phaseName, breadCrumb)
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
                        Call splitHryFullnameTo2(.referenceName, milestoneName, breadCrumb)
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
                            Call splitHryFullnameTo2(currentRule.referenceName, phaseName, breadCrumb)
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
                            Call splitHryFullnameTo2(currentRule.referenceName, milestoneName, breadCrumb)
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
                                .deliverables = currentElem.deliverables
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
                hproj.name = CType(.Range("Projekt_Name").Value, String)
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
                                    .Offset = 0

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
                                        If RoleDefinitions.Contains(hname) Then
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

                                                crole = New clsRolle(ende - anfang + 1)
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

                                        ElseIf CostDefinitions.Contains(hname) Then

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

                                                ccost = New clsKostenart(ende - anfang + 1)
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
        Dim hwert As Integer
        Dim ProjektdauerIndays As Integer = 0
        Dim endedateProjekt As Date


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
                    hproj.name = CType(.Range("Projekt_Name").Value, String)
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

                                    objectName = CType(CType(.Cells(zeile, columnOffset), Excel.Range).Value, String).Trim

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
                                                    Call logfileSchreiben(("Fehler, Lesen Termine: unzulässige Angaben für Offset (>=0) und Dauer (>=1): " & _
                                                                        "Offset= " & offset.ToString & _
                                                                        ", Duration= " & duration.ToString), hproj.name, anzFehler)
                                                    Throw New Exception("Fehler, Lesen Termine: unzulässige Angaben für Offset (>=0) und Dauer (>=1): " & _
                                                                        "Offset= " & offset.ToString & _
                                                                        ", Duration= " & duration.ToString)
                                                End If
                                            End If

                                            ' für die rootPhase muss gelten: offset = startoffset = 0 und duration = ProjektdauerIndays
                                            If duration <> ProjektdauerIndays Or offset <> 0 Then

                                                Call logfileSchreiben(("Fehler, Lesen Termine: unzulässige Angaben für Offset und Dauer: der ProjektPhase " & _
                                                                        "Offset= " & offset.ToString & _
                                                                        ", Duration=" & duration.ToString & _
                                                                        ", ProjektDauer=" & ProjektdauerIndays.ToString), hproj.name, anzFehler)
                                                Throw New Exception("Fehler, Lesen Termine: unzulässige Angaben für Offset und Dauer: der ProjektPhase " & _
                                                                        "Offset= " & offset.ToString & _
                                                                        ", Duration=" & duration.ToString & _
                                                                        ", ProjektDauer=" & ProjektdauerIndays.ToString)
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
                                                Call logfileSchreiben(("Fehler, Lesen Termine: unzulässige Angaben für Offset und Dauer: " & _
                                                                    offset.ToString & ", " & duration.ToString), hproj.name, anzFehler)
                                                Throw New Exception("Fehler, Lesen Termine: unzulässige Angaben für Offset und Dauer: " & _
                                                                    offset.ToString & ", " & duration.ToString)
                                            End If
                                        End If

                                        isPhase = True
                                        isMeilenstein = False

                                    End If

                                    ' eingelesener String objectname ist eine Phase

                                    If isPhase Then

                                        cphase = New clsPhase(parent:=hproj)

                                        If PhaseDefinitions.Contains(objectName) Or awinSettings.alwaysAcceptTemplateNames _
                                            Or awinSettings.addMissingPhaseMilestoneDef Then

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
                                        If MilestoneDefinitions.Contains(objectName) Or awinSettings.alwaysAcceptTemplateNames _
                                            Or awinSettings.addMissingPhaseMilestoneDef Then

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
                                            bewertungsAmpel = CType(CType(.Cells(zeile, columnOffset + 4), Excel.Range).Value, Integer)
                                            explanation = CType(CType(.Cells(zeile, columnOffset + 5), Excel.Range).Value, String)

                                            ' Ergänzung tk 2.11 deliverables ergänzt 
                                            deliverables = CType(CType(.Cells(zeile, columnOffset + 6), Excel.Range).Value, String)

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
                                                .deliverables = deliverables
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
                            ressOff = 1
                            ressSumOffset = -1              ' keine Summe vorhanden
                            Call logfileSchreiben("alte Version des ProjektSteckbriefes: ohne 'Summe'", hproj.name, anzFehler)
                        Else
                            ressOff = gefundenRange.Column - rng.Column - 1
                            ressSumOffset = gefundenRange.Column - rng.Column - 2
                            Call logfileSchreiben("neue Version des ProjektSteckbriefes: mit 'Summe'", hproj.name, anzFehler)

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

                                    If Not fertig Then

                                        lastelemID = cphase.nameID
                                        lastphase = cphase
                                        lastlevel = cphaseLevel
                                    End If

                                    ' in cphase wird die Phase mit Namen phaseName, bereits über Termine in der Hierarchie des Projekts eingetragen
                                    ' gespeichert
                                    If phaseName = hproj.name Or phaseName = elemNameOfElemID(rootPhaseName) Then

                                        cphase = hproj.getPhaseByID(rootPhaseName)


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
                                                Call splitHryFullnameTo2(breadcrumb, hhstr, breadcrumb)
                                                lastlevel = lastlevel - 1
                                            End While

                                        End If

                                        ' Prüfung, ob die Phase phaseName in der bereits aus Termine bestehenden Hierarchie mit dem gleiche breadcrumb existiert, sonst Fehler

                                        If Not hproj.hierarchy.containsPhase(phaseName, breadcrumb) Then

                                            Call logfileSchreiben("Fehler beim Lesen Ressourcen: bei Phase '" & phaseName & "#" & breadcrumb & "'", hproj.name, anzFehler)
                                            Throw New ArgumentException("Fehler beim Lesen Ressourcen: bei Phase '" & phaseName & "#" & breadcrumb & "'")
                                        Else

                                            Dim phaseIndex() As Integer = hproj.hierarchy.getPhaseIndices(phaseName, breadcrumb)

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


                                    If Not (anfang = cphase.relStart And ende = cphase.relEnde) Then

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
                                            If RoleDefinitions.Contains(hname) Then
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

                                                    crole = New clsRolle(ende - anfang + 1)
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

                                            ElseIf CostDefinitions.Contains(hname) Then

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

                                                    ccost = New clsKostenart(ende - anfang + 1)
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

    ''' <summary>
    ''' liest einen ProjektSteckbrief mit Hierarchie ein, vor mit PT113 die Spalte Summe hinzugefügt wurde
    ''' </summary>
    ''' <param name="hprojekt"></param>
    ''' <param name="hprojTemp"></param>
    ''' <param name="isTemplate"></param>
    ''' <param name="importDatum"></param>
    ''' <remarks></remarks>
    Public Sub awinImportProjectmitHrchy_beforePT113(ByRef hprojekt As clsProjekt, ByRef hprojTemp As clsProjektvorlage, ByVal isTemplate As Boolean, ByVal importDatum As Date)

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
                hproj.name = CType(.Range("Projekt_Name").Value, String)
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
                    Dim lastphase As clsPhase
                    Dim lasthrchyNode As clsHierarchyNode
                    Dim lastelemID As String = ""
                    Dim ccost As clsKostenart
                    Dim phaseName As String = ""
                    Dim aktLevel As Integer = 0   'speichert den Level direkt nach dem Lesen der Phase
                    Dim cphaseLevel As Integer = 0 'speichert den Level der momentan in cphase gespeicherten Phase
                    Dim lastlevel As Integer = 0  'speichert den Level des vorausgehenden elements

                    Dim anfang As Integer, ende As Integer  ', projDauer As Integer

                    Dim farbeAktuell As Object
                    Dim r As Integer, k As Integer


                    .Unprotect(Password:="x")       ' Blattschutz aufheben


                    Dim tmpws As Excel.Range = CType(wsRessourcen.Range("Phasen_des_Projekts"), Excel.Range)

                    rng = .Range("Phasen_des_Projekts")

                    Dim hstr As String = CStr(CType(.Range("Phasen_des_Projekts").Cells(1), Excel.Range).Value)
                    hstr = elemNameOfElemID(rootPhaseName)

                    If CStr(CType(.Range("Phasen_des_Projekts").Cells(1), Excel.Range).Value) <> elemNameOfElemID(rootPhaseName) Then


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
                        hstr = CStr(zelle.Value)
                        Dim x As Integer = CInt(zelle.IndentLevel)
                        If x Mod einrückTiefe <> 0 Then
                            Throw New ArgumentException("die Einrückung ist keine durch '" & CStr(einrückTiefe) & "' teilbare Zahl")
                        End If
                        aktLevel = CInt(x / einrückTiefe)

                        If Len(CType(zelle.Value, String)) > 0 Then
                            phaseName = CType(zelle.Value, String).Trim
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
                                    ' '' ''  hproj.AddPhase(cphase)

                                    Dim hrchynode As New clsHierarchyNode
                                    hrchynode.elemName = cphase.name


                                    If cphaseLevel = 0 Then
                                        hrchynode.parentNodeKey = ""

                                    ElseIf cphaseLevel = 1 Then
                                        hrchynode.parentNodeKey = rootPhaseName

                                    ElseIf cphaseLevel - lastlevel = 1 Then
                                        hrchynode.parentNodeKey = lastelemID

                                    ElseIf cphaseLevel - lastlevel = 0 Then
                                        hrchynode.parentNodeKey = hproj.hierarchy.getParentIDOfID(lastelemID)

                                    ElseIf lastlevel - cphaseLevel >= 1 Then
                                        Dim hilfselemID As String = lastelemID
                                        For l As Integer = 1 To lastlevel - cphaseLevel
                                            hilfselemID = hproj.hierarchy.getParentIDOfID(hilfselemID)
                                        Next l
                                        hrchynode.parentNodeKey = hproj.hierarchy.getParentIDOfID(hilfselemID)
                                    Else
                                        Throw New ArgumentException("Fehler beim Import! Hierarchie kann nicht richtig aufgebaut werden")
                                    End If

                                    hproj.AddPhase(cphase, parentID:=hrchynode.parentNodeKey)

                                    ' '' ''hproj.hierarchy.addNode(hrchynode, cphase.nameID)
                                    hrchynode.indexOfElem = hproj.AllPhases.Count
                                    ' merken von letzem Element (Knoten,Phase,Meilenstein)
                                    lasthrchyNode = hrchynode
                                    lastelemID = cphase.nameID
                                    lastphase = cphase
                                    lastlevel = cphaseLevel
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
                                    If phaseName = hproj.name Or phaseName = elemNameOfElemID(rootPhaseName) Then
                                        .nameID = rootPhaseName
                                        ' nichts tun, die erste Phase hat dann schon ihren richtigen Namen 
                                    Else
                                        .nameID = hproj.hierarchy.findUniqueElemKey(phaseName, False)
                                    End If
                                    cphaseLevel = aktLevel

                                    ' Änderung 28.11.13: jetzt wird die Phasen Länge exakt bestimmt , über startoffset in Tagen und dauerinDays als Länge
                                    Dim startOffset As Long
                                    Dim dauerIndays As Long
                                    startOffset = DateDiff(DateInterval.Day, hproj.startDate, hproj.startDate.AddMonths(anfang - 1))
                                    dauerIndays = calcDauerIndays(hproj.startDate.AddDays(startOffset), ende - anfang + 1, True)

                                    .changeStartandDauer(startOffset, dauerIndays)
                                    .Offset = 0

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
                                        If RoleDefinitions.Contains(hname) Then
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

                                                crole = New clsRolle(ende - anfang + 1)
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

                                        ElseIf CostDefinitions.Contains(hname) Then

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

                                                ccost = New clsKostenart(ende - anfang + 1)
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
                                            '''''hproj.AddPhase(cphase)

                                            Dim hrchynode As New clsHierarchyNode
                                            hrchynode.elemName = cphase.name


                                            If cphaseLevel = 0 Then
                                                hrchynode.parentNodeKey = ""

                                            ElseIf cphaseLevel = 1 Then
                                                hrchynode.parentNodeKey = rootPhaseName

                                            ElseIf cphaseLevel - lastlevel = 1 Then
                                                hrchynode.parentNodeKey = lastelemID

                                            ElseIf cphaseLevel - lastlevel = 0 Then
                                                hrchynode.parentNodeKey = hproj.hierarchy.getParentIDOfID(lastelemID)

                                            ElseIf lastlevel - cphaseLevel >= 1 Then
                                                Dim hilfselemID As String = lastelemID
                                                For l As Integer = 1 To lastlevel - cphaseLevel
                                                    hilfselemID = hproj.hierarchy.getParentIDOfID(hilfselemID)
                                                Next l
                                                hrchynode.parentNodeKey = hproj.hierarchy.getParentIDOfID(hilfselemID)
                                            Else
                                                Throw New ArgumentException("Fehler beim Import! Hierarchie kann nicht richtig aufgebaut werden")
                                            End If

                                            hproj.AddPhase(cphase, parentID:=hrchynode.parentNodeKey)

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
                        Dim phaseNameID As String
                        Dim milestoneName As String
                        Dim milestoneDate As Date
                        Dim resultVerantwortlich As String = ""
                        Dim bewertungsAmpel As Integer
                        Dim explanation As String
                        Dim deliverables As String
                        Dim bewertungsdatum As Date = importDatum
                        Dim Nummer As String
                        Dim tbl As Excel.Range
                        Dim rowOffset As Integer
                        Dim columnOffset As Integer


                        .Unprotect(Password:="x")       ' Blattschutz aufheben

                        tbl = .Range("ErgebnTabelle")
                        rowOffset = tbl.Row
                        columnOffset = tbl.Column

                        lastrow = CInt(CType(.Cells(2000, columnOffset), Excel.Range).End(XlDirection.xlUp).Row)

                        ' ur: 12.05.2015: hier wurde die Sortierung der ErgebnTabelle entfernt

                        Dim cphase As clsPhase = Nothing
                        Dim breadCrumb As String = ""
                        Dim lastLevel As Integer = 0

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

                            '' ''If zeile = 68 Then
                            '' ''    zeile = 68
                            '' ''End If
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
                                    ' String aus erster Spalte der Tabelle lesen

                                    objectName = CType(CType(.Cells(zeile, columnOffset), Excel.Range).Value, String).Trim

                                    ' Level abfragen

                                    Dim x As Integer = CInt(CType(.Cells(zeile, columnOffset), Excel.Range).IndentLevel)
                                    If x Mod einrückTiefe <> 0 Then
                                        Throw New ArgumentException("die Einrückung ist keine durch '" & CStr(einrückTiefe) & "' teilbare Zahl")
                                    End If
                                    aktLevel = CInt(x / einrückTiefe)


                                Catch ex As Exception
                                    objectName = Nothing
                                    Throw New Exception("In Tabelle 'Termine' ist der PhasenName nicht angegeben ")
                                    Exit For ' Ende der For-Schleife, wenn keine laufende Nummer mehr existiert
                                End Try


                                Try
                                    startDate = CDate(CType(.Cells(zeile, columnOffset + 2), Excel.Range).Value)
                                Catch ex As Exception
                                    startDate = Date.MinValue
                                End Try


                                If objectName = elemNameOfElemID(rootPhaseName) Or PhaseDefinitions.Contains(objectName) Then

                                    isPhase = True
                                    isMeilenstein = False


                                ElseIf startDate <> Date.MinValue Then
                                    Throw New ArgumentException("'" & objectName & "' ist eine Phase, die nicht im CustomizationFile definiert ist. Bitte korrigieren Sie dies!")
                                Else

                                    isPhase = False
                                    isMeilenstein = True

                                End If


                                '  ur: 12.05.2015: Änderung, damit Meilensteine, die den gleichen Namen haben wie Phasen, trotzdem als Meilensteine erkannt werden.
                                '                 gilt aktuell aber nur für den BMW-Import
                                If awinSettings.importTyp = 2 Then
                                    If PhaseDefinitions.Contains(objectName) _
                                        And startDate = Date.MinValue Then

                                        isPhase = False
                                        isMeilenstein = True
                                    End If
                                End If

                                Try
                                    endeDate = CDate(CType(.Cells(zeile, columnOffset + 3), Excel.Range).Value)
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

                                        'ur: 1.6.2015   Meilenstein hat den Namen einer Phase
                                        If PhaseDefinitions.Contains(objectName) _
                                            And startDate = Date.MinValue Then

                                            isPhase = False
                                            isMeilenstein = True
                                        End If

                                        'ur:12.05.2015:
                                        ' '' '' ''If IsNothing(cphase) Then
                                        ' '' '' ''    If hproj.AllPhases.Count > 0 Then
                                        ' '' '' ''        cphase = hproj.getPhase(1)
                                        ' '' '' ''    Else
                                        ' '' '' ''        ' Erzeuge ProjektPhase mit Länge des Projekts

                                        ' '' '' ''    End If

                                        ' '' '' ''End If
                                    End If


                                    'isPhase = False

                                Else
                                    'objectName ist eine Phase
                                    'isPhase = True

                                    ' ist der Phasen Name in der Liste der definitionen überhaupt bekannt ? 
                                    If Not PhaseDefinitions.Contains(objectName) Then

                                        ' jetzt noch prüfen, ob es sich um die Phase (1) handelt, dann kann sie ja nicht in der PhaseDefinitions enthalten sein  ..
                                        If elemNameOfElemID(rootPhaseName) = objectName Or hproj.name = objectName Then
                                            ' alles ok
                                        Else
                                            Throw New Exception("Phase '" & objectName & "' ist nicht definiert!" & vbLf &
                                                           "Bitte löschen Sie diese Phase aus '" & hproj.name & "'.xlsx, Tabellenblatt 'Termine'")

                                        End If

                                    End If

                                    ' an dieser stelle ist sichergestellt, daß der Phasen Name bekannt ist
                                    ' Prüfen, ob diese Phase bereits in hproj über das ressourcen Sheet angelegt wurde 
                                    ' tk: dieser Befehl holt jetzt die erste Phase mit deisem NAmen, berücksichtigt aber noch nicht die Position ind er Hierarchie; 
                                    ' das muss noch ergänzt werden 
                                    If hproj.name = objectName Or elemNameOfElemID(rootPhaseName) = objectName Then
                                        cphase = hproj.getPhaseByID(rootPhaseName)
                                        breadCrumb = ""
                                    Else

                                        If aktLevel > lastLevel Then

                                            If breadCrumb = "" Then
                                                breadCrumb = "."
                                            Else
                                                breadCrumb = breadCrumb & "#" & cphase.name
                                            End If

                                        ElseIf aktLevel = lastLevel Then
                                            ' aktlevel = lastlevel: also nicht tun
                                        Else

                                            While aktLevel < lastLevel
                                                Dim hstr As String = ""
                                                Call splitHryFullnameTo2(breadCrumb, hstr, breadCrumb)
                                                lastLevel = lastLevel - 1
                                            End While

                                        End If
                                        cphase = hproj.getPhase(objectName, breadCrumb)

                                        If IsNothing(cphase) Then
                                            If aktLevel <> hproj.hierarchy.getIndentLevel(cphase.nameID) Then

                                                ' ur: 11.05.2015: fehler, wenn die Phase nicht exisitiert, 
                                                '               nicht erzeugen
                                                ' Phase existiert nicht mit dem gleichen Breadcrumb
                                                Throw New ArgumentException("Die Phase '" & objectName & "' existiert nicht in dieser angegebenen Stufe" & vbLf & _
                                                                            "Bitte korrigieren Sie die Importdatei!" & "BreadCrumb = " & breadCrumb)

                                            End If



                                        End If

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
                                        ' '' '' ''If Not cphaseExisted Then
                                        ' '' '' ''    ' ur: 11.05.2015: parentID bestimmen fehlt hier noch
                                        ' '' '' ''    hproj.AddPhase(cphase, parentID:=rootPhaseName)
                                        ' '' '' ''End If


                                    Catch ex As Exception
                                        Throw New Exception(ex.Message)
                                    End Try

                                Else

                                    If aktLevel > lastLevel Then

                                        If breadCrumb = "" Then
                                            breadCrumb = "."
                                        Else
                                            breadCrumb = breadCrumb & "#" & cphase.name
                                        End If

                                    ElseIf aktLevel = lastLevel Then
                                        ' aktlevel = lastlevel: also nicht tun
                                    Else

                                        While aktLevel < lastLevel
                                            Dim hstr As String = ""
                                            Call splitHryFullnameTo2(breadCrumb, hstr, breadCrumb)
                                            lastLevel = lastLevel - 1
                                        End While

                                    End If

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
                                    Try
                                        bewertungsAmpel = CType(CType(.Cells(zeile, columnOffset + 4), Excel.Range).Value, Integer)
                                    Catch ex As Exception
                                        bewertungsAmpel = 0
                                    End Try

                                    explanation = CType(CType(.Cells(zeile, columnOffset + 5), Excel.Range).Value, String)

                                    ' Ergänzung tk 2.11 deliverables ergänzt 
                                    deliverables = CType(CType(.Cells(zeile, columnOffset + 6), Excel.Range).Value, String)


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
                                        .deliverables = deliverables
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



                                End If

                            Catch ex As Exception
                                If zeile <> lastrow Then
                                    ' beim lesen des ImportFiles ist ein Fehler aufgetreten
                                    Throw New Exception(ex.Message)
                                End If
                                ' letzte belegte Zeile wurde bereits bearbeitet.
                                zeile = lastrow + 1 ' erzwingt das Ende der For - Schleife
                                Nummer = Nothing


                            End Try

                            lastLevel = aktLevel                ' indentlevel merken
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
            projVorlage.hierarchy = hproj.hierarchy
            hprojTemp = projVorlage

        Else
            hprojekt = hproj
        End If

    End Sub

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
                lastConstellation = projectConstellations.getConstellation("Last")
            Catch ex As Exception
                'Call MsgBox("in awinProjekteInitialLaden Fehler ...")
            End Try

            ' jetzt Showprojekte aufbauen - und zwar so, dass Konstellation <Last> wiederhergestellt wird
            For Each kvp As KeyValuePair(Of String, clsConstellationItem) In lastConstellation.Liste

                Try
                    hproj = AlleProjekte.getProject(kvp.Key)
                    hproj.startDate = kvp.Value.Start
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
                        Call awinCreateBudgetWerte(kvp.Value)

                        AlleProjekte.Add(kvp.Key, kvp.Value)
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
                    Call awinCreateBudgetWerte(kvp.Value)
                    AlleProjekte.Add(kvp.Key, kvp.Value)

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
    ''' <summary>
    ''' lädt ein bestimmtes Portfolio von der Datenbank und zeigt es  
    ''' in der Projekttafel an.
    ''' 
    ''' </summary>
    ''' <param name="constellationName">
    ''' Name, unter dem das Portfolio in der Datenbank gespeichert wurde 
    ''' </param>
    ''' <remarks></remarks>
    ''' 
    Public Sub awinLoadConstellation(ByVal constellationName As String, ByRef successMessage As String)

        Dim activeConstellation As New clsConstellation
        Dim hproj As New clsProjekt
        Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
        Dim anzErrDB As Integer = 0
        Dim loadErrorMessage As String = " * Projekte, die nicht in der DB '" & awinSettings.databaseName & "' existieren:"
        Dim loadDateMessage As String = " * Das Datum kann nicht angepasst werden kann." & vbLf & _
                                        "   Das Projekt wurde bereits beauftragt."

        ' prüfen, ob diese Constellation bereits existiert ..
        Try
            activeConstellation = projectConstellations.getConstellation(constellationName)
        Catch ex As Exception
            Call MsgBox(" Projekt-Konstellation " & constellationName & " existiert nicht ")
            Exit Sub
        End Try

        ' die aktuelle Konstellation in "Last" speichern 
        Call storeSessionConstellation(ShowProjekte, "Last")

        ShowProjekte.Clear()

        ' 3.11.14 : das hier darf nicht gemacht werden ! denn dann nimmt er ausschließlich alle Daten aus der Datenbank !
        ' AlleProjekte.liste.Clear()


        ' jetzt werden die Start-Values entsprechend gesetzt ..

        For Each kvp As KeyValuePair(Of String, clsConstellationItem) In activeConstellation.Liste

            If AlleProjekte.Containskey(kvp.Key) Then
                ' Projekt ist bereits im Hauptspeicher geladen
                hproj = AlleProjekte.getProject(kvp.Key)
            Else
                If request.pingMongoDb() Then

                    If request.projectNameAlreadyExists(kvp.Value.projectName, kvp.Value.variantName) Then

                        ' Projekt ist noch nicht im Hauptspeicher geladen, es muss aus der Datenbank geholt werden.
                        hproj = request.retrieveOneProjectfromDB(kvp.Value.projectName, kvp.Value.variantName)

                        ' Projekt muss nun in die Liste der geladenen Projekte eingetragen werden
                        AlleProjekte.Add(kvp.Key, hproj)
                    Else
                        anzErrDB = anzErrDB + 1
                        If anzErrDB = 1 Then
                            successMessage = successMessage & loadErrorMessage & vbLf & _
                                                   "        " & kvp.Value.projectName
                        Else
                            successMessage = successMessage & vbLf & _
                                                   "        " & kvp.Value.projectName
                        End If

                        'Call MsgBox("Projekt '" & kvp.Value.projectName & "'konnte nicht geladen werden")
                        'Throw New ArgumentException("Projekt '" & kvp.Value.projectName & "'konnte nicht geladen werden")
                    End If
                Else
                    Throw New ArgumentException("Datenbank-Verbindung ist unterbrochen!" & vbLf & "Projekt '" & kvp.Value.projectName & "'konnte nicht geladen werden")
                End If
            End If
            If hproj.name = kvp.Value.projectName Then

                With hproj

                    ' Änderung THOMAS Start 
                    If .Status = ProjektStatus(0) Then
                        .startDate = kvp.Value.Start
                    ElseIf .startDate <> kvp.Value.Start Then
                        ' wenn das Datum nicht angepasst werden kann, weil das Projekt bereits beauftragt wurde  
                        successMessage = successMessage & vbLf & vbLf & _
                                            loadDateMessage & vbLf & _
                                            "        " & hproj.name & ": " & kvp.Value.Start.ToShortDateString
                    End If
                    ' Änderung THOMAS Ende 

                    .StartOffset = 0
                    .tfZeile = kvp.Value.zeile
                End With

                If kvp.Value.show Then

                    Try

                        ShowProjekte.Add(hproj)

                    Catch ex1 As Exception
                        Call MsgBox("Fehler in awinLoadConstellation aufgetreten: " & ex1.Message)
                    End Try

                End If

            End If

        Next


    End Sub

    ''' <summary>
    ''' fügt die in der Konstellation aufgeführten Projekte hinzu; 
    ''' wenn Sie bereits geladen sind, wird nachgesehen, ob die richtige Variante aktiviert ist 
    ''' ggf. wird diese Variante dann aktiviert 
    ''' </summary>
    ''' <param name="constellationName"></param>
    ''' <param name="successMessage"></param>
    ''' <remarks></remarks>
    Public Sub awinAddConstellation(ByVal constellationName As String, ByRef successMessage As String)

        Dim activeConstellation As New clsConstellation
        Dim hproj As New clsProjekt
        Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
        Dim anzErrDB As Integer = 0
        Dim loadErrorMessage As String = " * Projekte, die nicht in der DB '" & awinSettings.databaseName & "' existieren:"
        Dim loadDateMessage As String = " * Das Datum kann nicht angepasst werden kann." & vbLf & _
                                        "   Das Projekt wurde bereits beauftragt."
        Dim tryZeile As Integer

        ' ab diesem Wert soll neu gezeichnet werden 
        Dim startOfFreeRows As Integer = projectboardShapes.getMaxZeile

        ' prüfen, ob diese Constellation bereits existiert ..
        Try
            activeConstellation = projectConstellations.getConstellation(constellationName)
        Catch ex As Exception
            Call MsgBox(" Projekt-Konstellation " & constellationName & " existiert nicht ")
            Exit Sub
        End Try

        ' die aktuelle Konstellation in "Last" speichern 
        Call storeSessionConstellation(ShowProjekte, "Last")

        ' jetzt werden die einzelnen Projekte dazugeholt 

        For Each kvp As KeyValuePair(Of String, clsConstellationItem) In activeConstellation.Liste

            If AlleProjekte.Containskey(kvp.Key) Then
                ' Projekt ist bereits im Hauptspeicher geladen
                hproj = AlleProjekte.getProject(kvp.Key)

                ' wenn es bereits in Showprojekte ist , gar nichts machen
                If ShowProjekte.contains(hproj.name) Then
                    tryZeile = ShowProjekte.getProject(hproj.name).tfZeile
                Else
                    tryZeile = kvp.Value.zeile + startOfFreeRows - 1
                End If

                ' jetzt die Variante aktivieren 
                Call replaceProjectVariant(hproj.name, hproj.variantName, False, False, tryZeile)

            Else
                If request.pingMongoDb() Then

                    If request.projectNameAlreadyExists(kvp.Value.projectName, kvp.Value.variantName) Then

                        ' Projekt ist noch nicht im Hauptspeicher geladen, es muss aus der Datenbank geholt werden.
                        hproj = request.retrieveOneProjectfromDB(kvp.Value.projectName, kvp.Value.variantName)

                        ' Projekt muss nun in die Liste der geladenen Projekte eingetragen werden
                        AlleProjekte.Add(kvp.Key, hproj)
                        ' jetzt die Variante aktivieren 
                        tryZeile = kvp.Value.zeile + startOfFreeRows - 1
                        Call replaceProjectVariant(hproj.name, hproj.variantName, False, False, tryZeile)

                    Else
                        anzErrDB = anzErrDB + 1
                        If anzErrDB = 1 Then
                            successMessage = successMessage & loadErrorMessage & vbLf & _
                                                   "        " & kvp.Value.projectName
                        Else
                            successMessage = successMessage & vbLf & _
                                                   "        " & kvp.Value.projectName
                        End If

                        'Call MsgBox("Projekt '" & kvp.Value.projectName & "'konnte nicht geladen werden")
                        'Throw New ArgumentException("Projekt '" & kvp.Value.projectName & "'konnte nicht geladen werden")
                    End If
                Else
                    Throw New ArgumentException("Datenbank-Verbindung ist unterbrochen!" & vbLf & "Projekt '" & kvp.Value.projectName & "'konnte nicht geladen werden")
                End If
            End If

        Next



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
        Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)

        ' prüfen, ob diese Constellation überhaupt existiert ..
        Try
            activeConstellation = projectConstellations.getConstellation(constellationName)
        Catch ex As Exception
            Call MsgBox(" Projekt-Konstellation " & constellationName & " existiert nicht ")
            Exit Sub
        End Try

        If deleteDB Then
            If request.pingMongoDb() Then

                ' Konstellation muss aus der Datenbank gelöscht werden.

                returnValue = request.removeConstellationFromDB(activeConstellation)
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
    Public Sub loadProjectfromDB(ByVal pName As String, vName As String, ByVal show As Boolean)

        Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
        Dim hproj As clsProjekt
        Dim key As String = calcProjektKey(pName, vName)

        ' ab diesem Wert soll neu gezeichnet werden 
        Dim freieZeile As Integer = projectboardShapes.getMaxZeile

        hproj = request.retrieveOneProjectfromDB(pName, vName)

        ' prüfen, ob AlleProjekte das Projekt bereits enthält 
        ' danach ist sichergestellt, daß AlleProjekte das Projekt bereit enthält 
        If AlleProjekte.Containskey(key) Then
            AlleProjekte.Remove(key)
        End If

        AlleProjekte.Add(key, hproj)

        If show Then
            ' prüfen, ob es bereits in der Showprojekt enthalten ist
            ' diese Prüfung und die entsprechenden Aktionen erfolgen im 
            ' replaceProjectVariant

            Call replaceProjectVariant(pName, vName, False, True, freieZeile)

        End If



    End Sub

    ''' <summary>
    ''' löscht in der Datenbank alle Timestamps der Projekt-Variante pname, variantname
    ''' die Timestamps werden zudem alle im Papierkorb gesichert 
    ''' </summary>
    ''' <param name="pname">Projektname</param>
    ''' <param name="variantName">Variantenname</param>
    ''' <remarks></remarks>
    Public Sub deleteCompleteProjectVariant(ByVal pname As String, ByVal variantName As String, ByVal kennung As Integer)


        If kennung = PTTvActions.delFromDB Then

            Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
            Dim requestTrash As New Request(awinSettings.databaseURL, awinSettings.databaseName & "Trash", dbUsername, dbPasswort)

            If Not projekthistorie Is Nothing Then
                projekthistorie.clear() ' alte Historie löschen
            End If

            projekthistorie.liste = request.retrieveProjectHistoryFromDB _
                                    (projectname:=pname, variantName:=variantName, _
                                     storedEarliest:=Date.MinValue, storedLatest:=Date.Now)

            ' Speichern im Papierkorb 
            For Each kvp As KeyValuePair(Of Date, clsProjekt) In projekthistorie.liste
                If requestTrash.storeProjectToDB(kvp.Value) Then
                Else
                    ' es ging etwas schief


                    Call MsgBox("Fehler beim Speichern im Papierkorb:" & vbLf & _
                                kvp.Value.name & ", " & kvp.Value.timeStamp.ToShortDateString)
                End If
            Next

            ' jetzt alle Timestamps in der Datenbank löschen 
            Try
                If request.deleteProjectHistoryFromDB(projectname:=pname, variantName:=variantName, _
                                                  storedEarliest:=projekthistorie.First.timeStamp, _
                                                  storedLatest:=projekthistorie.Last.timeStamp) Then

                Else
                    Call MsgBox("Fehler beim Löschen von " & pname & ", " & variantName)
                End If
            Catch ex As Exception

            End Try
            


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

            If IsNothing(hproj) Or hproj.variantName <> variantName Then
                Dim key As String = calcProjektKey(pname, variantName)
                AlleProjekte.Remove(key)

            Else
                If variantName = "" Then

                    Call MsgBox("die Basis Variante kann nicht gelöscht werden")

                ElseIf hproj.variantName = variantName Then
                    ' es wird die Stand-Variante aktiviert 
                    Dim stdProj As clsProjekt
                    Dim stdkey As String = calcProjektKey(pname, "")

                    Dim key As String = calcProjektKey(pname, variantName)

                    Try
                        stdProj = AlleProjekte.getProject(stdkey)

                        ' jetzt muss die bisherige Variante aus Showprojekte rausgenommen werden ..
                        ShowProjekte.Remove(hproj.name)

                        ' die gewählte Variante wird rausgenommen
                        AlleProjekte.Remove(key)

                        ' die Standard Variante wird aufgenommen
                        ShowProjekte.Add(stdProj)

                        Call clearProjektinPlantafel(pname)

                        ' neu zeichnen des Projekts 
                        Dim tmpCollection As New Collection
                        Call ZeichneProjektinPlanTafel(tmpCollection, stdProj.name, hproj.tfZeile, tmpCollection, tmpCollection)


                    Catch ex As Exception

                    End Try

                Else
                    Call MsgBox("Fehler beim Löschen der Variante")
                End If
            End If



        End If


    End Sub




    ''' <summary>
    ''' löscht den angegebenen timestamp von pname#variantname aus der Datenbank
    ''' speichert den timestamp im Papierkorb
    ''' </summary>
    ''' <param name="pname"></param>
    ''' <param name="variantName"></param>
    ''' <param name="timeStamp"></param>
    ''' <param name="first"></param>
    ''' <remarks></remarks>
    Public Sub deleteProjectVariantTimeStamp(ByVal pname As String, ByVal variantName As String, _
                                                  ByVal timeStamp As Date, ByRef first As Boolean)

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
            Call MsgBox("hier ist was faul" & timeStamp.ToShortDateString & vbLf & _
                         hproj.timeStamp.ToShortDateString)
        End If
        timeStamp = hproj.timeStamp

        If IsNothing(hproj) Then
            Call MsgBox("Timestamp " & timeStamp.ToShortDateString & vbLf & _
                        "zu Projekt " & projekthistorie.First.getShapeText & " nicht gefunden")

        Else
            ' Speichern im Papierkorb, dann löschen
            If requestTrash.storeProjectToDB(hproj) Then
                If request.deleteProjectTimestampFromDB(projectname:=pname, variantName:=variantName, _
                                      stored:=timeStamp) Then
                    'Call MsgBox("ok, gelöscht")
                Else
                    Call MsgBox("Fehler beim Löschen von " & pname & ", " & variantName & ", " & _
                                timeStamp.ToShortDateString)
                End If
            Else
                ' es ging etwas schief


                Call MsgBox("Fehler beim Speichern im Papierkorb:" & vbLf & _
                            hproj.name & ", " & hproj.timeStamp.ToShortDateString)
            End If

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
    ''' baut den aktuell gültigen Treeview auf  
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub buildTreeview(ByRef projektHistorien As clsProjektDBInfos, _
                              ByRef TreeviewProjekte As TreeView, _
                              ByRef aktuelleGesamtListe As clsProjekteAlle, _
                              ByVal aKtionskennung As Integer, _
                              ByVal applyFilter As Boolean)

        Dim nodeLevel0 As TreeNode
        Dim nodeLevel1 As TreeNode
        Dim zeitraumVon As Date = StartofCalendar
        Dim zeitraumbis As Date = StartofCalendar.AddYears(20)
        Dim storedHeute As Date = Now
        Dim storedGestern As Date = StartofCalendar
        Dim pname As String = ""
        Dim variantName As String = ""
        Dim loadErrorMsg As String = ""
        Dim listOfVariantNamesDB As Collection


        Dim deletedProj As Integer = 0


        Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
        Dim requestTrash As New Request(awinSettings.databaseURL, awinSettings.databaseName & "Trash", dbUsername, dbPasswort)

        ' alles zurücksetzen 
        projektHistorien.clear()

        With TreeviewProjekte
            .Nodes.Clear()
        End With


        ' Alle Projekte aus DB
        ' projekteInDB = request.retrieveProjectsFromDB(pname, variantName, zeitraumVon, zeitraumbis, storedGestern, storedHeute, True)

        Select Case aKtionskennung

            Case PTTvActions.delFromDB
                pname = ""
                variantName = ""
                aktuelleGesamtListe.liste = request.retrieveProjectsFromDB(pname, variantName, zeitraumVon, zeitraumbis, storedGestern, storedHeute, True)
                loadErrorMsg = "es gibt keine Projekte in der Datenbank"

            Case PTTvActions.delFromSession
                aktuelleGesamtListe = AlleProjekte
                loadErrorMsg = "es sind keine Projekte geladen"

            Case PTTvActions.loadPVS    ' ur: 30.01.2015: aktuell nicht benutzt!!!
                pname = ""
                variantName = ""

                'ur: 25.01.2015 hier muss die "aktuelleGesamtListe.liste reduziert werden, da evt. ein Filter gesetzt wurde!!!!
                ' tk das applyFilter wird nachher gemacht , ausnahmslos für alle 
                aktuelleGesamtListe.liste = request.retrieveProjectsFromDB(pname, variantName, zeitraumVon, zeitraumbis, storedGestern, storedHeute, True)
                loadErrorMsg = "es gibt keine Projekte in der Datenbank"

            Case PTTvActions.loadPV
                pname = ""
                variantName = ""

                aktuelleGesamtListe.liste = request.retrieveProjectsFromDB(pname, variantName, zeitraumVon, zeitraumbis, storedGestern, storedHeute, True)
                loadErrorMsg = "es gibt keine passenden Projekte in der Datenbank"

            Case PTTvActions.activateV
                aktuelleGesamtListe = AlleProjekte
                loadErrorMsg = "es sind keine Projekte geladen"

            Case PTTvActions.deleteV
                aktuelleGesamtListe = AlleProjekte
                loadErrorMsg = "es sind keine Projekte geladen"

                ' '' ''ur: 10.08.2015: ''Case PTTvActions.definePortfolioDB
                ' '' '' ''    pname = ""
                ' '' '' ''    variantName = ""

                ' '' '' ''    aktuelleGesamtListe.liste = request.retrieveProjectsFromDB(pname, variantName, zeitraumVon, zeitraumbis, storedGestern, storedHeute, True)
                ' '' '' ''    loadErrorMsg = "es gibt keine Projekte in der Datenbank"

                ' '' '' ''Case PTTvActions.definePortfolioSE
                ' '' '' ''    pname = ""
                ' '' '' ''    variantName = ""
                ' '' '' ''    aktuelleGesamtListe = AlleProjekte
                ' '' '' ''    loadErrorMsg = "es sind keine Projekte geladen"


        End Select

        ' jetzt wird der Filter angewendet, wenn er angewendet werden soll 
        ' das wird jetzt in der Routine mitgegeben 
        If applyFilter Then
            aktuelleGesamtListe = reduzierenWgFilter(aktuelleGesamtListe)
        End If


        If aktuelleGesamtListe.Count >= 1 Then

            With TreeviewProjekte

                .CheckBoxes = True

                Dim projektliste As Collection = aktuelleGesamtListe.getProjectNames
                Dim showPname As Boolean

                For Each pname In projektliste

                    showPname = True
                    listOfVariantNamesDB = request.retrieveVariantNamesFromDB(pname)

                    ' im Falle activate Variante / Portfolio definieren: nur die Projekte anzeigen, die auch tatsächlich mehrere Varianten haben 
                    If aKtionskennung = PTTvActions.activateV Or aKtionskennung = PTTvActions.deleteV Then
                        If aktuelleGesamtListe.getVariantZahl(pname) = 0 Then
                            showPname = False
                        End If
                    End If

                    If showPname Then

                        nodeLevel0 = .Nodes.Add(pname)

                        ' Platzhalter einfügen; wird für alle Aktionskennungen benötigt
                        If aKtionskennung = PTTvActions.delFromSession Or _
                            aKtionskennung = PTTvActions.activateV Or _
                            aKtionskennung = PTTvActions.deleteV Or _
                             aKtionskennung = PTTvActions.loadPV Or _
                            aKtionskennung = PTTvActions.definePortfolioDB Or _
                            aKtionskennung = PTTvActions.definePortfolioSE Then
                            If aktuelleGesamtListe.getVariantZahl(pname) > 0 Or _
                                listOfVariantNamesDB.Count > 0 Then

                                nodeLevel0.Tag = "P"
                                nodeLevel1 = nodeLevel0.Nodes.Add("()")
                                nodeLevel1.Tag = "P"

                            Else
                                nodeLevel0.Tag = "X"
                            End If

                            '' ''ur:10.08.2015 '' hier muss im Falle Portfolio Definition das Kreuz dort gesetzt sein, was geladen ist 
                            ' '' ''If aKtionskennung = PTTvActions.definePortfolioSE Then
                            ' '' ''    If ShowProjekte.contains(pname) Then
                            ' '' ''        ' im aufrufenden Teil wird stopRecursion auf true gesetzt ... 
                            ' '' ''        nodeLevel0.Checked = True

                            ' '' ''    End If
                            ' '' ''End If


                        Else
                            nodeLevel0.Tag = "P"
                            nodeLevel1 = nodeLevel0.Nodes.Add("()")
                            nodeLevel1.Tag = "P"
                        End If
                    End If



                Next


            End With
        Else
            Call MsgBox(loadErrorMsg)
        End If


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

    ''' <summary>
    ''' Prozedur um Username und Passwort für die Datenbank-Benutzung abzufragen und auch zu testen.
    ''' </summary>
    ''' <remarks></remarks>
    Function loginProzedur() As Boolean

        appInstance.EnableEvents = True

        Dim loginDialog As New frmAuthentication
        Dim returnValue As DialogResult

        returnValue = DialogResult.Retry

        While returnValue = DialogResult.Retry

            returnValue = loginDialog.ShowDialog

        End While

        If returnValue = DialogResult.Abort Then
            'Call MsgBox("Customization-File schließen")
            Return False
        Else
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
                        newProjektliste.Add(kvp.Key, kvp.Value)
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
                    Call splitBreadCrumbFullnameTo3(fullName, elemName, breadcrumb, lfdNr)


                    cphase = kvp.Value.getPhase(elemName, breadcrumb, lfdNr)
                    Dim phaseName As String

                    If Not IsNothing(cphase) Then
                        Try

                            phaseName = kvp.Value.hierarchy.getBestNameOfID(cphase.nameID, True, False)
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

                Call splitBreadCrumbFullnameTo3(fullName, elemName, breadcrumb, lfdNr)

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

                Call splitBreadCrumbFullnameTo3(fullName, elemName, breadcrumb, lfdNr)

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
                    Call splitBreadCrumbFullnameTo3(fullName, elemName, breadcrumb, lfdNr)


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
                    Call splitBreadCrumbFullnameTo3(fullName, elemName, breadcrumb, lfdNr)

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
    ''' ruft das Formular auf, um Filter zu definieren
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub defineFilterDB()
        Dim auswahlFormular As New frmNameSelection
        Dim returnValue As DialogResult

        With auswahlFormular


            .Text = "Datenbank Filter definieren"

            '.chkbxShowObjects = False
            '.chkbxCreateCharts = False

            .chkbxOneChart.Checked = False
            .chkbxOneChart.Visible = False

            .rdbBU.Visible = True
            .pictureBU.Visible = True

            .rdbTyp.Visible = True
            .pictureTyp.Visible = True

            .repVorlagenDropbox.Visible = False
            .labelPPTVorlage.Visible = False

            ' Filter
            .filterDropbox.Visible = True
            .filterLabel.Visible = True
            .filterLabel.Text = "Name des Filters"

            ' Auswahl Speichern
            .auswSpeichern.Visible = False
            .auswSpeichern.Enabled = False

            '.showModePortfolio = True
            .menuOption = PTmenue.filterdefinieren

            .OKButton.Text = "Speichern"

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
                                              ByRef chtop As Double, ByRef chleft As Double)


        Dim repObj As Excel.ChartObject
        Dim myCollection As Collection

        Dim chWidth As Double
        Dim chHeight As Double

        ' Window Position festlegen 
        chWidth = 265 + (showRangeRight - showRangeLeft - 12 + 1) * boxWidth + (showRangeRight - showRangeLeft) * screen_correct
        chHeight = awinSettings.ChartHoehe1


        If oneChart = True Then


            ' alles in einem Chart anzeigen
            myCollection = New Collection
            For Each element As String In selCollection
                myCollection.Add(element, element)
            Next

            repObj = Nothing
            Call awinCreateprcCollectionDiagram(myCollection, repObj, chtop, chleft,
                                                              chWidth, chHeight, False, chTyp, False)

            chtop = chtop + 5
            chleft = chleft + 7
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

                chtop = chtop + 5
                chleft = chleft + 7
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
    Public Sub frmHryNameReadPPTVorlagen(ByVal menuOption As Integer, ByRef repVorlagenDropbox As System.Windows.Forms.ComboBox)


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
        ElseIf menuOption = PTmenue.reportBHTC Then

            dirname = awinPath & RepProjectVorOrdner

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
            If menuOption = PTmenue.visualisieren Or _
                menuOption = PTmenue.multiprojektReport Or _
                menuOption = PTmenue.einzelprojektReport Or _
                menuOption = PTmenue.leistbarkeitsAnalyse Then

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
        ElseIf showRangeRight - showRangeLeft > 5 Then
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
                Dim chtop As Double = 50.0 + awinSettings.ChartHoehe1
                Dim chleft As Double = (showRangeRight - 1) * boxWidth + 4

                If selectedPhases.Count > 0 Then
                    chTyp = DiagrammTypen(0)
                    Call zeichneLeistbarkeitsChart(selectedPhases, chTyp, oneChart, _
                                                   chtop, chleft)
                End If

                If selectedMilestones.Count > 0 Then
                    chTyp = DiagrammTypen(5)
                    Call zeichneLeistbarkeitsChart(selectedMilestones, chTyp, oneChart, _
                                                   chtop, chleft)
                End If

                If selectedRoles.Count > 0 Then
                    chTyp = DiagrammTypen(1)
                    Call zeichneLeistbarkeitsChart(selectedRoles, chTyp, oneChart, _
                                                   chtop, chleft)
                End If

                If selectedCosts.Count > 0 Then
                    chTyp = DiagrammTypen(2)
                    Call zeichneLeistbarkeitsChart(selectedCosts, chTyp, oneChart, _
                                                   chtop, chleft)
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

            Call MsgBox("ok, Filter gespeichert")

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
                Call splitHryFullnameTo2(fullName, pName, breadcrumb)

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
                Call splitHryFullnameTo2(fullName, msName, breadcrumb)

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

            ' Filter mit Namen "fName" in DB speichern
            Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)

            ' Datenbank ist gestartet
            If request.pingMongoDb() Then

                Dim filterToStoreInDB As clsFilter = filterDefinitions.retrieveFilter(fName)
                Dim returnvalue As Boolean = request.storeFilterToDB(filterToStoreInDB, False)
            Else
                Call MsgBox(" Datenbank-Verbindung ist unterbrochen!" & vbLf & " Filter kann nicht in DB gespeichert werden")
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
                    Call findAllTasksandInsert(aktTask_i, parentelemID, hproj, Rplan, protokollLine, zeile, protokollliste)

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
                    ImportProjekte.Add(calcProjektKey(hproj), hproj)
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

                Dim aktTask_j As rxfTask = RPLAN.task(j)

                Dim taskdauerinDays As Long = calcDauerIndays(aktTask_j.actualDate.start.Value, aktTask_j.actualDate.finish.Value)
                ' Herausfinden, ob aktTask_j Phase oder Meilenstein ist
                If taskdauerinDays > 0 And aktTask_j.taskType.type <> "MILESTONE" Then

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
                                If Not missingPhaseDefinitions.Contains(newPhaseDef.name) Then
                                    missingPhaseDefinitions.Add(newPhaseDef)
                                End If

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

                ElseIf taskdauerinDays = 0 And aktTask_j.taskType.type = "SUMMARY" Then

                    Call logfileSchreiben("Achtung, RXFImport: Die aktuelle Task '" & aktTask_j.name & "' ist eine Summary-Task mit der Duration = 0 ", hproj.name, anzFehler)

                ElseIf taskdauerinDays > 0 And aktTask_j.taskType.type = "MILESTONE" Then

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

                        ' Meilenstein wir nur in das aktuelle Projekt aufgenommen, wenn awinSettings.importUnkownNames = true 
                        ' und der Name bekannt ist (isUnkownName = false)

                        If Not isUnkownName Or awinSettings.importUnknownNames Then

                            cmilestone = New clsMeilenstein(parent:=parentphase)
                            cBewertung = New clsBewertung

                            origMSname = aktTask_j.name
                            If DateDiff(DateInterval.Month, aktTask_j.actualDate.start.Value, aktTask_j.actualDate.finish.Value) = 0 Then

                                milestonedate = aktTask_j.actualDate.start.Value
                            End If


                            ' wenn der freefloat nicht zugelassen ist und der Meilenstein ausserhalb der Phasen-Grenzen liegt 
                            ' muss abgebrochen werden 

                            If Not awinSettings.milestoneFreeFloat And _
                                (DateDiff(DateInterval.Day, parentphase.getStartDate, milestonedate) < 0 Or _
                                 DateDiff(DateInterval.Day, parentphase.getEndDate, milestonedate) > 0) Then

                                Call logfileSchreiben(("Fehler, RXFImport: Der Meilenstein liegt ausserhalb seiner Phase" & vbLf & _
                                                    origMSname & " nicht innerhalb " & parentphase.name & vbLf & _
                                                         "Korrigieren Sie bitte diese Inkonsistenz in der Datei '"), hproj.name, anzFehler)
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
                                .deliverables = deliverables
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


                                Try
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
                                Catch ex1 As Exception
                                    Throw New Exception(ex1.Message)
                                End Try
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


                Else

                    Call logfileSchreiben("Achtung, RXFImport: Die Task '" & aktTask_j.name & "' kann nicht als Meilenstein oder Phase identifiziert werden ", hproj.name, anzFehler)

                End If      '  Ende: ist MEILENSTEIN

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
    ''' Behandelt den Fehler UnkonwnNode beim Einlesen des RXF-Files
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub deserializer_UnknownNode(sender As Object, e As XmlNodeEventArgs)
        Call MsgBox(("RXFImport: Unknown Node:" & e.Name & ControlChars.Tab & e.Text))
    End Sub 'serializer_UnknownNode


    ''' <summary>
    ''' Behandelt den Fehler UnkonwnAttribute beim Einlesen des RXF-Files
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub deserializer_UnknownAttribute(sender As Object, e As XmlAttributeEventArgs)
        Dim attr As System.Xml.XmlAttribute = e.Attr
        Call MsgBox(("RXFImport: Unknown attribute " & attr.Name & "='" & attr.Value & "'"))
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



    ''' <summary>
    ''' initialisert das Logfile
    ''' </summary>
    ''' <remarks></remarks>
    Sub logfileInit()

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
    Sub logfileSchreiben(ByVal text As String, ByVal addOn As String, ByRef anzFehler As Long)

        Dim obj As Object

        Try
            obj = CType(CType(xlsLogfile.Worksheets("logBuch"), Excel.Worksheet).Rows(1), Excel.Range).Insert(Excel.XlInsertShiftDirection.xlShiftDown)

            With CType(xlsLogfile.Worksheets("logBuch"), Excel.Worksheet)
                CType(.Cells(1, 1), Excel.Range).Value = text
                CType(.Cells(1, 2), Excel.Range).Value = addOn
                CType(.Cells(1, 3), Excel.Range).Value = Date.Now
            End With
            anzFehler = anzFehler + 1


        Catch ex As Exception

        End Try

    End Sub
    ''' <summary>
    ''' öffnet das LogFile
    ''' </summary>
    ''' <remarks></remarks>
    Sub logfileOpen()

        appInstance.ScreenUpdating = False

        ' aktives Workbook merken im Variable actualWB
        Dim actualWB As String = appInstance.ActiveWorkbook.Name

        If My.Computer.FileSystem.FileExists(awinPath & logFileName) Then
            Try
                xlsLogfile = appInstance.Workbooks.Open(awinPath & logFileName)
                myLogfile = appInstance.ActiveWorkbook.Name
            Catch ex As Exception

                logmessage = "Öffnen von " & logFileName & " fehlgeschlagen" & vbLf & _
                                                "falls die Datei bereits geöffnet ist: Schließen Sie sie bitte"
                'Call logfileSchreiben(logMessage, " ")
                Throw New ArgumentException(logmessage)

            End Try

        Else
            ' Logfile neu anlegen 
            xlsLogfile = appInstance.Workbooks.Add
            Call logfileInit()
            xlsLogfile.SaveAs(awinPath & logFileName)
            myLogfile = xlsLogfile.Name

        End If

        ' Workbook, das vor dem öffnen des Logfiles aktiv war, wieder aktivieren
        appInstance.Workbooks(actualWB).Activate()

    End Sub



    ''' <summary>
    ''' schliesst  das logfile 
    ''' </summary>  
    ''' <remarks></remarks>
    Sub logfileSchliessen()

        appInstance.EnableEvents = False
        Try

            appInstance.Workbooks(myLogfile).Close(SaveChanges:=True)

        Catch ex As Exception
            Call MsgBox("Fehler beim Schließen des Logfiles")
        End Try
        appInstance.EnableEvents = True
    End Sub

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



        Dim xmlfilename As String = awinPath & ReportProfileOrdner & "\" & profil.name & ".xml"

        Try

            Dim serializer = New DataContractSerializer(GetType(clsReport))

            Dim file As New FileStream(xmlfilename, FileMode.Create)
            serializer.WriteObject(file, profil)
            file.Close()
        Catch ex As Exception

            Call MsgBox("Beim Schreiben der XML-Datei '" & xmlfilename & "' ist ein Fehler aufgetreten !")

        End Try

    End Sub

    Public Function XMLImportReportProfil(ByVal profilName As String) As clsReport

        Dim profil As New clsReport

        Dim serializer = New DataContractSerializer(GetType(clsReport))
        Dim xmlfilename As String = awinPath & ReportProfileOrdner & "\" & profilName & ".xml"
        Try

            ' XML-Datei Öffnen
            ' A FileStream is needed to read the XML document.
            Dim file As New FileStream(xmlfilename, FileMode.Open)
            profil = serializer.ReadObject(file)
            file.Close()

            XMLImportReportProfil = profil

        Catch ex As Exception

            Call MsgBox("Beim Lesen der XML-Datei '" & xmlfilename & "' ist ein Fehler aufgetreten !")
            XMLImportReportProfil = Nothing
        End Try

    End Function

    Public Sub retrieveProfilSelection(ByVal profilName As String, ByVal menuOption As Integer, _
                                     ByRef selectedBUs As Collection, ByRef selectedTyps As Collection, _
                                     ByRef selectedPhases As Collection, ByRef selectedMilestones As Collection, _
                                     ByRef selectedRoles As Collection, ByRef selectedCosts As Collection, ByRef reportProfil As clsReport)
        Try

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


            ' Einlesen des ausgewählten ReportProfils
            reportProfil = XMLImportReportProfil(profilName)

            ' Datumsangaben zurücksichern
            reportProfil.CalendarVonDate = PPTvondate_sav
            reportProfil.CalendarBisDate = PPTbisdate_sav
            reportProfil.VonDate = vondate_sav
            reportProfil.BisDate = bisdate_sav

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
                                     ByVal selectedRoles As Collection, ByVal selectedCosts As Collection, ByVal reportProfil As clsReport)



        '  und bereitstellen der Auswahl für Hierarchieselection
        reportProfil.Phases = copyColltoSortedList(selectedPhases)
        reportProfil.Milestones = copyColltoSortedList(selectedMilestones)
        reportProfil.Roles = copyColltoSortedList(selectedRoles)
        reportProfil.Costs = copyColltoSortedList(selectedCosts)
        reportProfil.BUs = copyColltoSortedList(selectedBUs)
        reportProfil.Typs = copyColltoSortedList(selectedTyps)


        With awinSettings

            .drawProjectLine = True
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
        End With


        ' Schreiben des ausgewählten ReportProfils
        Call XMLExportReportProfil(reportProfil)

    End Sub
    ''' <summary>
    ''' 
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
                    Dim srcDate As Date = My.Computer.FileSystem.GetFileInfo(srcFile).LastWriteTime.Date
                    Dim destDate As Date = My.Computer.FileSystem.GetFileInfo(destFile).LastWriteTime.Date
                    Dim ddiff As Long = DateDiff(DateInterval.Second, _
                                                 My.Computer.FileSystem.GetFileInfo(srcFile).LastWriteTime.Date, _
                                                 My.Computer.FileSystem.GetFileInfo(destFile).LastWriteTime.Date)

                    ' Wenn globales neuer als lokales, dann von globalPath nach awinPath kopieren
                    If ddiff < 0 Then
                        ' Kopieren der Datei, mit Overwrite erzwingen
                        My.Computer.FileSystem.CopyFile(srcFile, destFile, True)
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
                        Dim srcDate As Date = My.Computer.FileSystem.GetFileInfo(srcFile).LastWriteTime.Date
                        Dim destDate As Date = My.Computer.FileSystem.GetFileInfo(destFile).LastWriteTime.Date
                        Dim ddiff As Long = DateDiff(DateInterval.Second, _
                                                     My.Computer.FileSystem.GetFileInfo(srcFile).LastWriteTime.Date, _
                                                     My.Computer.FileSystem.GetFileInfo(destFile).LastWriteTime.Date)

                        ' Wenn globales neuer als lokales, dann von globalPath nach awinPath kopieren
                        If ddiff < 0 Then
                            ' Kopieren der Datei, mit Overwrite erzwingen
                            My.Computer.FileSystem.CopyFile(srcFile, destFile, True)
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

            Dim file As New FileStream(xmlfilename, FileMode.Create)
            serializer.WriteObject(file, lic)
            file.Close()
        Catch ex As Exception

            Call MsgBox("Beim Schreiben der XML-Datei '" & xmlfilename & "' ist ein Fehler aufgetreten !")

        End Try

    End Sub

    Public Function XMLImportLicences(ByVal licfile As String) As clsLicences

        Dim lic As New clsLicences

        Dim serializer = New DataContractSerializer(GetType(clsLicences))
        Dim xmlfilename As String = licfile
        Try

            ' XML-Datei Öffnen
            ' A FileStream is needed to read the XML document.
            Dim file As New FileStream(xmlfilename, FileMode.Open)
            lic = serializer.ReadObject(file)
            file.Close()

            XMLImportLicences = lic

        Catch ex As Exception

            Call MsgBox("Beim Lesen der XML-Datei '" & xmlfilename & "' ist ein Fehler aufgetreten !")
            XMLImportLicences = Nothing
        End Try

    End Function
End Module
