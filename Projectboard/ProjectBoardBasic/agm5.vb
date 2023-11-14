Imports ProjectBoardDefinitions
Imports DBAccLayer
Imports ProjectboardReports
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.ComponentModel
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
Imports System.Text.RegularExpressions


Public Module agm5
    Private Enum allianzSpalten
        Name = 0
        AmpelText = 1
        BusinessUnit = 2
        Description = 3
        Responsible = 4
        Budget = 5
        Projektnummer = 6
        Status = 7
        itemType = 8
        pvBudget = 9
    End Enum

    Private Enum allianzBOBSpalten
        budgetgruppe = 0
        satzart = 1
        bwla = 2
        bobname = 3
        scopename = 4
        scopeID = 5
        scopedesc = 6
        scopeStart = 7
        scopeEnd = 8
        businessUnit = 9
        scopeVerantw = 10
        budget = 11
        budgetExpl = 12
        itemType = 13
        scopeTyp = 14
        budgetBeauftragt = 15
        bobID = 16
        bobdesc = 17
        bobVerantw = 18
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
    ''' wird aktuell nirgends verwendet  
    ''' </summary>
    ''' <param name="myCollection"></param>
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
            Dim activeWSListe As Excel.Worksheet = CType(appInstance.ActiveWorkbook.Worksheets("Tabelle1"),
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
                                hproj = erstelleInventurProjekt(pName, vorlagenName, scenarioName,
                                                             start, ende, budget, zeile, sfit, risk,
                                                             capacityNeeded, Nothing, businessUnit, description, Nothing, "", 0.0, 0.0)

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

                                        hproj.AddPhase(parentPhase, origName:=phaseName,
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

        currentConstellationPvName = scenarioName

    End Sub



    ''' <summary>
    ''' Einlesen eines RXF-Files (XML-Ausleitung von RPLAN) und dazu ein Protokoll in Tabellenblatt 'xmlfilename'protokoll in Datei Logfile
    ''' </summary>
    ''' <param name="myCollection"></param>
    ''' <param name="xmlfilename"></param>Name des RXF-Files
    ''' <param name="isVorlage"></param>Ist Vorlage, oder nicht
    ''' <remarks></remarks>
    Public Sub RXFImport(ByRef myCollection As Collection, ByVal xmlfilename As String,
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
                    Dim hrchynode As New clsHierarchyNode With {
                        .elemName = cphase.name,
                        .parentNodeKey = ""
                    }
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
                        'hproj.farbe = vproj.farbe
                        hproj.Schrift = vproj.Schrift
                        hproj.Schriftfarbe = vproj.Schriftfarbe
                        hproj.earliestStart = vproj.earliestStart
                        hproj.latestStart = vproj.latestStart

                        'ElseIf Projektvorlagen.Contains("unknown") Then
                        '    vproj = Projektvorlagen.getProject("unknown")
                    Else
                        'Throw New Exception("es gibt weder die Vorlage 'unknown' noch die Vorlage " & vorlagenName)
                        hproj.VorlagenName = ""
                        'hproj.farbe = awinSettings.AmpelNichtBewertet
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
            Call logger(ptErrLevel.logError, ex.Message & vbLf & "Fehler bei Name " & CStr(aktuellerName), aktuellerName, anzFehler)
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
    Public Sub findAllTasksandInsert(ByVal task As rxfTask, ByVal parentelemID As String, ByRef hproj As clsProjekt, ByVal RPLAN As rxf, ByRef prtLine As clsProtokoll, ByRef zeile As Integer, ByRef prtliste As SortedList(Of Integer, clsProtokoll))


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

                    Dim isKnownMsName As Boolean = MilestoneDefinitions.Contains(aktTask_j.name) Or
                                                missingMilestoneDefinitions.Contains(aktTask_j.name)

                    Dim isKnownPhName As Boolean = PhaseDefinitions.Contains(aktTask_j.name) Or
                                                missingPhaseDefinitions.Contains(aktTask_j.name)

                    Dim taskdauerinDays As Long = calcDauerIndays(aktTask_j.actualDate.start.Value, aktTask_j.actualDate.finish.Value)
                    ' Herausfinden, ob aktTask_j Phase oder Meilenstein ist 

                    If taskdauerinDays > 1 Then
                        isMilestone = False

                        If aktTask_j.taskType.type = "MILESTONE" Then
                            Call logger(ptErrLevel.logError, "Korrektur, RXFImport: Phasen-Element mit verschiedenen Start- und Ende-Daten war als Meilenstein deklariert:",
                                                        aktTask_j.name & ": " & aktTask_j.actualDate.start.Value.ToShortDateString & " versus " &
                                                        aktTask_j.actualDate.finish.Value.ToShortDateString & vbLf &
                                                        "Projekt: " & hproj.name,
                                                        anzFehler)
                        End If

                    ElseIf aktTask_j.taskType.type = "MILESTONE" Then
                        isMilestone = True

                    ElseIf isKnownMsName And Not isKnownPhName Then
                        isMilestone = True
                        If aktTask_j.taskType.type <> "MILESTONE" Then
                            Call logger(ptErrLevel.logError, "Korrektur, RXFImport: bekanntes Meilenstein-Element  mit falscher Typ-Zuordnung:",
                                                        aktTask_j.name & " mit Typ " & aktTask_j.taskType.type & vbLf &
                                                        "Projekt: " & hproj.name,
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
                                            Dim duplicateSiblingID As String = hproj.getDuplicatePhaseSiblingID(mappedPhasename, parentphase.nameID,
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
                                    Throw New Exception("Fehler, RXFImport: Der Meilenstein hat verschiedene Start- und End-Daten:" & vbLf &
                                                        aktTask_j.actualDate.start.Value.ToShortDateString & " versus " &
                                                        aktTask_j.actualDate.finish.Value.ToShortDateString & vbLf &
                                                        "Projekt: " & hproj.name)
                                End If



                                ' wenn der freefloat nicht zugelassen ist und der Meilenstein ausserhalb der Phasen-Grenzen liegt 
                                ' muss abgebrochen werden 

                                If Not awinSettings.milestoneFreeFloat And
                                    (DateDiff(DateInterval.Day, parentphase.getStartDate, milestonedate) < 0 Or
                                     DateDiff(DateInterval.Day, parentphase.getEndDate, milestonedate) > 0) Then

                                    'Call logfileSchreiben(("Fehler, RXFImport: Der Meilenstein liegt ausserhalb seiner Phase" & vbLf & _
                                    '                    origMSname & " nicht innerhalb " & parentphase.name & vbLf & _
                                    '                         "Korrigieren Sie bitte diese Inkonsistenz in der Datei '"), hproj.name, anzFehler)
                                    Throw New Exception("Fehler, RXFImport: Der Meilenstein liegt ausserhalb seiner Phase" & vbLf &
                                                        origMSname & " nicht innerhalb " & parentphase.name & vbLf &
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
                                    Dim duplicateSiblingID As String = hproj.getDuplicateMsSiblingID(mappedMSname, parentphase.nameID,
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

                    Call logger(ptErrLevel.logError, ex.Message, hproj.name, anzFehler)


                End Try


            End If

        Next j    ' Ende Schleife über alle Tasks
    End Sub



    ''' <summary>
    ''' in der ganzen Datei sfilename wird der String searchstr durch replacestr ersetzt
    ''' </summary>
    ''' <param name="sfilename"></param>Name der Datei, in der die Ersetzung erfolgen soll
    ''' <param name="searchstr"></param>zu ersetzender String
    ''' <param name="replacestr"></param>neuer String
    ''' <returns></returns>Name der neuen Datei
    ''' <remarks></remarks>
    Public Function replaceStringInFile(ByVal sfilename As String, ByVal searchstr As String, ByVal replacestr As String) As String

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
    ''' nach BMW-Vorgaben:
    ''' bestimmt aus dem übergebenen VorlagenNamen ( =  der CustomValue "UsA_SERVICE_SPALTE_B" aus Phase "Projektphasen" ) 
    ''' den tatsächlichen VorlagenNamen des Projekts 
    '''     ''' </summary>
    ''' <param name="hproj"></param>aktuelles zu lesendes Projekt
    ''' <returns></returns>fertig zusammengesetzter VorlagenName des Projekts (gemäß BMW vorschriften
    ''' <remarks></remarks>
    Public Function findBMWVorlagenName(ByVal hproj As clsProjekt) As String


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
    Public Sub deserializer_UnknownNode(sender As Object, e As XmlNodeEventArgs)
        Call MsgBox(("XMLImport: Unknown Node:" & e.Name & ControlChars.Tab & e.Text))
    End Sub 'serializer_UnknownNode


    ''' <summary>
    ''' Behandelt den Fehler UnkonwnAttribute beim Einlesen eines XML-Files (oder RXF-Files)
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Public Sub deserializer_UnknownAttribute(sender As Object, e As XmlAttributeEventArgs)
        Dim attr As System.Xml.XmlAttribute = e.Attr
        Call MsgBox(("XMLImport: Unknown attribute " & attr.Name & "='" & attr.Value & "'"))
    End Sub 'serializer_UnknownAttribute




    ''' <summary>
    ''' liest aus einer Excel-Tabelle die ggf vorhandenen CustomFields, der arrayOfSpalten gibt an, welche Spalten ausgelesen werden müssen 
    ''' </summary>
    ''' <param name="arrayOfSpalten">Indices der Spalten, müssen nicht zusammenhängend sein</param>
    ''' <param name="Headerzeile"><include file='welcher Zeile steht die Überschrift ' path='[@name=""]'/></param>
    ''' <param name="curZeile">welche Zeile der Tabelle soll gerade ausgelesen werden </param>
    ''' <param name="currentWS">das Excel.Worksheet, das die Tabelle enthält</param>
    ''' <returns></returns>
    Public Function readCustomFieldsFromExcel(ByVal arrayOfSpalten() As Integer,
                                               ByVal Headerzeile As Integer, ByVal curzeile As Integer,
                                               ByVal currentWS As Excel.Worksheet) As Collection

        Dim custFields As New Collection

        If Not IsNothing(arrayOfSpalten) Then

            For i As Integer = 0 To arrayOfSpalten.Length - 1

                Dim spalte As Integer = arrayOfSpalten(i)

                With currentWS
                    Try
                        Dim cfName As String = CStr(CType(.Cells(Headerzeile, spalte), Global.Microsoft.Office.Interop.Excel.Range).Value)
                        Dim uniqueID As Integer = customFieldDefinitions.getUid(cfName)

                        If uniqueID > 0 Then
                            ' es ist eine Custom Field

                            Dim cfType As Integer = customFieldDefinitions.getTyp(uniqueID)
                            Dim cfValue As Object = Nothing
                            Dim tstStr As String

                            Select Case cfType
                                Case ptCustomFields.Str

                                    cfValue = CStr(CType(.Cells(curzeile, spalte), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                Case ptCustomFields.Dbl

                                    cfValue = CDbl(CType(.Cells(curzeile, spalte), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                Case ptCustomFields.bool

                                    cfValue = CBool(CType(.Cells(curzeile, spalte), Global.Microsoft.Office.Interop.Excel.Range).Value)
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
                        CType(.Cells(curzeile, spalte), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color = awinSettings.AmpelGelb
                        CType(.Cells(curzeile, spalte), Global.Microsoft.Office.Interop.Excel.Range).AddComment(Text:=ex.Message)
                    End Try
                End With


            Next

        End If

        readCustomFieldsFromExcel = custFields

    End Function


    ''' <summary>
    ''' wenn in der Spalte ein #ABP oder eine sonstige Allianz Gruppe vermerkt ist, dann zurückgeben 
    ''' </summary>
    ''' <param name="excelCell"></param>
    ''' <returns></returns>
    Private Function getAllianzTeamNameFromCell(ByVal excelCell As Excel.Range) As String
        Dim tmpResult As String = ""

        Try
            If Not IsNothing(excelCell) Then
                If Not IsNothing(excelCell.Value) Then
                    Dim cellValue As String = CStr(excelCell.Value).Trim
                    If cellValue.StartsWith("#") Then
                        ' 1. Versuch
                        Dim tmpStr1() As String = cellValue.Split(New Char() {CChar("-"), CChar("_"), CChar(" ")})

                        If RoleDefinitions.containsName(tmpStr1(0).Trim) Then

                            If RoleDefinitions.getRoledef(tmpStr1(0).Trim).isSkill Then
                                tmpResult = tmpStr1(0).Trim
                            End If
                        End If


                        If tmpResult = "" Then
                            ' 2. Versuch
                            Dim tmpStr2() As String = cellValue.Split(New Char() {CChar(" ")})
                            If tmpStr2.Length > 1 Then
                                Dim tmpName As String = "#" & tmpStr2(1).Trim
                                If RoleDefinitions.containsName(tmpName) Then
                                    If RoleDefinitions.getRoledef(tmpName).isSkill Then
                                        tmpResult = tmpName
                                    End If
                                End If
                            End If

                            If tmpResult = "" Then
                                ' 3. Versuch
                                Dim tmpstr4() As String = cellValue.Split(New Char() {CChar("("), CChar(")")})
                                If tmpstr4.Length > 1 Then
                                    Dim tmpName As String = "#" & tmpstr4(1).Trim
                                    If RoleDefinitions.containsName(tmpName) Then
                                        If RoleDefinitions.getRoledef(tmpName).isSkill Then
                                            tmpResult = tmpName
                                        End If
                                    End If
                                End If

                            End If

                        End If
                    End If
                End If

            End If

        Catch ex As Exception

        End Try



        getAllianzTeamNameFromCell = tmpResult
    End Function

    ''' <summary>
    ''' setzt die für den Allianz 1 Import Type notwendigen Felder
    ''' </summary>
    ''' <param name="PTroleNamesToConsider"></param>
    ''' <param name="PTcolRoleNamesToConsider"></param>
    ''' <param name="TEroleNamesToConsider"></param>
    ''' <param name="TEcolRoleNamesToConsider"></param>
    ''' <param name="currentWS"></param>
    ''' <param name="importTyp"></param>
    Private Sub setAllianzImportArrays(ByRef PTroleNamesToConsider() As String,
                                       ByRef PTcolRoleNamesToConsider() As Integer,
                                       ByRef TEroleNamesToConsider() As String,
                                       ByRef TEcolRoleNamesToConsider() As Integer,
                                       ByVal currentWS As Excel.Worksheet,
                                       ByVal importTyp As ptImportTypen)

        Dim tmpRoleNames() As String = Nothing
        Dim tmpColBz() As String = Nothing
        Dim tmpCols() As Integer

        Dim tmpTEroleNames() As String
        Dim tmpTEcolBZ() As String
        Dim tmpTECols() As Integer

        Dim errRoles As String = ""
        Dim ok As Boolean = True
        Dim zeile As Integer = 2

        ' wird im Fall BOB einlesen benötigt 
        Dim startCol As Integer
        Dim endCol As Integer



        If importTyp = ptImportTypen.allianzMassImport1 Then

            zeile = 2
            ' am besten hier aus awinsettings einlesen ...
            ' sowohl die PTRoleNames als auch die T€RoleNames 
            tmpRoleNames = {"D-BOSV-KB0", "D-BOSV-KB1", "D-BOSV-KB2", "D-BOSV-KB3", "D-BOSV-SBF1", "D-BOSV-SBF2", "DRUCK", "D-BOSV-SBP1", "D-BOSV-SBP2", "D-BOSV-SBP3", "AMIS",
                        "IT-BVG", "IT-KuV", "IT-PSQ", "A-IT04", "AZ Technology", "IT-SFK", "Op-DFS", "KaiserX IT"}

            tmpColBz = {"DB1", "DC1", "DD1", "DE1", "DG1", "DH1", "DI1", "DK1", "DL1", "DM1", "DN1", "DP1", "DQ1", "DR1", "DS1", "DT1", "DU1", "DV1", "DW1"}

            ReDim tmpCols(tmpRoleNames.Length - 1)

            tmpTEroleNames = Nothing
            tmpTEcolBZ = Nothing
            ReDim tmpTECols(0)

        ElseIf importTyp = ptImportTypen.allianzMassImport2 Then
            zeile = 3
            tmpRoleNames = {"D-BITSV-KB0", "D-BITSV-KB1", "D-BITSV-KB2", "D-BITSV-KB3", "D-BITSV-SBF1", "D-BITSV-SBF2", "D-BITSV-SBF-DRUCK", "D-BITSV-SBP1", "D-BITSV-SBP2", "D-BITSV-SBP3", "AMIS"}
            tmpColBz = {"CP1", "CQ1", "CR1", "CS1", "CU1", "CV1", "CW1", "CY1", "CZ1", "DA1", "DB1"}

            ReDim tmpCols(tmpRoleNames.Length - 1)

            tmpTEroleNames = {"D-BITKuV", "D-BITLuA", "D-BITKIS", "D-BITEPM", "D-BIT-FMV", "D-IT-BVG", "D-BITKVI", "D-IT-PSQ", "A-IT04", "D-IT-AS", "AMOS", "KX BIT", "KX IT", "D-IT-ISM"}
            tmpTEcolBZ = {"AN1", "AP1", "AQ1", "AR1", "AS1", "AT1", "AU1", "AV1", "AW1", "AX1", "AY1", "AZ1", "BA1", "BB1"}

            ReDim tmpTECols(tmpTEroleNames.Length - 1)

        Else
            ' BOB Import
            zeile = 6

            ' compiler Hygiene 30.11.19
            ' die beiden Zeilen sind nur nötig, damit Compiler keine Warning bringt; programatisch wird das hier nicht gebraucht 
            ReDim tmpCols(0)
            ReDim tmpTEcolBZ(0)
            ' Ende Compiler Hygiene

            With currentWS
                startCol = CInt(CType(.Range("BC1"), Excel.Range).Column)
                endCol = CInt(CType(.Range("DD1"), Excel.Range).Column)
            End With


            ReDim tmpTECols(endCol - startCol)
            ReDim tmpTEroleNames(endCol - startCol)

            ' tmpTERoleNames holen 
            For i As Integer = startCol To endCol
                tmpTECols(i - startCol) = i
                With currentWS
                    tmpTEroleNames(i - startCol) = CStr(CType(.Cells(zeile, i), Excel.Range).Value).Trim
                End With
            Next

        End If

        If importTyp = ptImportTypen.allianzBOBImport Then

            For i As Integer = 0 To endCol - startCol - 1
                If Not RoleDefinitions.containsName(tmpTEroleNames(i)) Then
                    errRoles = errRoles & tmpTEroleNames(i) & "; "
                End If
            Next

            If errRoles.Length > 0 Then
                Throw New ArgumentException("nicht bekannte Rolle(n: " & errRoles)
            End If
        Else

            If (tmpRoleNames.Length <> tmpColBz.Length) Or (tmpTEroleNames.Length <> tmpTEcolBZ.Length) Then
                Throw New ArgumentException("ungleiche Anzahl Namen und Spalten-Ids")
            Else
                Dim tmpAnzahl As Integer = tmpRoleNames.Length

                ' Plausibilitätsprüfung: nur weitermachen, wenn auch alle Rollen in der RollenDefinition drin sind 

                For Each tmpRoleName As String In tmpRoleNames
                    If RoleDefinitions.containsName(tmpRoleName) Then
                        ' ok 
                    Else
                        errRoles = errRoles & tmpRoleName & "; "
                    End If
                Next

                For Each tmpRoleName As String In tmpTEroleNames
                    If RoleDefinitions.containsName(tmpRoleName) Then
                        ' ok 
                    Else
                        errRoles = errRoles & tmpRoleName & "; "
                    End If
                Next

                If errRoles.Length = 0 Then
                    ' jetzt weiter machen .. die col
                    With currentWS
                        For i As Integer = 1 To tmpAnzahl
                            If Not IsNothing(tmpCols) Then

                            End If
                            tmpCols(i - 1) = CType(.Range(tmpColBz(i - 1)), Excel.Range).Column

                            ' test tk 9.6.18
                            If tmpRoleNames(i - 1).StartsWith("D-") Then

                                Dim tmpValue As String = CStr(CType(.Cells(zeile, tmpCols(i - 1)), Excel.Range).Value).Trim
                                Dim chkTxt As String = tmpRoleNames(i - 1).Trim.Substring(2)

                                ok = tmpValue.StartsWith(chkTxt) Or tmpRoleNames(i - 1) = "DRUCK"
                            Else
                                ok = ok And (CStr(CType(.Cells(zeile, tmpCols(i - 1)), Excel.Range).Value).StartsWith(tmpRoleNames(i - 1)) Or
                            tmpRoleNames(i - 1) = "DRUCK")
                            End If


                            If Not ok Then
                                Call MsgBox("Fehler in Spalte mit Angaben zu (?) " & tmpRoleNames(i - 1))
                                ok = True
                            End If
                        Next

                        tmpAnzahl = tmpTEroleNames.Length

                        For i As Integer = 1 To tmpAnzahl

                            tmpTECols(i - 1) = CType(.Range(tmpTEcolBZ(i - 1)), Excel.Range).Column

                            Dim tmpValue As String = CStr(CType(.Cells(zeile, tmpTECols(i - 1)), Excel.Range).Value).Trim

                            ok = tmpValue.StartsWith(tmpTEroleNames(i - 1))

                            If Not ok Then
                                Call MsgBox("Fehler in Spalte mit Angaben zu (?) " & tmpRoleNames(i - 1))
                                ok = True
                            End If
                        Next

                    End With

                Else
                    Throw New ArgumentException("nicht bekannte Rolle(n: " & errRoles)
                End If
            End If

        End If

        PTroleNamesToConsider = tmpRoleNames
        PTcolRoleNamesToConsider = tmpCols

        TEroleNamesToConsider = tmpTEroleNames
        TEcolRoleNamesToConsider = tmpTECols


    End Sub



    ''' <summary>
    ''' erzeugt die BOBs (Portfolios) und Scopes (Projekte  die in der Batch-Datei angegeben sind
    ''' stellt sie in ImportProjekte 
    ''' erstellt ein Szenario mit Namen der Batch-Datei; die Sortierung erfolgt über die Reihenfolge in der Batch-Datei 
    ''' das wird sichergestellt über Eintrag der tfzeile in hproj ... 
    ''' </summary>
    ''' <remarks></remarks>
    Sub importAllianzBOBS(ByVal startdate As Date, ByVal endDate As Date)

        Dim zeile As Integer, spalte As Integer

        Dim importType As ptImportTypen

        ' tk : nimmt die Start- bzw Ende-Daten auf ...
        Dim pStartDatum As Date
        Dim pEndDatum As Date


        Dim tfZeile As Integer = 2

        Dim pName As String = ""
        Dim variantName As String = ""
        Dim custFields As New Collection
        Dim description As String = ""
        Dim responsiblePerson As String = ""
        Dim sFit As Double = 5.0
        Dim risk As Double = 5.0
        Dim budget As Double = 0.0
        Dim budgetBeauftragt As Double = 0.0
        Dim businessUnit As String = ""
        Dim allianzProjektNummer As String = ""
        Dim allianzStatus As String = ""
        Dim ampelText As String
        Dim projVorhabensBudget As Double = 0.0

        Dim logmsg() As String

        Dim programName As String = ""
        Dim current1program As clsConstellation = Nothing

        Dim last1Budget As Double = 0.0
        Dim lfdNr1program As Integer = 2

        ' nimmt die vollen Namen der 
        Dim fullNameListe1 As New SortedList(Of String, String)

        Dim createdProjects As Integer = 0
        Dim createdPrograms As Integer = 0
        Dim emptyPrograms As Integer = 0



        Dim lastRow As Integer
        Dim lastColumn As Integer
        Dim geleseneProjekte As Integer
        Dim ok As Boolean = False

        ' für den Output 
        Dim outputFenster As New frmOutputWindow
        Dim outputCollection As New Collection
        Dim outPutLine As String = ""

        ' Standard-Definition
        Dim anzReleases As Integer = 5


        Dim vorlageName As String = "Run"

        Try
            anzReleases = Projektvorlagen.getProject(vorlageName).CountPhases - 1
        Catch ex As Exception

        End Try


        ' enthält die prozentualen Anteile in den Releases 
        Dim relPrz() As Double
        ReDim relPrz(anzReleases - 1)

        ' Projekt-Eintrag/Zeile, die in der Excel Datei ignoriert werden soll 
        Dim nameTobeIgnored As String = "xxx"

        ' enthält die Phasen Namen
        Dim phNames() As String
        ReDim phNames(anzReleases - 1)

        ' enthält die Spalten-Nummer, ab der die Release Phasen Anteile stehen 
        Dim colRelPrzStart As Integer

        ' enthält die Info, welche Rollen-Namen berücksichtigt werden sollen 
        Dim roleNamesToConsider() As String = Nothing

        ' enthält die Spalten-Nummern, wo die einzelnen Rollen-Namen zu finden sind
        Dim colRoleNamesToConsider() As Integer = Nothing

        ' enthält, wieviel Manntage von dieser Rolle insgesamt benötigt werden 
        Dim roleNeeds() As Double = Nothing

        ' enthält die Info, welche Rollen-Namen berücksichtigt werden sollen 
        Dim TEroleNamesToConsider() As String = Nothing

        ' enthält die Spalten-Nummern, wo die einzelnen Rollen-Namen zu finden sind
        Dim colTEroleNamesToConsider() As Integer = Nothing

        ' enthält, wieviel Manntage von dieser Rolle insgesamt benötigt werden 
        Dim costNeeds() As Double = Nothing

        ' enthält die Spalten, wo die einzelnen Felder stehen , korreliert mit der Enum allianzSpalten
        Dim colFields() As Integer

        Dim firstZeile As Excel.Range

        Dim enumAllianzCount As Integer = [Enum].GetNames(GetType(allianzBOBSpalten)).Length
        ReDim colFields(enumAllianzCount)



        spalte = 1
        geleseneProjekte = 0

        ' jetzt werden die Phase-Names besetzt
        Try
            For i = 1 To anzReleases
                phNames(i - 1) = Projektvorlagen.getProject(vorlageName).getPhase(i + 1).name
            Next
        Catch ex As Exception
            Call MsgBox("Probleme mit Vorlage " & vorlageName)
            Exit Sub
        End Try


        Try

            importType = ptImportTypen.allianzBOBImport
            Dim currentWS As Excel.Worksheet = bestimmeWsAndImporttype(importType)


            With currentWS

                colRelPrzStart = 0
                firstZeile = CType(.Rows(6), Excel.Range)
                zeile = 7


                ' damit werden die Arrays besetzt, welche Rollen gesucht sind und in welchen Spalten die Angaben dazu zu finden sind ... 
                Call setAllianzImportArrays(roleNamesToConsider, colRoleNamesToConsider,
                                            TEroleNamesToConsider, colTEroleNamesToConsider,
                                            currentWS, importType)


                'lastColumn = firstZeile.End(XlDirection.xlToLeft).Column

                lastColumn = CType(.Cells(1, 3000), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlToLeft).Column
                lastRow = CType(.Cells(5000, "B"), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlUp).Row


                ' um die CustomFields lesen zu können ... 
                Dim colCustomFields(1) As Integer


                ' BudgetGruppe
                colCustomFields(0) = CInt(CType(.Range("A1"), Excel.Range).Column)
                ' BWLA
                colCustomFields(1) = CInt(CType(.Range("C1"), Excel.Range).Column)

                ' jetzt die Spalten bestimmen, wo die Werte stehen
                Try
                    colFields(allianzBOBSpalten.bobname) = CType(.Range("F1"), Excel.Range).Column
                    colFields(allianzBOBSpalten.bobID) = CType(.Range("G1"), Excel.Range).Column
                    colFields(allianzBOBSpalten.bobdesc) = CType(.Range("H1"), Excel.Range).Column
                    colFields(allianzBOBSpalten.bobVerantw) = CType(.Range("D1"), Excel.Range).Column
                    colFields(allianzBOBSpalten.budget) = CType(.Range("AF1"), Excel.Range).Column
                    colFields(allianzBOBSpalten.budgetBeauftragt) = CType(.Range("AL1"), Excel.Range).Column
                    colFields(allianzBOBSpalten.budgetExpl) = CType(.Range("AK1"), Excel.Range).Column
                    colFields(allianzBOBSpalten.budgetgruppe) = CType(.Range("A1"), Excel.Range).Column
                    colFields(allianzBOBSpalten.businessUnit) = CType(.Range("O1"), Excel.Range).Column
                    colFields(allianzBOBSpalten.bwla) = CType(.Range("C1"), Excel.Range).Column
                    colFields(allianzBOBSpalten.itemType) = CType(.Range("B1"), Excel.Range).Column
                    colFields(allianzBOBSpalten.satzart) = CType(.Range("B1"), Excel.Range).Column
                    colFields(allianzBOBSpalten.scopedesc) = CType(.Range("K1"), Excel.Range).Column
                    colFields(allianzBOBSpalten.scopeEnd) = CType(.Range("N1"), Excel.Range).Column
                    colFields(allianzBOBSpalten.scopeID) = CType(.Range("J1"), Excel.Range).Column
                    colFields(allianzBOBSpalten.scopename) = CType(.Range("I1"), Excel.Range).Column
                    colFields(allianzBOBSpalten.scopeStart) = CType(.Range("M1"), Excel.Range).Column
                    colFields(allianzBOBSpalten.scopeVerantw) = CType(.Range("Q1"), Excel.Range).Column
                    colFields(allianzBOBSpalten.scopeTyp) = CType(.Range("L1"), Excel.Range).Column

                Catch ex As Exception
                    Dim errmsg As String = "fehlerhafte Range Definition ..."
                    Throw New ArgumentException(errmsg)
                End Try

                Dim realRoleNamesToConsider() As String = TEroleNamesToConsider


                ' tk Test Logfile schreiben ...
                If awinSettings.visboDebug Then
                    Try
                        ReDim logmsg(realRoleNamesToConsider.Count)
                        logmsg(0) = ""
                        For ix As Integer = 1 To realRoleNamesToConsider.Count
                            logmsg(ix) = realRoleNamesToConsider(ix - 1)
                        Next
                        Call logger(ptErrLevel.logDebug, "importAllianzBOBS", logmsg)
                    Catch ex As Exception

                    End Try
                End If


                ' jetzt die zugelassenen Werte für 
                Dim pgmlinie() As Integer
                Dim projektvorhaben() As Integer

                ReDim pgmlinie(0)
                ' nach Relco mit Rupi am 8.8.19 so entschieden
                ReDim projektvorhaben(0)

                pgmlinie(0) = 3
                projektvorhaben(0) = 4


                ' jetzt müssen die Dimensionen gesetzt werden 
                Dim tmpLen As Integer = realRoleNamesToConsider.Length
                ReDim roleNeeds(tmpLen - 1)


                While zeile <= lastRow

                    Dim boBName As String = ""
                    Dim scopeName As String = ""
                    Dim itemType As Integer = 0

                    pName = ""
                    vorlageName = "Run"

                    ' Werte zurücksetzen ..
                    ReDim roleNeeds(tmpLen - 1)

                    ok = False

                    ' Kommentare zurücksetzen ...
                    Try
                        CType(.Range(.Cells(zeile, 1), .Cells(zeile, lastColumn)), Global.Microsoft.Office.Interop.Excel.Range).ClearComments()
                    Catch ex As Exception

                    End Try

                    Try
                        If Not IsNothing(CType(.Cells(zeile, colFields(allianzBOBSpalten.itemType)), Excel.Range).Value) Then
                            itemType = CInt(CType(.Cells(zeile, colFields(allianzBOBSpalten.itemType)), Excel.Range).Value)
                        Else
                            itemType = 0
                        End If

                    Catch ex As Exception
                        itemType = 0
                    End Try

                    If pgmlinie.Contains(itemType) Or projektvorhaben.Contains(itemType) Then

                        ' lese den Scope bzw Bob-Name 
                        Try
                            If pgmlinie.Contains(itemType) Then
                                If Not IsNothing(CType(.Cells(zeile, colFields(allianzBOBSpalten.bobname)), Excel.Range).Value) Then
                                    boBName = CStr(CType(.Cells(zeile, colFields(allianzBOBSpalten.bobname)), Excel.Range).Value).Trim
                                    If boBName <> "" Then
                                        pName = boBName

                                        ' jetzt description, ID  und verantwortlich rauslesen ... 
                                        If Not IsNothing(CType(.Cells(zeile, colFields(allianzBOBSpalten.bobdesc)), Excel.Range).Value) Then
                                            description = CStr(CType(.Cells(zeile, colFields(allianzBOBSpalten.bobdesc)), Excel.Range).Value).Trim
                                        Else
                                            description = ""
                                        End If


                                        If Not IsNothing(CType(.Cells(zeile, colFields(allianzBOBSpalten.bobID)), Excel.Range).Value) Then
                                            allianzProjektNummer = CStr(CType(.Cells(zeile, colFields(allianzBOBSpalten.bobID)), Excel.Range).Value).Trim
                                        Else
                                            allianzProjektNummer = ""
                                        End If

                                        If Not IsNothing(CType(.Cells(zeile, colFields(allianzBOBSpalten.bobVerantw)), Excel.Range).Value) Then
                                            responsiblePerson = CStr(CType(.Cells(zeile, colFields(allianzBOBSpalten.bobVerantw)), Excel.Range).Value).Trim
                                        Else
                                            responsiblePerson = ""
                                        End If

                                    End If

                                Else
                                    boBName = ""
                                    description = ""
                                    allianzProjektNummer = ""
                                    responsiblePerson = ""
                                End If

                            ElseIf projektvorhaben.Contains(itemType) Then
                                If Not IsNothing(CType(.Cells(zeile, colFields(allianzBOBSpalten.scopename)), Excel.Range).Value) Then
                                    scopeName = CStr(CType(.Cells(zeile, colFields(allianzBOBSpalten.scopename)), Excel.Range).Value).Trim
                                    pName = scopeName

                                    ' jetzt description, ID  und verantwortlich rauslesen ... 
                                    If Not IsNothing(CType(.Cells(zeile, colFields(allianzBOBSpalten.scopedesc)), Excel.Range).Value) Then
                                        description = CStr(CType(.Cells(zeile, colFields(allianzBOBSpalten.scopedesc)), Excel.Range).Value).Trim
                                    Else
                                        description = ""
                                    End If

                                    If Not IsNothing(CType(.Cells(zeile, colFields(allianzBOBSpalten.scopeID)), Excel.Range).Value) Then
                                        allianzProjektNummer = CStr(CType(.Cells(zeile, colFields(allianzBOBSpalten.scopeID)), Excel.Range).Value).Trim
                                    Else
                                        allianzProjektNummer = ""
                                    End If

                                    If Not IsNothing(CType(.Cells(zeile, colFields(allianzBOBSpalten.scopeVerantw)), Excel.Range).Value) Then
                                        responsiblePerson = CStr(CType(.Cells(zeile, colFields(allianzBOBSpalten.scopeVerantw)), Excel.Range).Value).Trim
                                    Else
                                        responsiblePerson = ""
                                    End If

                                Else
                                    scopeName = ""
                                    description = ""
                                    allianzProjektNummer = ""
                                    responsiblePerson = ""
                                End If

                            End If


                            If boBName <> "" Or scopeName <> "" Then
                                ok = True
                            End If

                        Catch ex As Exception
                            pName = Nothing
                        End Try


                        If IsNothing(pName) Then
                            CType(.Cells(zeile, lastColumn), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color = awinSettings.AmpelGelb
                            CType(.Cells(zeile, lastColumn), Global.Microsoft.Office.Interop.Excel.Range).AddComment(Text:="Projekt-Name fehlt oder Fehler in PRojekt-Nummer ..")

                        ElseIf pName.Trim = nameTobeIgnored Then
                            CType(.Cells(zeile, lastColumn), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color = awinSettings.AmpelGelb
                            CType(.Cells(zeile, lastColumn), Global.Microsoft.Office.Interop.Excel.Range).AddComment(Text:="wird ignoriert ..")

                        ElseIf pName.Trim.Length < 2 Then

                            Try
                                CType(.Cells(zeile, lastColumn), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color = awinSettings.AmpelGelb
                                CType(.Cells(zeile, lastColumn), Global.Microsoft.Office.Interop.Excel.Range).AddComment(Text:="Projekt-Name muss mindestens 2 Buchstaben haben und eindeutig sein ..")
                            Catch ex As Exception

                            End Try

                        Else
                            custFields.Clear()

                            If Not isValidPVName(pName) Then
                                pName = makeValidProjectName(pName)
                            End If

                            Try
                                ' weitere Informationen auslesen 



                                ok = False

                                If projektvorhaben.Contains(itemType) Then
                                    ' ok weitermachen
                                    ok = True

                                    Try
                                        If Not IsNothing(CType(.Cells(zeile, colFields(allianzBOBSpalten.budget)), Excel.Range).Value) Then
                                            projVorhabensBudget = CDbl(CType(.Cells(zeile, colFields(allianzBOBSpalten.budget)), Excel.Range).Value)
                                        Else
                                            projVorhabensBudget = 0.0
                                        End If

                                    Catch ex As Exception
                                        projVorhabensBudget = 0.0
                                    End Try

                                Else
                                    ok = False
                                    ' jetzt muss geschaut werden, ob es sich um eine Programmlinie handelt, dann soll 
                                    ' ein neues Portfolio aufgemacht werden .. 
                                    If pgmlinie.Contains(itemType) Then
                                        ' die bisherige Constellation wegschreiben ...


                                        '  ur: 20211108: accelerate the load of an portfolio without summaryProject - calcUnionProject needs to m


                                        If Not IsNothing(current1program) Then
                                            ' ggf hier wieder rausnehmen ...

                                            If current1program.count > 0 Then
                                                If projectConstellations.Contains(current1program.constellationName) Then
                                                    projectConstellations.Remove(current1program.constellationName)
                                                End If

                                                createdPrograms = createdPrograms + 1
                                                projectConstellations.Add(current1program)

                                                '        ' tk 10.8.19 das wird jetzt wieder gemacht , aber nur um zu überprüfen ob Summe(POBs) <= lastProgramProj
                                                '        ' jetzt das union-Projekt erstellen ; 
                                                '        Dim unionProj As clsProjekt = calcUnionProject(current1program, True, Date.Now.Date.AddHours(23).AddMinutes(59), budget:=last1Budget)

                                                '        Try
                                                '            ' Test, ob das Budget auch ausreicht
                                                '            ' wenn nein, einfach Warning ausgeben 
                                                '            Dim tmpGesamtCost As Double = unionProj.getGesamtKostenBedarf.Sum
                                                '            If unionProj.Erloes - tmpGesamtCost < 0 Then

                                                '                Dim goOn As Boolean = True
                                                '                If unionProj.Erloes > 0 Then
                                                '                    goOn = (tmpGesamtCost - unionProj.Erloes) / unionProj.Erloes > 0.05
                                                '                End If

                                                '                If goOn Then
                                                '                    outPutLine = "Warnung: Budget-Überschreitung bei BOB: " & unionProj.name & " (Budget=" & unionProj.Erloes.ToString("#0.##") & ", Gesamtkosten=" & tmpGesamtCost.ToString("#0.##")
                                                '                    outputCollection.Add(outPutLine)

                                                '                    Dim logtxt(2) As String
                                                '                    logtxt(0) = "Budget-Überschreitung"
                                                '                    logtxt(1) = "Programmlinie"
                                                '                    logtxt(2) = unionProj.name
                                                '                    Dim values(2) As Double
                                                '                    values(0) = unionProj.Erloes
                                                '                    values(1) = tmpGesamtCost
                                                '                    If values(0) > 0 Then
                                                '                        values(2) = tmpGesamtCost / unionProj.Erloes
                                                '                    Else
                                                '                        values(2) = 9999999999
                                                '                    End If
                                                '                    Call logger(ptErrLevel.logWarning, "importAllianzBOBS", logtxt, values)
                                                '                End If

                                                '            End If

                                                '        Catch ex As Exception

                                                '        End Try

                                                '        Dim bobProj As clsProjekt = Nothing
                                                '        Dim bPKey As String = calcProjektKey(unionProj)

                                                '        If ImportProjekte.Containskey(bPKey) Then
                                                '            bobProj = ImportProjekte.getProject(bPKey)
                                                '            Dim updatedProj As clsProjekt = bobProj.updateProjectWithRessourcesFrom(unionProj)

                                                '            ' nur ersetzen , wenn es auch was zum Updaten gab
                                                '            If Not IsNothing(updatedProj) Then

                                                '                ImportProjekte.Remove(bPKey, updateCurrentConstellation:=False)
                                                '                ImportProjekte.Add(updatedProj, updateCurrentConstellation:=False)
                                                '                '' test
                                                '                Dim everythingOK As Boolean = testUProjandSingleProjs(current1program)
                                                '                If Not everythingOK Then

                                                '                    outPutLine = "Summary Projekt nicht identisch mit der Liste der Projekt-Vorhaben: " & current1program.constellationName
                                                '                    outputCollection.Add(outPutLine)

                                                '                    ReDim logmsg(2)
                                                '                    logmsg(0) = "Summary Projekt nicht identisch mit der Liste der Projekt-Vorhaben:"
                                                '                    logmsg(1) = ""
                                                '                    logmsg(2) = current1program.constellationName
                                                '                    Call logger(ptErrLevel.logError, "importAllianzBOBS", logmsg)

                                                '                    ' wieder zurücksetzen ... 
                                                '                    ImportProjekte.Remove(bPKey, updateCurrentConstellation:=False)
                                                '                    ImportProjekte.Add(bobProj, updateCurrentConstellation:=False)
                                                '                End If
                                                '                ' ende test
                                                '            Else
                                                '                ' nur dann was ausgeben, wenn unionproj auch Ressourcen hat ... 
                                                '                If unionProj.getAllPersonalKosten.Sum > 0 Then
                                                '                    outPutLine = "updatedProjekt mit Ressourcen fehlgeschlagen: " & bobProj.name
                                                '                    outputCollection.Add(outPutLine)

                                                '                    ReDim logmsg(2)
                                                '                    logmsg(0) = "updatedProjekt mit Ressourcen fehlgeschlagen: "
                                                '                    logmsg(1) = ""
                                                '                    logmsg(2) = bobProj.name
                                                '                    Call logger(ptErrLevel.logError, "importAllianzBOBS", logmsg)
                                                '                End If


                                                '            End If

                                                '        End If
                                            Else
                                                emptyPrograms = emptyPrograms + 1
                                            End If

                                        End If

                                        current1program = New clsConstellation(ptSortCriteria.customTF, itemType.ToString & " - " & pName)
                                        lfdNr1program = 2

                                        Try
                                            last1Budget = CDbl(CType(.Cells(zeile, colFields(allianzBOBSpalten.budget)), Excel.Range).Value)
                                        Catch ex As Exception
                                            last1Budget = 0.0
                                        End Try

                                        'With current1program
                                        '    .constellationName = itemType.ToString & " - " & pName
                                        'End With

                                        ' wenn jetzt als nächstes gleich wieder eine Programm-Linie kommt, dann muss dem Program als sein erstes und einziges Projekt 
                                        ' die Programm-Linie sein
                                        Dim programItemfound As Boolean = False
                                        Dim nextItemType As Integer
                                        Dim kidItemFound As Boolean = False
                                        Dim tmpZ As Integer = zeile + 1


                                        Do While tmpZ <= lastRow And Not kidItemFound And Not programItemfound
                                            Try
                                                nextItemType = CInt(CType(.Cells(tmpZ, colFields(allianzBOBSpalten.itemType)), Excel.Range).Value)
                                            Catch ex As Exception
                                                nextItemType = 0
                                            End Try

                                            kidItemFound = projektvorhaben.Contains(nextItemType)
                                            programItemfound = pgmlinie.Contains(nextItemType)
                                            tmpZ = tmpZ + 1

                                        Loop

                                        If Not kidItemFound And programItemfound Then
                                            ' jetzt wird sichergestellt, dass diese Programm-Linie jetzt als Projekt angelegt wird ..
                                            ok = True
                                            projVorhabensBudget = last1Budget
                                        ElseIf kidItemFound And tmpZ <= lastRow Then
                                            ok = True
                                            projVorhabensBudget = last1Budget
                                        End If

                                    End If
                                End If


                                If ok Then
                                    Try

                                        custFields = readCustomFieldsFromExcel(colCustomFields, 6, zeile, currentWS)


                                        ' lese , wieviel Prozent der Gesamtsumme jeweils auf die Release verteilt werden soll 
                                        For i As Integer = 0 To anzReleases - 1
                                            Try
                                                If IsNothing(CType(.Cells(zeile, colRelPrzStart + i), Excel.Range).Value) Then
                                                    relPrz(i) = 0.0
                                                Else
                                                    relPrz(i) = CDbl(CType(.Cells(zeile, colRelPrzStart + i), Excel.Range).Value)
                                                End If
                                            Catch ex As Exception
                                                relPrz(i) = 0.0
                                            End Try
                                        Next

                                        ' Plausibilitäts-Check - wenn es sich nicht auf 100% summiert, dann lieber alles auf die rootPhase verteilen und nichts auf die Release Phasen
                                        Dim a As Double = relPrz.Sum
                                        If relPrz.Sum > 0 Then
                                            If relPrz.Sum < 0.99 Or relPrz.Sum > 1.01 Then
                                                CType(.Cells(zeile, lastColumn), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color = awinSettings.AmpelGelb
                                                CType(.Cells(zeile, lastColumn), Global.Microsoft.Office.Interop.Excel.Range).AddComment(Text:="Prz-Sätze addieren nicht auf 100% ... alles in Projektphase ")
                                                If relPrz.Sum < 0.99 Then
                                                    outPutLine = pName & " Prozent-Sätze > 0 , aber < 1; Gesamt-Summe auf Gesamt-Projekt verteilt  .."
                                                Else
                                                    outPutLine = pName & " Prozent-Sätze > 1.0 , Gesamt-Summe auf Gesamt-Projekt verteilt  .."
                                                End If

                                                outputCollection.Add(outPutLine)

                                                Dim logtxt(2) As String
                                                logtxt(0) = "Prozent-Sätze > 0 , aber < 1; Gesamt-Summe auf Gesamt-Projekt verteilt  .. "
                                                logtxt(1) = ""
                                                logtxt(2) = pName

                                                Call logger(ptErrLevel.logError, "importAllianzBOBS", logtxt)

                                                ReDim relPrz(anzReleases - 1)
                                            End If
                                        End If


                                        For i As Integer = 0 To colTEroleNamesToConsider.Length - 1
                                            Try
                                                If IsNothing(CType(.Cells(zeile, colTEroleNamesToConsider(i)), Excel.Range).Value) Then
                                                    roleNeeds(i) = 0.0
                                                Else
                                                    roleNeeds(i) = 0.0

                                                    If Not IsNothing(CType(.Cells(zeile, colTEroleNamesToConsider(i)), Excel.Range).Value) Then

                                                        If IsNumeric(CType(.Cells(zeile, colTEroleNamesToConsider(i)), Excel.Range).Value) Then
                                                            Dim cellValue As Double = CDbl(CType(.Cells(zeile, colTEroleNamesToConsider(i)), Excel.Range).Value)
                                                            If cellValue > 0 Then

                                                                Dim tmpRoleDef As clsRollenDefinition = RoleDefinitions.getRoledef(TEroleNamesToConsider(i))

                                                                If Not IsNothing(tmpRoleDef) Then
                                                                    ' jetzt handelt es sich um T€ - Werte , das heisst die anzahl Manntage erreichnet sich aus value*1000/tagessatz
                                                                    Dim tagessatz As Double
                                                                    Dim tmpValue As Double
                                                                    Try
                                                                        tagessatz = RoleDefinitions.getRoledef(TEroleNamesToConsider(i)).tagessatzIntern
                                                                        If tagessatz = 0 Then
                                                                            tagessatz = 1000
                                                                            Call MsgBox("tagessatz = 0 ! Rolle " & TEroleNamesToConsider(i))
                                                                        End If

                                                                        tmpValue = CDbl(CType(.Cells(zeile, colTEroleNamesToConsider(i)), Excel.Range).Value) * 1000 / tagessatz

                                                                        If tmpValue >= 0 Then
                                                                            roleNeeds(i) = tmpValue
                                                                        Else
                                                                            roleNeeds(i) = 0.0
                                                                        End If
                                                                    Catch ex As Exception

                                                                    End Try
                                                                End If
                                                            End If
                                                        Else
                                                            roleNeeds(i) = 0.0
                                                        End If

                                                    Else
                                                        roleNeeds(i) = 0.0
                                                    End If

                                                End If
                                            Catch ex As Exception
                                                roleNeeds(i) = 0.0
                                            End Try

                                        Next


                                    Catch ex As Exception
                                        ok = False
                                    End Try

                                    ' jetzt werden noch weitere Infos eingelesen ..
                                    Try ' Ampelbeschreibung
                                        ampelText = ""
                                        If Not IsNothing(CType(.Cells(zeile, colFields(allianzBOBSpalten.budgetExpl)), Excel.Range).Value) Then
                                            ampelText = CStr(CType(.Cells(zeile, colFields(allianzBOBSpalten.budgetExpl)), Excel.Range).Value).Trim
                                        End If

                                    Catch ex As Exception
                                        ampelText = ""
                                    End Try

                                    ' Description wird in abhängigkeit von ItemType / satzartz eingelesen 

                                    Try ' Business Unit
                                        businessUnit = ""
                                        If Not IsNothing(CType(.Cells(zeile, colFields(allianzBOBSpalten.businessUnit)), Excel.Range).Value) Then
                                            businessUnit = CStr(CType(.Cells(zeile, colFields(allianzBOBSpalten.businessUnit)), Excel.Range).Value).Trim
                                        End If
                                    Catch ex As Exception
                                        businessUnit = ""
                                    End Try

                                    ' Projektleiter wird in Abhängigkeit von ItemType / Satzartz eingelesen 


                                    Try ' ProjektVorlage 
                                        vorlageName = "Std"
                                        If Not Projektvorlagen.Contains(vorlageName) Then
                                            vorlageName = Projektvorlagen.getProject(0).VorlagenName
                                        End If

                                        If Not IsNothing(CType(.Cells(zeile, colFields(allianzBOBSpalten.scopeTyp)), Excel.Range).Value) Then
                                            Dim tmpVorlagenName As String = CStr(CType(.Cells(zeile, colFields(allianzBOBSpalten.scopeTyp)), Excel.Range).Value).Trim
                                            If Projektvorlagen.Contains(tmpVorlagenName) Then
                                                vorlageName = tmpVorlagenName
                                            End If

                                        End If

                                    Catch ex As Exception
                                        vorlageName = "Run"
                                    End Try

                                    Try ' Budget
                                        budget = 0.0

                                        If Not IsNothing(CType(.Cells(zeile, colFields(allianzBOBSpalten.budget)), Excel.Range).Value) Then
                                            budget = CDbl(CType(.Cells(zeile, colFields(allianzBOBSpalten.budget)), Excel.Range).Value)
                                            If budget < 0 Then
                                                ' solche Projekte nicht einlesen 
                                                ok = False

                                                ' logfile und PRotokoll schreiben 
                                                Dim logtxt(2) As String
                                                logtxt(0) = "BOB / Scope mit negativem Budget wird nicht eingelesen: "
                                                If pgmlinie.Contains(itemType) Then
                                                    outPutLine = logtxt(0) & itemType.ToString & " - " & boBName & " : " & budget.ToString
                                                    logtxt(1) = boBName & " : " & budget.ToString
                                                Else
                                                    outPutLine = logtxt(0) & itemType.ToString & " - " & scopeName & " : " & budget.ToString
                                                    logtxt(1) = scopeName & " : " & budget.ToString
                                                End If

                                                outputCollection.Add(outPutLine)
                                                logtxt(2) = pName

                                                Call logger(ptErrLevel.logError, "importAllianzBOBS", logtxt)

                                            End If
                                        End If

                                    Catch ex As Exception
                                        budget = 0.0
                                    End Try

                                    Try ' Budget beauftragt
                                        budgetBeauftragt = 0.0
                                        If Not IsNothing(CType(.Cells(zeile, colFields(allianzBOBSpalten.budgetBeauftragt)), Excel.Range).Value) Then
                                            budgetBeauftragt = CDbl(CType(.Cells(zeile, colFields(allianzBOBSpalten.budgetBeauftragt)), Excel.Range).Value)
                                            If budgetBeauftragt < 0 Then
                                                budgetBeauftragt = 0.0
                                            End If
                                        End If
                                    Catch ex As Exception
                                        budgetBeauftragt = 0.0
                                    End Try


                                    Try ' Status

                                        If budget = budgetBeauftragt And budget > 0 Then
                                            'ur: 211202: 
                                            'allianzStatus = ProjektStatus(PTProjektStati.beauftragt)
                                            allianzStatus = VProjectStatus(PTVPStati.ordered)
                                        Else
                                            'ur: 211202: 
                                            'allianzStatus = ProjektStatus(PTProjektStati.geplant)
                                            allianzStatus = VProjectStatus(PTVPStati.initialized)
                                        End If

                                    Catch ex As Exception
                                        'ur: 211202: 
                                        'allianzStatus = ProjektStatus(PTProjektStati.geplant)
                                        allianzStatus = VProjectStatus(PTVPStati.initialized)
                                    End Try

                                    Try ' im richtigen Zeitfenster ?
                                        Dim valid As Boolean = True
                                        If projektvorhaben.Contains(itemType) Then


                                            If Not IsNothing(CType(.Cells(zeile, colFields(allianzBOBSpalten.scopeStart)), Excel.Range).Value) Then
                                                pStartDatum = CDate(CType(.Cells(zeile, colFields(allianzBOBSpalten.scopeStart)), Excel.Range).Value)
                                            Else
                                                pStartDatum = startdate
                                            End If



                                            If Not IsNothing(CType(.Cells(zeile, colFields(allianzBOBSpalten.scopeEnd)), Excel.Range).Value) Then
                                                pEndDatum = CDate(CType(.Cells(zeile, colFields(allianzBOBSpalten.scopeEnd)), Excel.Range).Value)
                                            Else
                                                pEndDatum = endDate
                                            End If

                                            If DateDiff(DateInterval.Month, pStartDatum, startdate) > 0 Or DateDiff(DateInterval.Month, pStartDatum, endDate) <= 0 Then
                                                valid = False
                                            End If

                                            If DateDiff(DateInterval.Month, pEndDatum, endDate) < 0 Or DateDiff(DateInterval.Month, pEndDatum, startdate) >= 0 Then
                                                valid = False
                                            End If

                                            If DateDiff(DateInterval.Month, pStartDatum, pEndDatum) < 0 Then
                                                valid = False
                                            End If

                                            If Not valid Then
                                                ok = False

                                                ' logfile und PRotokoll schreiben 
                                                Dim logtxt(2) As String
                                                logtxt(0) = "Scope ist nicht im aktuellen Zeitfenster, wird nicht eingelesen: "

                                                outPutLine = logtxt(0) & itemType.ToString & " - " & scopeName & " : " & pStartDatum.ToShortDateString & " - " & pEndDatum.ToShortDateString
                                                logtxt(1) = scopeName & " : " & pStartDatum.ToShortDateString & " - " & pEndDatum.ToShortDateString

                                                outputCollection.Add(outPutLine)

                                                Call logger(ptErrLevel.logError, "importAllianzBOBS", logtxt)


                                            End If
                                        Else
                                            pStartDatum = startdate
                                            pEndDatum = endDate
                                        End If


                                    Catch ex As Exception
                                        pStartDatum = startdate
                                        pEndDatum = endDate
                                    End Try

                                End If



                            Catch ex As Exception
                                Call MsgBox("Fehler bei Informationen auslesen: Projekt " & pName)
                                ok = False
                            End Try



                            If ok Then


                                'Projekt anlegen ,Verschiebung um 
                                Dim hproj As clsProjekt = Nothing

                                ' #####################################################################
                                ' Erstellen des Projekts nach den Angaben aus der Batch-Datei 
                                '
                                pName = itemType.ToString & " - " & pName
                                Dim combinedName As Boolean = False
                                If projektvorhaben.Contains(itemType) Then
                                    combinedName = True
                                End If
                                ' lege ein Allianz IT - Projekt an
                                hproj = erstelleProjektausParametern(pName, variantName, vorlageName, pStartDatum, pEndDatum, budget, sFit, risk, allianzProjektNummer,
                                                                     description, custFields, businessUnit, responsiblePerson, allianzStatus,
                                                                     zeile, realRoleNamesToConsider, roleNeeds, Nothing, Nothing, phNames, relPrz, combinedName)

                                ' tk 21.7.19 es wird ein Summary Projekt für die Programm-Linie angelegt  
                                If pgmlinie.Contains(itemType) Then
                                    'pName = itemType.ToString & " - " & pName
                                    'hproj.name = pName
                                    hproj.projectType = ptPRPFType.portfolio
                                End If

                                Try
                                    ' Test, ob das Budget auch ausreicht
                                    ' wenn nein, einfach Warning ausgeben 
                                    Dim tmpGesamtCost As Double = hproj.getGesamtKostenBedarf.Sum
                                    If hproj.Erloes - tmpGesamtCost < 0 Then

                                        Dim goOn As Boolean = True
                                        If hproj.Erloes > 0 Then
                                            goOn = (tmpGesamtCost - hproj.Erloes) / hproj.Erloes > 0.05
                                        End If

                                        If goOn Then
                                            Dim logtxt(2) As String
                                            logtxt(0) = "Budget-Überschreitung"
                                            If pgmlinie.Contains(itemType) Then
                                                outPutLine = "Warnung: Budget-Überschreitung bei " & pName & " (Budget=" & hproj.Erloes.ToString("#0.##") & ", Gesamtkosten=" & tmpGesamtCost.ToString("#0.##")
                                                logtxt(1) = "Programm-Linie"
                                            Else
                                                outPutLine = "Warnung: Budget-Überschreitung bei " & pName & " (Budget=" & hproj.Erloes.ToString("#0.##") & ", Gesamtkosten=" & tmpGesamtCost.ToString("#0.##")
                                                logtxt(1) = "Projekt"
                                            End If

                                            outputCollection.Add(outPutLine)
                                            logtxt(2) = pName

                                            Dim values(2) As Double
                                            values(0) = hproj.Erloes
                                            values(1) = tmpGesamtCost
                                            If values(0) > 0 Then
                                                values(2) = tmpGesamtCost / hproj.Erloes
                                            Else
                                                values(2) = 9999999999
                                            End If
                                            Call logger(ptErrLevel.logWarning, "importAllianzBOBS", logtxt, values)
                                        End If

                                    End If

                                Catch ex As Exception

                                End Try

                                ' Test tk 
                                Try
                                    ReDim logmsg(2)
                                    logmsg(0) = "Importiert: "
                                    logmsg(1) = ""
                                    logmsg(2) = pName

                                    For ix As Integer = 1 To realRoleNamesToConsider.Count

                                        Dim tmpRollenName As String = realRoleNamesToConsider(ix - 1)
                                        Dim sollBedarf As Double = roleNeeds(ix - 1)


                                        Dim tmpCollection As New Collection
                                        tmpCollection.Add(tmpRollenName)
                                        Dim istBedarf As Double = hproj.getRessourcenBedarf(tmpRollenName,
                                                                                            inclSubRoles:=True).Sum

                                        If Math.Abs(sollBedarf - istBedarf) > 0.001 Then
                                            outPutLine = "Differenz bei " & pName & ", " & tmpRollenName & ": " & Math.Abs(sollBedarf - istBedarf).ToString("#0.##")
                                            outputCollection.Add(outPutLine)
                                        End If

                                    Next

                                    Dim sollBedarfGesamt As Double = roleNeeds.Sum
                                    Dim istBedarfGesamt As Double = hproj.getAlleRessourcen.Sum

                                    If Math.Abs(sollBedarfGesamt - istBedarfGesamt) > 0.001 Then
                                        outPutLine = "Gesamt Differenz bei " & pName & ": " & Math.Abs(sollBedarfGesamt - istBedarfGesamt).ToString("#0.##")
                                        outputCollection.Add(outPutLine)
                                    End If

                                    If awinSettings.visboDebug Then
                                        Call logger(ptErrLevel.logWarning, "importAllianzBOBS", logmsg, roleNeeds)
                                    End If


                                Catch ex As Exception

                                End Try



                                ' Ende Test tk 

                                If Not IsNothing(hproj) Then


                                    ' jetzt ist alles so weit ok 
                                    Dim pkey As String = ""
                                    If Not IsNothing(hproj) Then
                                        Try
                                            pkey = calcProjektKey(hproj)

                                            If ImportProjekte.Containskey(pkey) Then
                                                outPutLine = "Name existiert mehrfach: " & pName
                                                outputCollection.Add(outPutLine)

                                                Dim logtxt(2) As String
                                                logtxt(0) = "Name existiert mehrfach: "
                                                logtxt(1) = ""
                                                logtxt(2) = pName
                                                Call logger(ptErrLevel.logWarning, "importAllianzBOBS", logtxt)
                                            Else
                                                ImportProjekte.Add(hproj, False)
                                                If projektvorhaben.Contains(itemType) Then
                                                    createdProjects = createdProjects + 1


                                                    ' jetzt soll das in die Constellation 
                                                    Dim cItem As New clsConstellationItem
                                                    With cItem
                                                        .projectName = hproj.name
                                                        .variantName = hproj.variantName
                                                        .show = True
                                                        .projectTyp = CType(hproj.projectType, ptPRPFType).ToString
                                                        .zeile = lfdNr1program
                                                    End With

                                                    current1program.add(cItem)
                                                    lfdNr1program = lfdNr1program + 1
                                                End If

                                            End If

                                        Catch ex As Exception
                                            outPutLine = "Fehler bei " & pName & vbLf & "Error: " & ex.Message
                                            outputCollection.Add(outPutLine)
                                        End Try

                                    End If


                                Else
                                    ok = False
                                    If pgmlinie.Contains(itemType) Then
                                        outPutLine = "Fehler beim Erzeugen der Programm-Linie " & pName
                                    Else
                                        outPutLine = "Fehler beim Erzeugen des Projektes " & pName
                                    End If

                                    outputCollection.Add(outPutLine)
                                End If

                            End If

                        End If

                        geleseneProjekte = geleseneProjekte + 1

                    End If

                    zeile = zeile + 1

                End While

                ' jetzt die letzte ggf vorkommende Constellation aufnehmen 
                If Not IsNothing(current1program) Then

                    If current1program.count > 0 Then
                        ' ggf aus der Liste aller Constellations wieder rausnehmen 

                        If projectConstellations.Contains(current1program.constellationName) Then
                            projectConstellations.Remove(current1program.constellationName)
                        End If

                        createdPrograms = createdPrograms + 1
                        projectConstellations.Add(current1program)

                        ' ur: 20211108: accelerate the load of an portfolio without summaryProject - calcUnionProject needs to m

                        ' tk 10.8.19 das wird jetzt wieder gemacht , aber nur um zu überprüfen ob Summe(POBs) <= lastProgramProj
                        ' jetzt das union-Projekt erstellen 


                        ' Dim unionProj As clsProjekt = calcUnionProject(current1program, True, Date.Now.Date.AddHours(23).AddMinutes(59), budget:=last1Budget)

                        'Try
                        '    ' Test, ob das Budget auch ausreicht
                        '    ' wenn nein, einfach Warning ausgeben 
                        '    Dim tmpGesamtCost As Double = unionProj.getGesamtKostenBedarf.Sum
                        '    If unionProj.Erloes - tmpGesamtCost < 0 Then
                        '        Dim goOn As Boolean = True
                        '        If unionProj.Erloes > 0 Then
                        '            goOn = (tmpGesamtCost - unionProj.Erloes) / unionProj.Erloes > 0.05
                        '        End If

                        '        If goOn Then
                        '            outPutLine = "Warnung: Budget-Überschreitung bei BOB: " & unionProj.name & " (Budget=" & unionProj.Erloes.ToString("#0.##") & ", Gesamtkosten=" & tmpGesamtCost.ToString("#0.##")
                        '            outputCollection.Add(outPutLine)

                        '            Dim logtxt(2) As String
                        '            logtxt(0) = "Budget-Überschreitung"
                        '            logtxt(1) = "Programmlinie"
                        '            logtxt(2) = unionProj.name
                        '            Dim values(2) As Double
                        '            values(0) = unionProj.Erloes
                        '            values(1) = tmpGesamtCost
                        '            If values(0) > 0 Then
                        '                values(2) = tmpGesamtCost / unionProj.Erloes
                        '            Else
                        '                values(2) = 9999999999
                        '            End If
                        '            Call logger(ptErrLevel.logWarning, "importAllianzBOBS", logtxt, values)
                        '        End If

                        '    End If

                        'Catch ex As Exception

                        'End Try

                        'Dim bobProj As clsProjekt = Nothing
                        'Dim bPKey As String = calcProjektKey(unionProj)

                        'If ImportProjekte.Containskey(bPKey) Then
                        '    bobProj = ImportProjekte.getProject(bPKey)
                        '    Dim updatedProj As clsProjekt = bobProj.updateProjectWithRessourcesFrom(unionProj)

                        '    ' nur ersetzen , wenn es auch was zum Updaten gab
                        '    If Not IsNothing(updatedProj) Then

                        '        ImportProjekte.Remove(bPKey, updateCurrentConstellation:=False)
                        '        ImportProjekte.Add(updatedProj, updateCurrentConstellation:=False)
                        '        '' test
                        '        Dim everythingOK As Boolean = testUProjandSingleProjs(current1program)
                        '        If Not everythingOK Then

                        '            outPutLine = "Summary Projekt nicht identisch mit der Liste der Projekt-Vorhaben: " & current1program.constellationName
                        '            outputCollection.Add(outPutLine)

                        '            ReDim logmsg(2)
                        '            logmsg(0) = "Summary Projekt nicht identisch mit der Liste der Projekt-Vorhaben:"
                        '            logmsg(1) = ""
                        '            logmsg(2) = current1program.constellationName
                        '            Call logger(ptErrLevel.logWarning, "importAllianzBOBS", logmsg)

                        '            ' wieder zurücksetzen ... 
                        '            ImportProjekte.Remove(bPKey, updateCurrentConstellation:=False)
                        '            ImportProjekte.Add(bobProj, updateCurrentConstellation:=False)
                        '        End If
                        '        ' ende test
                        '    Else
                        '        If unionProj.getAllPersonalKosten.Sum > 0 Then
                        '            outPutLine = "updatedProjekt mit Ressourcen fehlgeschlagen: " & bobProj.name
                        '            outputCollection.Add(outPutLine)

                        '            ReDim logmsg(2)
                        '            logmsg(0) = "updatedProjekt mit Ressourcen fehlgeschlagen: "
                        '            logmsg(1) = ""
                        '            logmsg(2) = bobProj.name
                        '            Call logger(ptErrLevel.logWarning, "importAllianzBOBS", logmsg)
                        '        End If

                        '    End If

                        'End If




                    Else
                        emptyPrograms = emptyPrograms + 1
                    End If


                End If

            End With
        Catch ex As Exception

            Throw New Exception("Fehler in Import-Datei: " & ex.Message)

        End Try


        If outputCollection.Count > 0 Then
            Call showOutPut(outputCollection, "Import Type 1", "")
        End If

        If emptyPrograms = 0 Then
            Call MsgBox("Zeilen gelesen: " & geleseneProjekte & vbLf &
                    "Scopes erzeugt: " & createdProjects & vbLf &
                    "BOBs erzeugt: " & createdPrograms & vbLf &
                    "insgesamt importiert: " & ImportProjekte.Count)
        Else
            Call MsgBox("Zeilen gelesen: " & geleseneProjekte & vbLf &
                    "Scopes erzeugt: " & createdProjects & vbLf &
                    "BOBs erzeugt: " & createdPrograms & vbLf &
                    "BOBs nicht erzeugt, weil leer: " & emptyPrograms & vbLf &
                    "insgesamt importiert: " & ImportProjekte.Count)
        End If


    End Sub



    ''' <summary>
    ''' erzeugt die Projekte, die in der Batch-Datei angegeben sind
    ''' stellt sie in ImportProjekte 
    ''' erstellt ein Szenario mit Namen der Batch-Datei; die Sortierung erfolgt über die Reihenfolge in der Batch-Datei 
    ''' das wird sichergestellt über Eintrag der tfzeile in hproj ... 
    ''' </summary>
    ''' <remarks></remarks>
    Sub importAllianzType1(ByVal startdate As Date, ByVal endDate As Date)

        Dim zeile As Integer, spalte As Integer

        Dim importType As ptImportTypen

        Dim tfZeile As Integer = 2

        Dim pName As String = ""
        Dim variantName As String = ""
        Dim custFields As New Collection
        Dim description As String = ""
        Dim responsiblePerson As String = ""
        Dim sFit As Double = 5.0
        Dim risk As Double = 5.0
        Dim budget As Double = 0.0
        Dim businessUnit As String = ""
        Dim allianzProjektNummer As String = ""
        Dim allianzStatus As String = ""
        Dim ampelText As String
        Dim projVorhabensBudget As Double = 0.0

        Dim logmsg() As String

        Dim programName As String = ""
        Dim current1program As clsConstellation = Nothing

        Dim last1Budget As Double = 0.0
        Dim lfdNr1program As Integer = 2

        ' nimmt die vollen Namen der 
        Dim fullNameListe1 As New SortedList(Of String, String)

        Dim createdProjects As Integer = 0
        Dim createdPrograms As Integer = 0
        Dim emptyPrograms As Integer = 0



        Dim lastRow As Integer
        Dim lastColumn As Integer
        Dim geleseneProjekte As Integer
        Dim ok As Boolean = False

        ' für den Output 
        Dim outputFenster As New frmOutputWindow
        Dim outputCollection As New Collection
        Dim outPutLine As String = ""

        ' Standard-Definition
        Dim anzReleases As Integer = 5


        Dim vorlageName As String = "Rel"
        If awinSettings.databaseName.EndsWith("20") Then
            vorlageName = "Rel20"
        End If

        Try
            anzReleases = Projektvorlagen.getProject(vorlageName).CountPhases - 1
        Catch ex As Exception

        End Try


        ' enthält die prozentualen Anteile in den Releases 
        Dim relPrz() As Double
        ReDim relPrz(anzReleases - 1)

        ' Projekt-Eintrag/Zeile, die in der Excel Datei ignoriert werden soll 
        Dim nameTobeIgnored As String = "xxx"

        ' enthält die Phasen Namen
        Dim phNames() As String
        ReDim phNames(anzReleases - 1)

        ' enthält die Spalten-Nummer, ab der die Release Phasen Anteile stehen 
        Dim colRelPrzStart As Integer

        ' enthält die Info, welche Rollen-Namen berücksichtigt werden sollen 
        Dim roleNamesToConsider() As String = Nothing

        ' enthält die Spalten-Nummern, wo die einzelnen Rollen-Namen zu finden sind
        Dim colRoleNamesToConsider() As Integer = Nothing

        ' enthält, wieviel Manntage von dieser Rolle insgesamt benötigt werden 
        Dim roleNeeds() As Double = Nothing

        ' enthält die Info, welche Rollen-Namen berücksichtigt werden sollen 
        Dim TEroleNamesToConsider() As String = Nothing

        ' enthält die Spalten-Nummern, wo die einzelnen Rollen-Namen zu finden sind
        Dim colTEroleNamesToConsider() As Integer = Nothing

        ' enthält, wieviel Manntage von dieser Rolle insgesamt benötigt werden 
        Dim costNeeds() As Double = Nothing

        ' enthält die Spalten, wo die einzelnen Felder stehen , korreliert mit der Enum allianzSpalten
        Dim colFields() As Integer

        Dim firstZeile As Excel.Range

        Dim enumAllianzCount As Integer = [Enum].GetNames(GetType(allianzSpalten)).Length
        ReDim colFields(enumAllianzCount)



        spalte = 1
        geleseneProjekte = 0

        ' jetzt werden die Phase-Names besetzt
        Try
            For i = 1 To anzReleases
                phNames(i - 1) = Projektvorlagen.getProject(vorlageName).getPhase(i + 1).name
            Next
        Catch ex As Exception
            Call MsgBox("Probleme mit Vorlage " & vorlageName)
            Exit Sub
        End Try


        Try


            Dim currentWS As Excel.Worksheet = bestimmeWsAndImporttype(importType)

            If IsNothing(currentWS) Then
                Call MsgBox("Import File nicht erkannt - bitte " & visboImportKennung & "-Feld in Excel-Datei eintragen!")
            ElseIf (importType <> ptImportTypen.allianzMassImport1 And importType <> ptImportTypen.allianzMassImport2) Then
                Call MsgBox("keine Allianz-Projektliste: " & visboImportKennung & "muss Wert 5 oder 6 haben!")
                Exit Sub
            End If

            Dim isOldAllianzImport As Boolean = (importType = ptImportTypen.allianzMassImport1)


            With currentWS



                ' jetzt wird festgelegt, ab wo die relativen Verteilungs-Werte für die Releases stehen 
                If isOldAllianzImport Then
                    colRelPrzStart = .Range("AI1").Column
                    firstZeile = CType(.Rows(2), Excel.Range)
                    zeile = 3
                Else
                    colRelPrzStart = .Range("AB1").Column
                    firstZeile = CType(.Rows(3), Excel.Range)
                    zeile = 5
                End If


                ' damit werden die Arrays besetzt, welche Rollen gesucht sind und in welchen Spalten die Angaben dazu zu finden sind ... 
                Call setAllianzImportArrays(roleNamesToConsider, colRoleNamesToConsider,
                                            TEroleNamesToConsider, colTEroleNamesToConsider,
                                            currentWS, importType)


                'lastColumn = firstZeile.End(XlDirection.xlToLeft).Column

                lastColumn = CType(.Cells(1, 3000), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlToLeft).Column

                If isOldAllianzImport Then
                    lastRow = CType(.Cells(5000, "B"), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlUp).Row
                Else
                    lastRow = CType(.Cells(5000, "A"), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlUp).Row
                End If



                ' um die CustomFields lesen zu können ... 
                Dim colCustomFields(3) As Integer

                If isOldAllianzImport Then
                    ' T-BWLA
                    colCustomFields(0) = CInt(CType(.Range("A1"), Excel.Range).Column)
                    ' PGML
                    colCustomFields(1) = CInt(CType(.Range("B1"), Excel.Range).Column)
                    ' POB
                    colCustomFields(2) = CInt(CType(.Range("C1"), Excel.Range).Column)
                    ' Key Cluster
                    colCustomFields(3) = CInt(CType(.Range("D1"), Excel.Range).Column)

                    ' jetzt die Spalten bestimmen, wo die Werte stehen
                    Try
                        colFields(allianzSpalten.Name) = CType(.Range("H1"), Excel.Range).Column
                        colFields(allianzSpalten.itemType) = CType(.Range("G1"), Excel.Range).Column
                        colFields(allianzSpalten.AmpelText) = CType(.Range("U1"), Excel.Range).Column
                        colFields(allianzSpalten.BusinessUnit) = CType(.Range("AB1"), Excel.Range).Column
                        colFields(allianzSpalten.Responsible) = CType(.Range("AC1"), Excel.Range).Column
                        colFields(allianzSpalten.Projektnummer) = CType(.Range("AD1"), Excel.Range).Column
                        colFields(allianzSpalten.Status) = CType(.Range("AF1"), Excel.Range).Column
                        colFields(allianzSpalten.Budget) = CType(.Range("M1"), Excel.Range).Column
                        colFields(allianzSpalten.pvBudget) = CType(.Range("N1"), Excel.Range).Column
                    Catch ex As Exception
                        Dim errmsg As String = "fehlerhafte Range Definition ..."
                        Throw New ArgumentException(errmsg)
                    End Try
                Else

                    ' T-BWLA
                    colCustomFields(0) = CInt(CType(.Range("D1"), Excel.Range).Column)
                    ' PGML
                    colCustomFields(1) = CInt(CType(.Range("E1"), Excel.Range).Column)
                    ' POB
                    colCustomFields(2) = CInt(CType(.Range("F1"), Excel.Range).Column)
                    ' Key Cluster
                    colCustomFields(3) = CInt(CType(.Range("H1"), Excel.Range).Column)

                    ' jetzt die Spalten bestimmen, wo die Werte stehen
                    Try
                        colFields(allianzSpalten.Name) = CType(.Range("J1"), Excel.Range).Column
                        colFields(allianzSpalten.itemType) = CType(.Range("A1"), Excel.Range).Column
                        colFields(allianzSpalten.AmpelText) = CType(.Range("L1"), Excel.Range).Column
                        colFields(allianzSpalten.BusinessUnit) = CType(.Range("Y1"), Excel.Range).Column
                        colFields(allianzSpalten.Responsible) = CType(.Range("Z1"), Excel.Range).Column
                        colFields(allianzSpalten.Projektnummer) = CType(.Range("K1"), Excel.Range).Column
                        colFields(allianzSpalten.Status) = CType(.Range("EA1"), Excel.Range).Column
                        colFields(allianzSpalten.Budget) = CType(.Range("M1"), Excel.Range).Column
                        colFields(allianzSpalten.pvBudget) = CType(.Range("O1"), Excel.Range).Column
                    Catch ex As Exception
                        Dim errmsg As String = "fehlerhafte Range Definition ..."
                        Throw New ArgumentException(errmsg)
                    End Try

                End If

                Dim realRoleNamesToConsider() As String = Nothing
                If isOldAllianzImport Then
                    realRoleNamesToConsider = roleNamesToConsider
                Else
                    ReDim realRoleNamesToConsider(roleNamesToConsider.Length + TEroleNamesToConsider.Length - 1)
                    For i As Integer = 0 To roleNamesToConsider.Length - 1
                        realRoleNamesToConsider(i) = roleNamesToConsider(i)
                    Next
                    Dim i_offset As Integer = roleNamesToConsider.Length

                    For i As Integer = 0 To TEroleNamesToConsider.Length - 1
                        realRoleNamesToConsider(i + i_offset) = TEroleNamesToConsider(i)
                    Next
                End If

                ' tk Test Logfile schreiben ...
                If awinSettings.visboDebug Then
                    Try
                        ReDim logmsg(realRoleNamesToConsider.Count)
                        logmsg(0) = ""
                        For ix As Integer = 1 To realRoleNamesToConsider.Count
                            logmsg(ix) = realRoleNamesToConsider(ix - 1)
                        Next
                        Call logger(ptErrLevel.logDebug, "importAllianzType1", logmsg)
                    Catch ex As Exception

                    End Try
                End If


                ' jetzt die zugelassenen Werte für 
                Dim pgmlinie() As Integer
                Dim projektvorhaben() As Integer

                If isOldAllianzImport Then
                    ReDim pgmlinie(0)
                    ReDim projektvorhaben(0)
                    pgmlinie(0) = 1
                    projektvorhaben(0) = 4
                Else
                    ReDim pgmlinie(0)
                    ' nach Relco mit Rupi am 8.8.19 so entschieden
                    ReDim projektvorhaben(0)

                    pgmlinie(0) = 2
                    projektvorhaben(0) = 4
                    'projektvorhaben(1) = 6
                    'projektvorhaben(2) = 7
                End If

                ' jetzt müssen die Dimensionen gesetzt werden 
                Dim tmpLen As Integer = roleNamesToConsider.Length

                If Not IsNothing(roleNamesToConsider) Then

                    If importType = ptImportTypen.allianzMassImport2 Then

                        If Not IsNothing(TEroleNamesToConsider) Then
                            tmpLen = tmpLen + TEroleNamesToConsider.Length
                        End If

                    End If

                    ReDim roleNeeds(tmpLen - 1)

                ElseIf Not IsNothing(TEroleNamesToConsider) Then

                    tmpLen = TEroleNamesToConsider.Length
                    ReDim roleNeeds(tmpLen - 1)

                End If


                While zeile <= lastRow

                    ' Werte zurücksetzen ..
                    ReDim roleNeeds(tmpLen - 1)

                    ok = False

                    ' Kommentare zurücksetzen ...
                    Try
                        CType(.Range(.Cells(zeile, 1), .Cells(zeile, lastColumn)), Global.Microsoft.Office.Interop.Excel.Range).ClearComments()
                    Catch ex As Exception

                    End Try


                    ' lese den Projekt-Namen
                    Try
                        If Not IsNothing(CType(.Cells(zeile, colFields(allianzSpalten.Name)), Excel.Range).Value) Then
                            pName = CStr(CType(.Cells(zeile, colFields(allianzSpalten.Name)), Excel.Range).Value).Trim
                        Else
                            pName = Nothing
                        End If
                        ok = True
                    Catch ex As Exception
                        pName = Nothing
                    End Try


                    If IsNothing(pName) Then
                        CType(.Cells(zeile, lastColumn), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color = awinSettings.AmpelGelb
                        CType(.Cells(zeile, lastColumn), Global.Microsoft.Office.Interop.Excel.Range).AddComment(Text:="Projekt-Name fehlt ..")

                    ElseIf pName.Trim = nameTobeIgnored Then
                        CType(.Cells(zeile, lastColumn), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color = awinSettings.AmpelGelb
                        CType(.Cells(zeile, lastColumn), Global.Microsoft.Office.Interop.Excel.Range).AddComment(Text:="wird ignoriert ..")

                    ElseIf pName.Trim.Length < 2 Then

                        Try
                            CType(.Cells(zeile, lastColumn), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color = awinSettings.AmpelGelb
                            CType(.Cells(zeile, lastColumn), Global.Microsoft.Office.Interop.Excel.Range).AddComment(Text:="Projekt-Name muss mindestens 2 Buchstaben haben und eindeutig sein ..")
                        Catch ex As Exception

                        End Try

                    Else

                        Dim itemType As Integer

                        custFields.Clear()
                        description = pName

                        If Not isValidPVName(pName) Then
                            pName = makeValidProjectName(pName)
                        End If

                        Try
                            ' weitere Informationen auslesen 

                            Try
                                itemType = CInt(CType(.Cells(zeile, colFields(allianzSpalten.itemType)), Excel.Range).Value)
                            Catch ex As Exception
                                itemType = 0
                            End Try

                            ok = False

                            If projektvorhaben.Contains(itemType) Then
                                ' ok weitermachen
                                ok = True

                                Try
                                    projVorhabensBudget = CDbl(CType(.Cells(zeile, colFields(allianzSpalten.pvBudget)), Excel.Range).Value)
                                Catch ex As Exception
                                    projVorhabensBudget = 0.0
                                End Try

                            Else
                                ok = False
                                ' jetzt muss geschaut werden, ob es sich um eine Programmlinie handelt, dann soll 
                                ' ein neues Portfolio aufgemacht werden .. 
                                If pgmlinie.Contains(itemType) Then
                                    ' die bisherige Constellation wegschreiben ...


                                    If Not IsNothing(current1program) Then
                                        ' ggf hier wieder rausnehmen ...

                                        If current1program.count > 0 Then
                                            If projectConstellations.Contains(current1program.constellationName) Then
                                                projectConstellations.Remove(current1program.constellationName)
                                            End If

                                            createdPrograms = createdPrograms + 1
                                            projectConstellations.Add(current1program)

                                            ' 
                                            ' tk 10.8.19 das wird jetzt wieder gemacht , aber nur um zu überprüfen ob Summe(POBs) <= lastProgramProj
                                            ' jetzt das union-Projekt erstellen ;


                                            '  ur: 20211108: accelerate the load of an portfolio without summaryProject - calcUnionProject needs to m

                                            'Dim unionProj As clsProjekt = calcUnionProject(current1program, True, Date.Now.Date.AddHours(23).AddMinutes(59), budget:=last1Budget)

                                            'Try
                                            '    ' Test, ob das Budget auch ausreicht
                                            '    ' wenn nein, einfach Warning ausgeben 
                                            '    Dim tmpGesamtCost As Double = unionProj.getGesamtKostenBedarf.Sum
                                            '    If unionProj.Erloes - tmpGesamtCost < 0 Then

                                            '        Dim goOn As Boolean = True
                                            '        If unionProj.Erloes > 0 Then
                                            '            goOn = (tmpGesamtCost - unionProj.Erloes) / unionProj.Erloes > 0.05
                                            '        End If

                                            '        If goOn Then
                                            '            outPutLine = "Warnung: Budget-Überschreitung bei Programmlinie" & unionProj.name & " (Budget=" & unionProj.Erloes.ToString("#0.##") & ", Gesamtkosten=" & tmpGesamtCost.ToString("#0.##")
                                            '            outputCollection.Add(outPutLine)

                                            '            Dim logtxt(2) As String
                                            '            logtxt(0) = "Budget-Überschreitung"
                                            '            logtxt(1) = "Programmlinie"
                                            '            logtxt(2) = unionProj.name
                                            '            Dim values(2) As Double
                                            '            values(0) = unionProj.Erloes
                                            '            values(1) = tmpGesamtCost
                                            '            If values(0) > 0 Then
                                            '                values(2) = tmpGesamtCost / unionProj.Erloes
                                            '            Else
                                            '                values(2) = 9999999999
                                            '            End If
                                            '            Call logfileSchreiben(logtxt, values)
                                            '        End If

                                            '    End If

                                            'Catch ex As Exception

                                            'End Try

                                            ' Status gleich auf 1: beauftragt setzen 
                                            'unionProj.Status = ProjektStatus(PTProjektStati.beauftragt)

                                            'If ImportProjekte.Containskey(calcProjektKey(unionProj)) Then
                                            '    ImportProjekte.Remove(calcProjektKey(unionProj), updateCurrentConstellation:=False)
                                            'End If

                                            'ImportProjekte.Add(unionProj, updateCurrentConstellation:=False)
                                            '' test
                                            Dim everythingOK As Boolean = testUProjandSingleProjs(current1program)
                                            If Not everythingOK Then
                                                outPutLine = "Summary Projekt nicht identisch mit der Liste der Projekt-Vorhaben: " & current1program.constellationName
                                                outputCollection.Add(outPutLine)

                                                ReDim logmsg(2)
                                                logmsg(0) = "Summary Projekt nicht identisch mit der Liste der Projekt-Vorhaben:"
                                                logmsg(1) = ""
                                                logmsg(2) = current1program.constellationName
                                                Call logger(ptErrLevel.logError, "importAllianzBOBS", logmsg)
                                            End If
                                            '' ende tk Änderung 21.7.19 
                                        Else
                                            emptyPrograms = emptyPrograms + 1
                                        End If

                                    End If

                                    current1program = New clsConstellation(ptSortCriteria.customTF, itemType.ToString & " - " & pName)
                                    lfdNr1program = 2

                                    Try
                                        last1Budget = CDbl(CType(.Cells(zeile, colFields(allianzSpalten.Budget)), Excel.Range).Value)
                                    Catch ex As Exception
                                        last1Budget = 0.0
                                    End Try

                                    'With current1program
                                    '    .constellationName = itemType.ToString & " - " & pName
                                    'End With

                                    ' wenn jetzt als nächstes gleich wieder eine Programm-Linie kommt, dann muss dem Program als sein erstes und einziges Projekt 
                                    ' die Programm-Linie sein
                                    Dim programItemfound As Boolean = False
                                    Dim nextItemType As Integer
                                    Dim kidItemFound As Boolean = False
                                    Dim tmpZ As Integer = zeile + 1


                                    Do While tmpZ <= lastRow And Not kidItemFound And Not programItemfound
                                        Try
                                            nextItemType = CInt(CType(.Cells(tmpZ, colFields(allianzSpalten.itemType)), Excel.Range).Value)
                                        Catch ex As Exception
                                            nextItemType = 0
                                        End Try

                                        kidItemFound = projektvorhaben.Contains(nextItemType)
                                        programItemfound = pgmlinie.Contains(nextItemType)
                                        tmpZ = tmpZ + 1

                                    Loop

                                    If Not kidItemFound And programItemfound Then
                                        ' jetzt wird sichergestellt, dass diese Programm-Linie jetzt als Projekt angelegt wird ..
                                        ok = True
                                        projVorhabensBudget = last1Budget
                                    ElseIf kidItemFound And tmpZ <= lastRow Then
                                        ok = True
                                        projVorhabensBudget = last1Budget
                                    End If

                                End If
                            End If


                            If ok Then
                                Try

                                    If isOldAllianzImport Then
                                        custFields = readCustomFieldsFromExcel(colCustomFields, 2, zeile, currentWS)
                                    Else
                                        custFields = readCustomFieldsFromExcel(colCustomFields, 2, zeile, currentWS)
                                    End If


                                    ' lese , wieviel Prozent der Gesamtsumme jeweils auf die Release verteilt werden soll 
                                    For i As Integer = 0 To anzReleases - 1
                                        Try
                                            If IsNothing(CType(.Cells(zeile, colRelPrzStart + i), Excel.Range).Value) Then
                                                relPrz(i) = 0.0
                                            Else
                                                relPrz(i) = CDbl(CType(.Cells(zeile, colRelPrzStart + i), Excel.Range).Value)
                                            End If
                                        Catch ex As Exception
                                            relPrz(i) = 0.0
                                        End Try
                                    Next

                                    ' Plausibilitäts-Check - wenn es sich nicht auf 100% summiert, dann lieber alles auf die rootPhase verteilen und nichts auf die Release Phasen
                                    Dim a As Double = relPrz.Sum
                                    If relPrz.Sum > 0 Then
                                        If relPrz.Sum < 0.99 Or relPrz.Sum > 1.01 Then
                                            CType(.Cells(zeile, lastColumn), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color = awinSettings.AmpelGelb
                                            CType(.Cells(zeile, lastColumn), Global.Microsoft.Office.Interop.Excel.Range).AddComment(Text:="Prz-Sätze addieren nicht auf 100% ... alles in Projektphase ")
                                            If relPrz.Sum < 0.99 Then
                                                outPutLine = pName & " Prozent-Sätze > 0 , aber < 1; Gesamt-Summe auf Gesamt-Projekt verteilt  .."
                                            Else
                                                outPutLine = pName & " Prozent-Sätze > 1.0 , Gesamt-Summe auf Gesamt-Projekt verteilt  .."
                                            End If

                                            outputCollection.Add(outPutLine)

                                            Dim logtxt(2) As String
                                            logtxt(0) = "Prozent-Sätze > 0 , aber < 1; Gesamt-Summe auf Gesamt-Projekt verteilt  .. "
                                            logtxt(1) = ""
                                            logtxt(2) = pName

                                            Call logger(ptErrLevel.logInfo, "importAllianzBOBS", logtxt)

                                            ReDim relPrz(anzReleases - 1)
                                        End If
                                    End If


                                    ' was ist der Gesamtbedarf dieser Rolle in dem besagten Vorhaben ? 
                                    For i As Integer = 0 To colRoleNamesToConsider.Length - 1
                                        Try
                                            If IsNothing(CType(.Cells(zeile, colRoleNamesToConsider(i)), Excel.Range).Value) Then
                                                roleNeeds(i) = 0.0
                                            Else
                                                Dim tmpValue As Double = CDbl(CType(.Cells(zeile, colRoleNamesToConsider(i)), Excel.Range).Value) * nrOfDaysMonth
                                                If tmpValue >= 0 Then
                                                    roleNeeds(i) = tmpValue
                                                Else
                                                    roleNeeds(i) = 0.0
                                                End If
                                            End If
                                        Catch ex As Exception
                                            roleNeeds(i) = 0.0
                                        End Try

                                    Next

                                    If Not isOldAllianzImport Then
                                        Dim i_offset As Integer = colRoleNamesToConsider.Length

                                        For i As Integer = 0 To colTEroleNamesToConsider.Length - 1
                                            Try
                                                If IsNothing(CType(.Cells(zeile, colTEroleNamesToConsider(i)), Excel.Range).Value) Then
                                                    roleNeeds(i + i_offset) = 0.0
                                                Else
                                                    roleNeeds(i + i_offset) = 0.0
                                                    Dim cellValue As Double = CDbl(CType(.Cells(zeile, colTEroleNamesToConsider(i)), Excel.Range).Value)
                                                    If cellValue > 0 Then

                                                        Dim tmpRoleDef As clsRollenDefinition = RoleDefinitions.getRoledef(TEroleNamesToConsider(i))

                                                        If Not IsNothing(tmpRoleDef) Then
                                                            ' jetzt handelt es sich um T€ - Werte , das heisst die anzahl Manntage erreichnet sich aus value*1000/tagessatz
                                                            Dim tagessatz As Double
                                                            Dim tmpValue As Double
                                                            Try
                                                                tagessatz = RoleDefinitions.getRoledef(TEroleNamesToConsider(i)).tagessatzIntern
                                                                If tagessatz = 0 Then
                                                                    tagessatz = 800
                                                                    Call MsgBox("tagessatz = 0 ! Rolle " & TEroleNamesToConsider(i))
                                                                End If

                                                                tmpValue = CDbl(CType(.Cells(zeile, colTEroleNamesToConsider(i)), Excel.Range).Value) * 1000 / tagessatz

                                                                If tmpValue >= 0 Then
                                                                    roleNeeds(i + i_offset) = tmpValue
                                                                Else
                                                                    roleNeeds(i + i_offset) = 0.0
                                                                End If
                                                            Catch ex As Exception

                                                            End Try
                                                        End If
                                                    End If


                                                End If
                                            Catch ex As Exception
                                                roleNeeds(i + i_offset) = 0.0
                                            End Try

                                        Next
                                    End If


                                Catch ex As Exception
                                    ok = False
                                End Try

                                ' jetzt werden noch weitere Infos eingelesen ..
                                Try ' Ampelbeschreibung

                                    ampelText = CStr(CType(.Cells(zeile, colFields(allianzSpalten.AmpelText)), Excel.Range).Value)
                                Catch ex As Exception
                                    ampelText = ""
                                End Try

                                Try ' Business Unit

                                    businessUnit = CStr(CType(.Cells(zeile, colFields(allianzSpalten.BusinessUnit)), Excel.Range).Value)
                                Catch ex As Exception
                                    businessUnit = ""
                                End Try

                                Try ' Projektleiter
                                    responsiblePerson = CStr(CType(.Cells(zeile, colFields(allianzSpalten.Responsible)), Excel.Range).Value)
                                Catch ex As Exception
                                    responsiblePerson = ""
                                End Try

                                Try ' Budget
                                    budget = 0.0
                                    If projektvorhaben.Contains(itemType) Then
                                        budget = CStr(CType(.Cells(zeile, colFields(allianzSpalten.pvBudget)), Excel.Range).Value)

                                    ElseIf pgmlinie.Contains(itemType) Then
                                        budget = CStr(CType(.Cells(zeile, colFields(allianzSpalten.Budget)), Excel.Range).Value)
                                        ' wenn dieses Null ist, so soll die andere Spalte genommen werden 
                                        Try
                                            If budget = 0 Then
                                                budget = CStr(CType(.Cells(zeile, colFields(allianzSpalten.pvBudget)), Excel.Range).Value)
                                            End If
                                        Catch ex As Exception

                                        End Try

                                    End If

                                Catch ex As Exception
                                    budget = 0.0
                                End Try


                                Try ' Projekt-Nummer

                                    allianzProjektNummer = CStr(CType(.Cells(zeile, colFields(allianzSpalten.Projektnummer)), Excel.Range).Value)
                                Catch ex As Exception
                                    allianzProjektNummer = ""
                                End Try

                                Try ' Status
                                    If itemType = 6 Then
                                        'ur: 211202: 
                                        'allianzStatus = ProjektStatus(PTProjektStati.geplant)
                                        allianzStatus = VProjectStatus(PTVPStati.initialized)
                                    Else
                                        'ur: 211202: 
                                        'allianzStatus = ProjektStatus(PTProjektStati.beauftragt)
                                        allianzStatus = VProjectStatus(PTVPStati.ordered)
                                    End If

                                Catch ex As Exception
                                    'ur: 211202: 
                                    'allianzStatus = ProjektStatus(PTProjektStati.geplant)
                                    allianzStatus = VProjectStatus(PTVPStati.initialized)
                                End Try

                            End If



                        Catch ex As Exception
                            Call MsgBox("Fehler bei Informationen auslesen: Projekt " & pName)
                            ok = False
                        End Try



                        If ok Then


                            'Projekt anlegen ,Verschiebung um 
                            Dim hproj As clsProjekt = Nothing

                            ' #####################################################################
                            ' Erstellen des Projekts nach den Angaben aus der Batch-Datei 
                            '
                            pName = itemType.ToString & " - " & pName
                            ' lege ein Allianz IT - Projekt an
                            hproj = erstelleProjektausParametern(pName, variantName, vorlageName, startdate, endDate, budget, sFit, risk, allianzProjektNummer,
                                                                 description, custFields, businessUnit, responsiblePerson, allianzStatus,
                                                                 zeile, realRoleNamesToConsider, roleNeeds, Nothing, Nothing, phNames, relPrz, False)

                            ' tk 21.7.19 es wird ein Summary Projekt für die Programm-Linie angelegt  
                            If pgmlinie.Contains(itemType) Then
                                'pName = itemType.ToString & " - " & pName
                                'hproj.name = pName
                                hproj.projectType = ptPRPFType.portfolio
                            End If

                            Try
                                ' Test, ob das Budget auch ausreicht
                                ' wenn nein, einfach Warning ausgeben 
                                Dim tmpGesamtCost As Double = hproj.getGesamtKostenBedarf.Sum
                                If hproj.Erloes - tmpGesamtCost < 0 Then

                                    Dim goOn As Boolean = True
                                    If hproj.Erloes > 0 Then
                                        goOn = (tmpGesamtCost - hproj.Erloes) / hproj.Erloes > 0.05
                                    End If

                                    If goOn Then
                                        Dim logtxt(2) As String
                                        logtxt(0) = "Budget-Überschreitung"
                                        If pgmlinie.Contains(itemType) Then
                                            outPutLine = "Warnung: Budget-Überschreitung bei " & pName & " (Budget=" & hproj.Erloes.ToString("#0.##") & ", Gesamtkosten=" & tmpGesamtCost.ToString("#0.##")
                                            logtxt(1) = "Programm-Linie"
                                        Else
                                            outPutLine = "Warnung: Budget-Überschreitung bei " & pName & " (Budget=" & hproj.Erloes.ToString("#0.##") & ", Gesamtkosten=" & tmpGesamtCost.ToString("#0.##")
                                            logtxt(1) = "Projekt"
                                        End If

                                        outputCollection.Add(outPutLine)
                                        logtxt(2) = pName

                                        Dim values(2) As Double
                                        values(0) = hproj.Erloes
                                        values(1) = tmpGesamtCost
                                        If values(0) > 0 Then
                                            values(2) = tmpGesamtCost / hproj.Erloes
                                        Else
                                            values(2) = 9999999999
                                        End If
                                        Call logger(ptErrLevel.logWarning, "importAllianzBOBS", logtxt, values)
                                    End If

                                End If

                            Catch ex As Exception

                            End Try

                            ' Test tk 
                            Try
                                ReDim logmsg(2)
                                logmsg(0) = "Importiert: "
                                logmsg(1) = ""
                                logmsg(2) = pName

                                For ix As Integer = 1 To realRoleNamesToConsider.Count

                                    Dim tmpRollenName As String = realRoleNamesToConsider(ix - 1)
                                    Dim sollBedarf As Double = roleNeeds(ix - 1)


                                    Dim tmpCollection As New Collection
                                    tmpCollection.Add(tmpRollenName)
                                    Dim istBedarf As Double = hproj.getRessourcenBedarf(tmpRollenName,
                                                                                        inclSubRoles:=True).Sum

                                    If Math.Abs(sollBedarf - istBedarf) > 0.001 Then
                                        outPutLine = "Differenz bei " & pName & ", " & tmpRollenName & ": " & Math.Abs(sollBedarf - istBedarf).ToString("#0.##")
                                        outputCollection.Add(outPutLine)
                                    End If

                                Next

                                Dim sollBedarfGesamt As Double = roleNeeds.Sum
                                Dim istBedarfGesamt As Double = hproj.getAlleRessourcen.Sum

                                If Math.Abs(sollBedarfGesamt - istBedarfGesamt) > 0.001 Then
                                    outPutLine = "Gesamt Differenz bei " & pName & ": " & Math.Abs(sollBedarfGesamt - istBedarfGesamt).ToString("#0.##")
                                    outputCollection.Add(outPutLine)
                                End If

                                If awinSettings.visboDebug Then
                                    Call logger(ptErrLevel.logDebug, "importAllianzBOBS", logmsg, roleNeeds)
                                End If


                            Catch ex As Exception

                            End Try



                            ' Ende Test tk 

                            If Not IsNothing(hproj) Then


                                ' jetzt ist alles so weit ok 
                                Dim pkey As String = ""
                                If Not IsNothing(hproj) Then
                                    Try
                                        pkey = calcProjektKey(hproj)

                                        If ImportProjekte.Containskey(pkey) Then
                                            outPutLine = "Name existiert mehrfach: " & pName
                                            outputCollection.Add(outPutLine)
                                        Else
                                            ImportProjekte.Add(hproj, False)
                                            If projektvorhaben.Contains(itemType) Then
                                                createdProjects = createdProjects + 1


                                                ' jetzt soll das in die Constellation 
                                                Dim cItem As New clsConstellationItem
                                                With cItem
                                                    .projectName = hproj.name
                                                    .variantName = hproj.variantName
                                                    .show = True
                                                    .projectTyp = CType(hproj.projectType, ptPRPFType).ToString
                                                    .zeile = lfdNr1program
                                                End With

                                                current1program.add(cItem)
                                                lfdNr1program = lfdNr1program + 1
                                            End If

                                        End If

                                    Catch ex As Exception
                                        outPutLine = "Fehler bei " & pName & vbLf & "Error: " & ex.Message
                                        outputCollection.Add(outPutLine)
                                    End Try

                                End If


                            Else
                                ok = False
                                If pgmlinie.Contains(itemType) Then
                                    outPutLine = "Fehler beim Erzeugen der Programm-Linie " & pName
                                Else
                                    outPutLine = "Fehler beim Erzeugen des Projektes " & pName
                                End If

                                outputCollection.Add(outPutLine)
                            End If

                        End If

                    End If


                    geleseneProjekte = geleseneProjekte + 1
                    zeile = zeile + 1

                End While

                ' jetzt die letzte ggf vorkommende Constellation aufnehmen 
                If Not IsNothing(current1program) Then

                    If current1program.count > 0 Then
                        ' ggf aus der Liste aller Constellations wieder rausnehmen 

                        If projectConstellations.Contains(current1program.constellationName) Then
                            projectConstellations.Remove(current1program.constellationName)
                        End If

                        createdPrograms = createdPrograms + 1
                        projectConstellations.Add(current1program)

                        ' tk 10.8.19 das wird jetzt wieder gemacht , aber nur um zu überprüfen ob Summe(POBs) <= lastProgramProj
                        ' jetzt das union-Projekt erstellen 


                        '        '  ur: 20211108: accelerate the load of an portfolio without summaryProject - calcUnionProject needs to m

                        ' Dim unionProj As clsProjekt = calcUnionProject(current1program, True, Date.Now.Date.AddHours(23).AddMinutes(59), budget:=last1Budget)

                        'Try
                        '    ' Test, ob das Budget auch ausreicht
                        '    ' wenn nein, einfach Warning ausgeben 
                        '    Dim tmpGesamtCost As Double = unionProj.getGesamtKostenBedarf.Sum
                        '    If unionProj.Erloes - tmpGesamtCost < 0 Then
                        '        Dim goOn As Boolean = True
                        '        If unionProj.Erloes > 0 Then
                        '            goOn = (tmpGesamtCost - unionProj.Erloes) / unionProj.Erloes > 0.05
                        '        End If

                        '        If goOn Then
                        '            outPutLine = "Warnung: Budget-Überschreitung bei Programmlinie " & unionProj.name & " (Budget=" & unionProj.Erloes.ToString("#0.##") & ", Gesamtkosten=" & tmpGesamtCost.ToString("#0.##")
                        '            outputCollection.Add(outPutLine)

                        '            Dim logtxt(2) As String
                        '            logtxt(0) = "Budget-Überschreitung"
                        '            logtxt(1) = "Programmlinie"
                        '            logtxt(2) = unionProj.name
                        '            Dim values(2) As Double
                        '            values(0) = unionProj.Erloes
                        '            values(1) = tmpGesamtCost
                        '            If values(0) > 0 Then
                        '                values(2) = tmpGesamtCost / unionProj.Erloes
                        '            Else
                        '                values(2) = 9999999999
                        '            End If
                        '            Call logfileSchreiben(logtxt, values)
                        '        End If

                        '    End If

                        'Catch ex As Exception

                        'End Try

                        ' Status wird gleich auf 1: beauftragt gesetzt
                        'unionProj.Status = ProjektStatus(PTProjektStati.beauftragt)

                        'If ImportProjekte.Containskey(calcProjektKey(unionProj)) Then
                        '    ImportProjekte.Remove(calcProjektKey(unionProj), updateCurrentConstellation:=False)
                        'End If

                        'ImportProjekte.Add(unionProj, updateCurrentConstellation:=False)

                        '' test
                        Dim everythingOK As Boolean = testUProjandSingleProjs(current1program)
                        If Not everythingOK Then

                            outPutLine = "Summary Projekt nicht identisch mit der Liste der Projekt-Vorhaben: " & current1program.constellationName
                            outputCollection.Add(outPutLine)

                            ReDim logmsg(2)
                            logmsg(0) = "Summary Projekt nicht identisch mit der Liste der Projekt-Vorhaben:"
                            logmsg(1) = ""
                            logmsg(2) = current1program.constellationName
                            Call logger(ptErrLevel.logError, "importAllianzBOBS", logmsg)

                        End If
                        ' ende test
                    Else
                        emptyPrograms = emptyPrograms + 1
                    End If


                End If

            End With
        Catch ex As Exception

            Throw New Exception("Fehler in Import-Datei: " & ex.Message)

        End Try

        If outputCollection.Count > 0 Then
            Call showOutPut(outputCollection, "Import Type 1", "")
        End If

        If emptyPrograms = 0 Then
            Call MsgBox("Zeilen gelesen: " & geleseneProjekte & vbLf &
                    "Projekte erzeugt: " & createdProjects & vbLf &
                    "Programme erzeugt: " & createdPrograms & vbLf &
                    "insgesamt importiert: " & ImportProjekte.Count)
        Else
            Call MsgBox("Zeilen gelesen: " & geleseneProjekte & vbLf &
                    "Projekte erzeugt: " & createdProjects & vbLf &
                    "Programme erzeugt: " & createdPrograms & vbLf &
                    "Programme nicht erzeugt, weil leer: " & emptyPrograms & vbLf &
                    "insgesamt importiert: " & ImportProjekte.Count)
        End If


    End Sub


    ''' <summary>
    ''' aktualisiert Projekte mit den für BOSV-KB angegebenen Werten 
    ''' dabei werden die neuen Daten in das Projekt "gemerged"; d.h alle Werte zu anderen Rollen als BOSV-KB bleiben erhalten 
    ''' Ebenso alle Attribute ; es werden also nur die Rollen-Bedarfe zu BOSV-KB ausgetauscht ...  
    ''' </summary>
    Public Sub importAllianzType2()
        Dim zeile As Integer, spalte As Integer

        Dim tfZeile As Integer = 2

        Dim pName As String = ""
        Dim variantName As String = ""


        Dim upDatedProjects As Integer = 0
        Dim errorProjects As Integer = 0

        ' für den Output 
        Dim outputFenster As New frmOutputWindow
        Dim outputCollection As New Collection
        Dim outPutLine As String = ""


        Dim vorlageName As String = "Rel"
        Dim lastRow As Integer
        Dim lastColumn As Integer
        Dim geleseneProjekte As Integer
        Dim ok As Boolean = False

        ' die Projekte

        Dim hproj As clsProjekt = Nothing
        Dim newProj As clsProjekt = Nothing
        Dim projektKundenNummer As String = ""

        ' welche Rollen sollen gelöscht werden
        Dim deleteRoles As New Collection

        ' jetzt werden die aufgebaut ...
        If awinSettings.ActualdataOrgaUnits = "" Then

            deleteRoles.Add("D-BOSV-KB0")
            deleteRoles.Add("D-BOSV-KB1")
            deleteRoles.Add("D-BOSV-KB2")
            deleteRoles.Add("D-BOSV-KB3")
            deleteRoles.Add("Grp-BOSV-KB")

        Else
            Dim tmpStr() As String = awinSettings.ActualdataOrgaUnits.Split(New Char() {CChar(";")})
            For Each tmpRCName As String In tmpStr
                If RoleDefinitions.containsName(tmpRCName.Trim) Then
                    deleteRoles.Add(tmpRCName.Trim)
                End If
            Next
        End If

        ' diese Rollen und Subroles sollen alle vorher gelöscht werden und dann mit den neuen Werten ersetzt werden 
        ' Amis soll nicht gelöscht werden, deshalb die explizite Aufführung


        ' Standard-Definition
        Dim anzPhasen As Integer = 5

        Try
            anzPhasen = Projektvorlagen.getProject(vorlageName).CountPhases
        Catch ex As Exception
            Call MsgBox("in ImportAllianzType2: " & vbLf & "es gibt keine Projektvorlage " & vorlageName & ".xlsx!" & vbLf & "-> Abbruch ...")
            Exit Sub
        End Try


        ' enthält die eingeplanten PT für die einzelnen Releases  
        Dim phValues() As Double
        ReDim phValues(anzPhasen - 1)

        ' nimmt die Farbe auf, die steuert, dass diese Zeile nicht eingelesen wird ... 
        Dim projectStartingColor As Integer

        Dim currentColor As Integer


        ' enthält die Phasen Namen
        Dim phNameIDs() As String
        ReDim phNameIDs(anzPhasen - 1)

        ' enthält die Spalten-Nummer, ab der die Release Phasen Mann-Tage stehen 
        Dim colRelValues As Integer

        ' enthät die Saplte, wo der ProjektName steht ...
        Dim colPname As Integer

        ' enthält die Spalten-Nummer, wo die einzelnen Rollen-Namen zu finden sind
        Dim colRoleName As Integer = -1

        ' jetzt werden die ImportProjekte zurückgesetzt ...
        ImportProjekte.Clear()

        Dim firstZeile As Excel.Range


        zeile = 2
        spalte = 1
        geleseneProjekte = 0

        ' jetzt werden die Phase-Names besetzt
        Try
            For i = 1 To anzPhasen
                phNameIDs(i - 1) = Projektvorlagen.getProject(vorlageName).getPhase(i).nameID
            Next
        Catch ex As Exception
            Call MsgBox("Probleme mit Vorlage " & vorlageName)
            Exit Sub
        End Try

        ' enthält, wieviel Manntage von dieser Rolle insgesamt benötigt werden 
        Dim rolePhaseValues As New SortedList(Of String, Double())


        Try

            Dim found As Boolean = False
            Dim wsi As Integer = 1
            Dim wsCount As Integer = appInstance.ActiveWorkbook.Worksheets.Count


            While Not found And wsi <= wsCount
                If CType(appInstance.ActiveWorkbook.Worksheets.Item(wsi),
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet).Name.StartsWith("Projekte") Then
                    found = True
                Else
                    wsi = wsi + 1
                End If
            End While

            If Not found Then
                Call MsgBox("keine Projekte-Tabelle gefunden ...")
                Exit Sub
            End If

            Dim currentWS As Excel.Worksheet = CType(appInstance.ActiveWorkbook.Worksheets.Item(wsi),
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet)
            With currentWS


                firstZeile = CType(.Rows(2), Excel.Range)

                ' jetzt wird festgelegt, ab wo die absoluten PT-Werte für die Releases stehen 
                colRelValues = CType(.Range("M1"), Excel.Range).Column

                colPname = CType(.Range("B1"), Excel.Range).Column

                ' wo stehen die Team-Bezeichner
                colRoleName = .Range("D1").Column

                projectStartingColor = CInt(CType(.Cells(2, 2), Excel.Range).Interior.Color)


                'lastColumn = firstZeile.End(XlDirection.xlToLeft).Column

                lastColumn = CType(.Cells(1, 2000), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlToLeft).Column
                lastRow = CType(.Cells(20000, "B"), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlUp).Row



                While zeile < lastRow

                    Dim oldProj As clsProjekt = Nothing

                    currentColor = CInt(CType(.Cells(zeile, 2), Excel.Range).Interior.Color)
                    If currentColor = projectStartingColor Then
                        ' jetzt kommt die Behandlung ...   

                        Try

                            pName = CStr(CType(.Cells(zeile, colPname), Excel.Range).Value).Trim
                            geleseneProjekte = geleseneProjekte + 1
                            ok = isKnownProject(pName, projektKundenNummer, AlleProjekte)

                        Catch ex As Exception
                            ok = False
                        End Try

                        ' startzeile muss jetzt gemerkt werden ...
                        zeile = zeile + 1
                        currentColor = CInt(CType(.Cells(zeile, 2), Excel.Range).Interior.Color)
                        Dim startzeile As Integer = zeile
                        Dim endeZeile As Integer = startzeile

                        ' jetzt schon zeile auf das nächste Projekt positionieren ...
                        Do Until currentColor = projectStartingColor And Not zeile > lastRow
                            zeile = zeile + 1
                            currentColor = CInt(CType(.Cells(zeile, 2), Excel.Range).Interior.Color)
                        Loop
                        ' in zeile ist jetzt das nächste Projekt 

                        If Not ok Then

                            outPutLine = "Projekt nicht bekannt: " & pName
                            outputCollection.Add(outPutLine)

                        Else

                            Dim pvKey As String = calcProjektKey(pName, "")
                            oldProj = AlleProjekte.getProject(pvKey)

                            ' jetzt werden die Values für ein Projekt ausgelsen 
                            rolePhaseValues.Clear()
                            ' in zeile steht das nächste Projekt, in zeile-1 dann der letzte Eintrag des aktuellen Projekts
                            endeZeile = zeile - 1


                            ' jetzt kann rolePhaseValues dimensioniert werden 
                            For iz As Integer = startzeile To endeZeile
                                Dim phaseValues(anzPhasen - 1) As Double
                                Dim roleName As String = ""

                                If Not IsNothing(CType(.Cells(iz, colRoleName), Excel.Range).Value) Then
                                    roleName = CStr(CType(.Cells(iz, colRoleName), Excel.Range).Value).Trim


                                    If roleName <> "" Then

                                        If RoleDefinitions.containsName(roleName) Then

                                            ' jetzt muss die RCNameID bestimmt werden 
                                            Dim rcNameID As String = RoleDefinitions.getRoledef(roleName).UID.ToString
                                            For ip As Integer = 1 To anzPhasen - 1
                                                phaseValues(ip) = CDbl(CType(.Cells(iz, colRelValues + ip - 1), Excel.Range).Value)
                                            Next

                                            If phaseValues.Sum = 0 Then
                                                ' nichts tun
                                            Else
                                                If rolePhaseValues.ContainsKey(rcNameID) Then
                                                    ' addieren ...
                                                    For px As Integer = 1 To anzPhasen - 1
                                                        rolePhaseValues.Item(rcNameID)(px) = rolePhaseValues.Item(rcNameID)(px) + phaseValues(px)
                                                    Next
                                                Else
                                                    ' neu aufnehmen 
                                                    rolePhaseValues.Add(rcNameID, phaseValues)
                                                End If
                                            End If
                                        Else
                                            outPutLine = "Team / Rolle nicht bekannt: " & roleName
                                            outputCollection.Add(outPutLine)
                                        End If

                                    End If

                                End If

                            Next

                            ' jetzt wird der Merge auf das Projekt gemacht 
                            ' dabei wird die updateSummaryRole und alle dazu gehörenden SubRoles gelöscht 
                            ' es müssen aber auch die Gruppe gelöscht werden ... 

                            ' test tk 
                            Dim formerLeft As Integer = showRangeLeft
                            Dim formerRight As Integer = showRangeRight
                            showRangeLeft = getColumnOfDate(CDate("1.1.2018"))
                            showRangeRight = getColumnOfDate(CDate("31.12.2018"))

                            Dim testprojekte As New clsProjekte
                            testprojekte.Add(oldProj)

                            Dim gesamtVorher As Double = oldProj.getAlleRessourcen().Sum
                            Dim gesamtVorher2 As Double = testprojekte.getRoleValuesInMonth("Orga", considerAllSubRoles:=True).Sum
                            Dim bosvVorher As Double = oldProj.getRessourcenBedarf("D-BOSV-KB", inclSubRoles:=True).Sum

                            ' tk test ...
                            If Math.Abs(gesamtVorher - gesamtVorher2) >= 0.001 Then
                                Call MsgBox(oldProj.name & " Einzelproj <> Portfolio" & gesamtVorher.ToString & " <> " & gesamtVorher2.ToString)
                            End If
                            ' tk test ...

                            ' jetzt alle Rollen und SubRoles von updateSummaryRole löschen 
                            newProj = oldProj.deleteRolesAndCosts(deleteRoles, Nothing, True)
                            Dim gesamtNachher As Double = newProj.getAlleRessourcen().Sum

                            ' tk test ...
                            For Each tmpRoleName As String In deleteRoles
                                Dim bosvNachher As Double = newProj.getRessourcenBedarf(tmpRoleName, inclSubRoles:=True).Sum

                                If Not bosvNachher = 0 Then
                                    Call MsgBox(tmpRoleName & " wurde nicht gelöscht ... Fehler bei" & newProj.name)
                                End If
                            Next
                            ' tk test ...


                            ' jetzt alle Rollen / Phasen Werte hinzufügen 
                            Dim addValues As Double = 0.0
                            For Each kvp As KeyValuePair(Of String, Double()) In rolePhaseValues
                                addValues = addValues + kvp.Value.Sum
                            Next
                            newProj = newProj.merge(rolePhaseValues, Nothing, phNameIDs, True)

                            Dim bosvErgebnis As Double = newProj.getRessourcenBedarf("Grp-BOSV-KB", inclSubRoles:=True).Sum

                            If Math.Abs(bosvErgebnis - addValues) >= 0.001 Then
                                outPutLine = "addValues ungleich ergebnis: " & addValues.ToString("#0.##") & " <> " & bosvErgebnis.ToString("#0.##")
                                outputCollection.Add(outPutLine)
                            End If

                            ' jetzt in die Import-Projekte eintragen 
                            upDatedProjects = upDatedProjects + 1
                            ImportProjekte.Add(newProj, updateCurrentConstellation:=False)

                            ' wegen test 
                            showRangeLeft = formerLeft
                            showRangeRight = formerRight
                        End If

                    End If

                End While


            End With
        Catch ex As Exception

            Throw New Exception("Fehler in Import-Datei" & ex.Message)

        End Try

        If outputCollection.Count > 0 Then
            Call showOutPut(outputCollection, "Import Detail-Planungs Typ 2", "")
        End If

        Call MsgBox("Zeilen gelesen: " & geleseneProjekte & vbLf &
                    "Projekte aktualisiert: " & upDatedProjects)


    End Sub


    ''' <summary>
    ''' übernimmt für Projekte, die bislang noch keine Projekt-Nummern hatten, die Projekt-Nummer  
    ''' </summary>
    Sub importAllianzType4()
        Dim zeile As Integer, spalte As Integer

        Dim tfZeile As Integer = 2

        Dim pName As String = ""
        Dim variantName As String = ""


        Dim upDatedProjects As Integer = 0
        Dim errorProjects As Integer = 0

        ' für den Output 
        Dim outputFenster As New frmOutputWindow
        Dim outputCollection As New Collection
        Dim outPutLine As String = ""


        Dim lastRow As Integer
        Dim geleseneProjekte As Integer
        Dim ok As Boolean = False

        ' die Projekte

        Dim hproj As clsProjekt = Nothing
        Dim projektKundenNummer As String = ""



        ' enthät die Saplte, wo der "alte" ProjektName steht ...
        Dim colPname As Integer = 4
        ' enthält die Spalte, wo die PRojekt-Nummer drin steht 

        ' enthält die Spalten-Nummer, wo die einzelnen Rollen-Namen zu finden sind
        Dim colPNr As Integer = 3


        ' jetzt werden die ImportProjekte zurückgesetzt ...
        ImportProjekte.Clear()

        Dim firstZeile As Excel.Range


        zeile = 2
        spalte = 1
        geleseneProjekte = 0


        Try

            Dim found As Boolean = False
            Dim wsi As Integer = 1
            Dim wsCount As Integer = appInstance.ActiveWorkbook.Worksheets.Count


            While Not found And wsi <= wsCount
                If CType(appInstance.ActiveWorkbook.Worksheets.Item(wsi),
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet).Name.StartsWith("logBuch") Then
                    found = True
                Else
                    wsi = wsi + 1
                End If
            End While

            If Not found Then
                Call MsgBox("keine Projekte-Tabelle gefunden ...")
                Exit Sub
            End If

            Dim currentWS As Excel.Worksheet = CType(appInstance.ActiveWorkbook.Worksheets.Item(wsi),
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet)
            With currentWS


                firstZeile = CType(.Rows(2), Excel.Range)
                lastRow = CType(.Cells(20000, "D"), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlUp).Row



                While zeile < lastRow

                    Dim oldProj As clsProjekt = Nothing

                    Try

                        pName = CStr(CType(.Cells(zeile, colPname), Excel.Range).Value).Trim
                        projektKundenNummer = CStr(CType(.Cells(zeile, colPNr), Excel.Range).Value).Trim

                        If pName <> "" And projektKundenNummer <> "" Then
                            ' jetzt kann ggf der Update erfolgen ... 
                            geleseneProjekte = geleseneProjekte + 1
                            ok = isKnownProject(pName, projektKundenNummer, AlleProjekte)
                        End If


                    Catch ex As Exception
                        ok = False
                    End Try

                    If ok Then
                        hproj = getProjektFromSessionOrDB(pName, "", AlleProjekte, Date.Now)
                        If Not IsNothing(hproj) Then

                            If hproj.kundenNummer = "" Then
                                hproj.kundenNummer = projektKundenNummer
                                ' jetzt in die Import-Projekte eintragen 
                                upDatedProjects = upDatedProjects + 1
                                ImportProjekte.Add(hproj, updateCurrentConstellation:=False)
                            Else
                                outPutLine = "Projekt hat bereits eine Kunden-Nummer: " & pName & " old-Nr: " & hproj.kundenNummer & "; new-Nr: " & projektKundenNummer
                                outputCollection.Add(outPutLine)
                            End If



                        Else
                            outPutLine = "Projekt nicht in Datenbank gefunden: " & pName
                            outputCollection.Add(outPutLine)
                        End If
                    Else
                        outPutLine = "Projekt nicht bekannt: " & pName
                        outputCollection.Add(outPutLine)
                    End If

                    zeile = zeile + 1

                End While


            End With
        Catch ex As Exception

            Throw New Exception("Fehler in Import-Datei" & ex.Message)

        End Try

        If outputCollection.Count > 0 Then
            Call showOutPut(outputCollection, "Import Detail-Planungs Typ 4", "")
        End If

        Call MsgBox("Zeilen gelesen: " & geleseneProjekte & vbLf &
                    "Projekte aktualisiert: " & upDatedProjects)


    End Sub


    ''' <summary>
    ''' bestimmt den Import-Typ und das Worksheet, das eingelesen werden soll ..
    ''' </summary>
    ''' <param name="importType"></param>
    ''' <returns></returns>
    Public Function bestimmeWsAndImporttype(ByRef importType As ptImportTypen) As Excel.Worksheet
        Dim resultWS As Excel.Worksheet = Nothing
        Dim wb As Excel.Worksheet = CType(appInstance.ActiveSheet, Excel.Worksheet)
        Dim tmpImportType As Integer = -1

        Try

            If wb.Name = "BOBundScope" Then
                resultWS = wb
                importType = ptImportTypen.allianzBOBImport
            Else
                tmpImportType = CInt(wb.Range(visboImportKennung).Value)

                If [Enum].IsDefined(GetType(ptImportTypen), tmpImportType) Then
                    resultWS = CType(CType(wb.Range(visboImportKennung), Excel.Range).Parent, Excel.Worksheet)
                    importType = tmpImportType
                End If
            End If

        Catch ex As Exception
            resultWS = Nothing
        End Try

        bestimmeWsAndImporttype = resultWS

    End Function


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





    ''' <summary>
    ''' hiermit werden die missingMilestoneDefinitions zu den MilestoneDefinitions und 
    ''' die missingPhaseDefinitions zu den PhaseDefinitions hinzugefügt
    ''' </summary>
    Public Sub addMissingDefs2Defs()

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
        ElseIf menuOption = PTmenue.reportBHTC Or
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
    ''' Umwandlung eines Datum des Typs ISO-Datums-String in ein Date
    ''' </summary>
    ''' <param name="ISODate"></param>
    ''' <returns></returns>
    Public Function ISODateToDateTime(ByVal ISODate As String) As DateTime

        Dim newDate As DateTime = Nothing
        Try
            newDate = Date.Parse(ISODate, Nothing, DateTimeStyles.RoundtripKind)
        Catch ex As Exception
            newDate = Nothing
            Throw New ArgumentException("´Fehler bei der Datumsumwandlung von ISO-Datum-String in Date:  " & ISODate)
        End Try

        ISODateToDateTime = newDate
    End Function

End Module
