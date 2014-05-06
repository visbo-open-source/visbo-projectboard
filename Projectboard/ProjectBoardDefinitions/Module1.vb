Imports System.Collections.Generic
Imports System.Math
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Core



Public Module Module1


    ' in Modul 1 sollten jetzt alle Konstanten und Einstellungen in einer Klasse zusammengefasst werden
    ' awinSettings: für StartOfCalendar, linker Rand, rechter Rand, ...
    ' Laufzeit Parameter;

    Public awinSettings As New clsawinSettings
    Public magicBoardCmdBar As New clsCommandBarEvents
    Public anzahlCalls As Integer = 0
    Public iProjektFarbe As Object
    Public iWertFarbe As Object
    'Public HoehePrcChart As Double

    Public vergleichsfarbe0 As Object
    Public vergleichsfarbe1 As Object
    Public vergleichsfarbe2 As Object
    Public ergebnisfarbe1 As Object
    Public ergebnisfarbe2 As Object

    ' diese Variable steuert, ob die Ereignis-Routine Cmdbar.onupdate durchlaufen wird oder gleich zu beginn wieder verlassen wird
    ' wird immer dann auf false gesetzt , wenn in eigenen Routinen Projekte gesetzt, gelöscht oder ins Show/Noshow gestellt werden 
    Public enableOnUpdate As Boolean = True


    Public Projektvorlagen As New clsProjektvorlagen
    Public ShowProjekte As New clsProjekte
    Public selectedProjekte As New clsProjekte
    Public AlleProjekte As New SortedList(Of String, clsProjekt)
    Public ImportProjekte As New clsProjekte
    Public DeletedProjekte As New clsProjekte
    Public projectConstellations As New clsConstellations
    Public currentConstellation As String = "" ' hier wird mitgeführt, was die aktuelle Projekt-Konstellation ist 
    Public allDependencies As New clsDependencies

    ' hier wird die Projekt Historie eines Projektes aufgenommen 
    Public projekthistorie As New clsProjektHistorie

    Public specialListofPhases As New Collection

    Public feierTage As New SortedSet(Of Date)

    Public timeMachineIsOn As Boolean = False


    Public PfChartBubbleNames() As String

    Public RoleDefinitions As New clsRollen
    Public PhaseDefinitions As New clsPhasen
    Public CostDefinitions As New clsKostenarten
    Public DiagramList As New clsDiagramme
    Public awinButtonEvents As New clsAwinEvents

    ' Variable gibt ab, ob die Time Zone, die auf Diagramme wirkt, gezeigt werden soll oder nicht
    Public showtimezone As Boolean


    ' damit ist das Formular Milestone / Status / Phase überall verfügbar
    Public formMilestone As New frmMilestoneInformation
    Public formStatus As New frmStatusInformation
    Public formPhase As New frmPhaseInformation



    ' variable gibt an, zu welchem Objekt-Rolle (Rolle, Kostenart, Ergebnis, ..)  der Röntgen Blick gezeigt wird 
    Public roentgenBlick As New clsBestFitObject

    ' diese beiden folgenden Variablen steuern im Sheet "Ressourcen", welcher Bereich in den Diagrammen angezeigt werden soll
    Public showRangeLeft As Integer
    Public showRangeRight As Integer

    ' diese beiden Variablen nehmen die Farben auf für Showtimezone bzw. Noshowtimezone
    Public showtimezone_color As Object, noshowtimezone_color As Object


    ' maxScreenHeight, maxScreenWidth gibt die maximale Höhe/Breite des Bildschirms in Punkten an 
    Public maxScreenHeight As Double, maxScreenWidth As Double
    Public boxWidth As Double = 19.3, boxHeight As Double, topOfMagicBoard As Double
    Public screen_correct As Double = 0.26
    Public miniWidth As Double = 126 ' wird aber noch in Abhängigkeit von maxscreenwidth gesetzt 
    Public miniHeight As Double = 70 ' wird aber noch in abhängigkeit von maxscreenheight gesetzt

    ' diese Konstanten werden benötigt, um die Diagramme gemäß des gewählten Zeitraums richtig zu positionieren
    Public Const summentitel1 As String = "Prognose Ergebniskennzahl"
    Public Const summentitel2 As String = "strategischer Fit, Risiko & Marge"
    Public Const summentitel3 As String = "Personal-Kosten intern/extern"
    Public Const summentitel4 As String = "Personal Kosten Struktur"
    Public Const summentitel5 As String = "Ergebnis Verbesserungs-Potentiale"
    Public Const summentitel6 As String = "Bisherige Ziel-Erreichung"
    Public Const summentitel7 As String = "Prognose zukünftige Ziel-Erreichung"
    Public Const summentitel8 As String = "Bisherige & zukünftige Ziel-Erreichung"
    Public Const summentitel9 As String = "Auslastungs-Übersicht"
    Public Const summentitel10 As String = "Details zur Über-Auslastung"
    Public Const summentitel11 As String = "Details zur Unter-Auslastung"
    Public Const maxProjektdauer As Integer = 60

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
    End Enum

    ' wird in awinSetTypen dimensioniert und gesetzt 
    Public portfolioDiagrammtitel() As String



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

    ' dieser array nimmt die Koordinaten der Formulare auf 
    ' die Koordinaten werden in der Reihenfolge gespeichert: top, left, width, height 
    Public frmCoord(19, 3) As Double

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
    End Enum

    Public Enum PTpinfo
        top = 0
        left = 1
        width = 2
        height = 3
    End Enum

   
    Public StartofCalendar As Date = #1/1/2012# ' wird in Customization File gesetzt - dies hier ist nur die Default Einstellung 

    Public weightStrategicFit As Double

    '
    '
    ' Projektstatus kann sein:
    ' beendet
    ' geplant
    ' beauftragt
    ' abgeschlossen
    Public ProjektStatus(4) As String


    '
    '
    ' Diagramm-Typ kann sein:
    ' Phase
    ' Rolle
    ' Kostenart
    ' Summe
    ' portfolio

    ' Variable nimmt die Namen der Diagramm-Typen auf 
    Public DiagrammTypen(6) As String

    ' Variable nimmt die Namen der Windows auf  
    Public windowNames(5) As String

    ' Variable nimmt die Namen der Ergebnis Charts auf  
    Public ergebnisChartName(3) As String

    ' diese Variabe nimmt die Farbe der Kapa-Linie an
    Public rollenKapaFarbe As Object

    ' diese Variable nimmt die Farbe der internen Ressourcen, ohne Projekte an auf
    Public farbeInternOP As Object

    ' diese Variable nimmt die Farbe der externen Ressourcen auf
    Public farbeExterne As Object

    ' Variable nimmt die Namen der Worksheets für Portfolio und Ressourcen auf
    Public arrWsNames(0 To 20) As String

    ' variable nimmt auf, wieviel Tage ein Monat hat
    Public nrOfDaysMonth As Double

    ' so werden in Visual Basic die Worksheets der aktuell geladenen Excel Applikation zugänglich gemacht   
    Public appInstance As _Application

    ' nimmt den Pfad Namen auf - also wo liegen Customization File und Projekt-Details
    Public awinPath As String
    Public requirementsOrdner As String = "requirements\"
    Public customizationFile As String = requirementsOrdner & "Project Board Customization.xlsx" ' Projekt Tafel Customization.xlsx
    Public projektFilesOrdner As String = "ProjectFiles"
    Public deletedFilesOrdner As String = "DeletedFiles"
    Public rplanimportFilesOrdner As String = "RPLANImport"
    Public projektVorlagenOrdner As String = requirementsOrdner & "ProjectTemplates"
    ' Public projektDetail As String = "Project Detail.xlsx"
    Public projektAustausch As String = requirementsOrdner & "Projekt-Steckbrief.xlsx"
    Public projektRessOrdner As String = requirementsOrdner & "Ressource Manager"
    Public RepProjectVorOrdner As String = requirementsOrdner & "ReportTemplatesProject"
    Public RepPortfolioVorOrdner As String = requirementsOrdner & "ReportTemplatesPortfolio"
    Public demoModusHistory As Boolean = False
    Public historicDate As Date = #6/6/2012#

    Public FirstX As Double = -1.0
    Public FirstY As Double = -1.0
    Public LastX As Double = -1.0
    Public LastY As Double = -1.0
    Public firstPress As Boolean = True







    '
    ' Funktion prüft, ob der angegebene Name bereits Element der Projektliste ist
    '
    Function inProjektliste(strName As String) As Boolean

        Dim found As Boolean
        Dim foundinDatabase As Boolean = False

        If Len(strName) < 2 Then
            found = False
        ElseIf AlleProjekte.ContainsKey(strName & "#") Then
            found = True
        ElseIf foundinDatabase Then
            ' hier muss noch geprüft werden, ob das File in der Datenbank vorkommt 
            found = True
        Else
            found = False
        End If

        inProjektliste = found

    End Function

    ''' <summary>
    ''' prüft ob in der Plantafel in Zeile "zeile" ab Spalte "spalte" Platz für ein Projekt der Länge "laenge" ist
    ''' </summary>
    ''' <param name="zeile"></param>
    ''' <param name="spalte"></param>
    ''' <param name="laenge"></param>
    ''' <returns>true, wenn Platz ist; false, wenn kein ausreichender Platz ist</returns>
    ''' <remarks></remarks>
    Function istFrei(zeile As Integer, spalte As Integer, laenge As Integer) As Boolean
        Dim i As Integer
        Dim b As Boolean

        b = True

        With appInstance.Worksheets(arrWsNames(3))

            i = 0
            While i < laenge And b
                If .Cells(zeile, spalte).Offset(0, i).Interior.ColorIndex <> -4142 Then
                    b = False
                End If
                i = i + 1
            End While
        End With

        istFrei = b

    End Function

    Function istStartOfProject(ByVal wsname As String, ByVal zeile As Integer, ByVal spalte As Integer, ByRef pname As String) As Boolean
        Dim found As Boolean
        Dim inputstr As String, commentstr As String

        commentstr = ""
        inputstr = ""

        With appInstance.Worksheets(wsname)

            If Not IsNumeric(.Cells(zeile, spalte).Value) And Len(.Cells(zeile, spalte).Value) > 2 Then
                inputstr = .Cells(zeile, spalte).Value
            End If

            If Not .Cells(zeile, spalte).Comment Is Nothing Then
                commentstr = .Cells(zeile, spalte).Comment.text
            End If

            If inProjektliste(inputstr) Then
                found = True
                pname = inputstr
            ElseIf inProjektliste(commentstr) Then
                found = True
                pname = commentstr
            Else
                found = False
                pname = " "
            End If

        End With

        istStartOfProject = found

    End Function

    Sub awinLoescheProjekt(pname As String)
        '
        'Prozedur löscht in Ws Ressourcen alle zeilen, die den Projektnamen enthalten
        '
        '
        'Dim zeile As Integer, endpunkt As Integer
        Dim hproj As clsProjekt

        Dim tfz As Integer, tfs As Integer
        Dim key As String



        ' prüfen, ob es in der ShowProjektListe ist ...
        If ShowProjekte.Liste.ContainsKey(pname) Then

            ' Shape wird gelöscht - ausserdem wird der Verweis in hproj auf das Shape gelöscht 
            Call clearProjektinPlantafel(pname)


            Try
                hproj = ShowProjekte.getProject(pname)
                key = hproj.name & "#" & hproj.variantName
                Try
                    DeletedProjekte.Add(hproj)
                Catch ex As Exception
                    ' nichts tun, dann wurde das eben schon mal gelöscht ..
                End Try

            Catch ex As Exception
                Call MsgBox(" Fehler in Delete " & pname & " , Modul: awinLoescheProjekt")
                Exit Sub
            End Try



            With hproj
                tfz = .tfZeile
                tfs = .tfspalte
            End With


            ShowProjekte.Remove(pname)
            AlleProjekte.Remove(key)


            Dim abstand As Integer ' eigentlich nur Dummy Variable, wird aber in Tabelle2 benötigt ...
            Call awinClkReset(abstand)

            ' ein Projekt wurde gelöscht bzw aus Showprojekte entfernt  - typus = 3
            Call awinNeuZeichnenDiagramme("3")



        Else
            Call MsgBox("Projekt " & pname & " wurde nicht gefunden")
        End If




    End Sub

    '
    ' prüft , ob übergebenes Diagramm ein Ergebnis Diagramm ist - in index steht ggf als Ergebnis die entsprechende Nummer; 0 wenn es kein Ergebnis Diagramm ist
    '
    Function istErgebnisDiagramm(ByRef chtobj As ChartObject, ByRef index As Integer) As Boolean
        Dim e As Integer
        Dim found As Boolean
        Dim anzErgebnisArten As Integer = 2
        'Dim chtTitle As String


        e = 1
        index = 0
        found = False

        'Try
        '    chtTitle = chtobj.Chart.ChartTitle.Text
        'Catch ex As Exception
        '    chtTitle = " "
        'End Try

        'While Not found And e <= anzErgebnisArten
        '    If chtTitle Like ergebnisChartName(e - 1) & "*" Then
        '        found = True
        '    Else
        '        e = e + 1
        '    End If
        'End While

        'If found Then
        '    index = e
        'End If

        istErgebnisDiagramm = found

    End Function

    '
    ' prüft , ob übergebenes Diagramm ein Rollen Diagramm ist - in R steht ggf als Ergebnis die entsprechende Rollen-Nummer; 0 wenn es kein Rollen Diagramm ist
    '
    Function istRollenDiagramm(ByRef chtobj As ChartObject) As Boolean

        Dim found As Boolean
        Dim chtobjName As String
        Dim tmpStr(20) As String




        found = False
        chtobjName = chtobj.Name

        Try

            tmpStr = chtobjName.Split(New Char() {"#"}, 20)
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
        Dim ergebnis As Boolean = False

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
        Dim r As Integer
        Dim found As Boolean
        Dim chtTitle As String


        Try
            chtTitle = chtobj.Chart.ChartTitle.Text
        Catch ex As Exception
            chtTitle = " "
        End Try

        r = 1
        rwert = 0
        found = False

        With chtobj
            If chtTitle Like (summentitel1 & "*") Then
                found = True
                rwert = 1
            ElseIf chtTitle Like (summentitel2 & "*") Then
                found = True
                rwert = 2
                ' summentitel3, summentitel4 nicht mehr relevant
            ElseIf chtTitle Like (summentitel3 & "*") Then
                found = True
                rwert = 3
            ElseIf chtTitle Like (summentitel4 & "*") Then
                found = True
                rwert = 4
            ElseIf chtTitle Like (ergebnisChartName(0) & " / " & ergebnisChartName(1) & "*") Then
                found = True
                rwert = 5
            ElseIf chtTitle Like (summentitel6 & "*") Then
                found = True
                rwert = 6
            ElseIf chtTitle Like (summentitel7 & "*") Then
                found = True
                rwert = 7
            ElseIf chtTitle Like (summentitel8 & "*") Then
                found = True
                rwert = 8
            ElseIf chtTitle Like (summentitel9 & "*") Then
                found = True
                rwert = 9
            ElseIf chtTitle Like (summentitel10 & "*") Then
                found = True
                rwert = 10
            ElseIf chtTitle Like (summentitel11 & "*") Then
                found = True
                rwert = 11
            End If
        End With

        ' das folgende soll das zukünftige Schema werden 
        Dim chtobjName As String
        Dim tmpStr(20) As String
        chtobjName = chtobj.Name

        Try

            tmpStr = chtobjName.Split(New Char() {"#"}, 20)
            If tmpStr(0) = "pf" And tmpStr.Length >= 2 Then

                If CInt(tmpStr(1)) = PTpfdk.Budget Then

                    found = True
                    rwert = PTpfdk.Budget

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

            tmpStr = chtobjName.Split(New Char() {"#"}, 20)
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

            tmpStr = chtobjName.Split(New Char() {"#"}, 20)
            If tmpStr(0) = "pf" And tmpStr.Length >= 2 Then

                If CInt(tmpStr(1)) = PTpfdk.Phasen Then

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

            tmpStr = chtobjName.Split(New Char() {"#"}, 20)
            If tmpStr(0) = "pf" And tmpStr.Length >= 2 Then

                If CInt(tmpStr(1)) = PTpfdk.Meilenstein Then
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

            tmpStr = chtobjName.Split(New Char() {"#"}, 20)
            If tmpStr(0) = "pf" And tmpStr.Length >= 2 Then

                If CInt(tmpStr(1)) = PTpfdk.FitRisiko Or _
                    CInt(tmpStr(1)) = PTpfdk.FitRisikoVol Or _
                    CInt(tmpStr(1)) = PTpfdk.ComplexRisiko Or _
                    CInt(tmpStr(1)) = PTpfdk.Dependencies Or _
                    CInt(tmpStr(1)) = PTpfdk.ZeitRisiko Then

                    found = True

                End If

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

    Sub awinClkReset(abstand As Integer)
        'Dim hproj As clsProjekt
        'Dim tfz As Integer, tfs As Integer
        'Dim plaenge As Integer
        'Dim pcolor As Object = 0
        'Dim failfree As Boolean = True


        'Exit Sub

        'Call DeleteStartMarkers()

        'If selectedProjects(1) <> "" Then
        '    ' jetzt muss das bisher selektierte Projekt zurückgesetzt werden 
        '    Try
        '        hproj = ShowProjekte.getProject(selectedProjects(1))
        '        With hproj
        '            tfz = .tfZeile
        '            tfs = .tfSpalte
        '            plaenge = .Dauer
        '            pcolor = .farbe
        '        End With

        '    Catch ex As Exception
        '        failfree = False
        '    End Try

        'End If


        abstand = 0


    End Sub



    Sub awinRightClickinPortfolioAendern()
        Dim myBar As CommandBar
        Dim myitem As CommandBarButton
        'Dim myitem As CommandBarControl
        Dim i As Integer, endofsearch As Integer
        Dim found As Boolean
        Dim awinevent As clsEventsPfCharts

        found = False
        i = 1

        With appInstance.CommandBars
            endofsearch = .Count

            While i <= endofsearch And Not found
                If .Item(i).Name = "awinRightClickinPortfolio" Then
                    found = True
                Else
                    i = i + 1
                End If
            End While
        End With

        If found Then
            Exit Sub
        End If

        'CommandBars.Item.Name
        myBar = appInstance.CommandBars.Add(Name:="awinRightClickinPortfolio", Position:=MsoBarPosition.msoBarPopup, Temporary:=True)


        ' Add a menu item
        myitem = myBar.Controls.Add(Type:=MsoControlType.msoControlButton)
        With myitem
            .Caption = "Umbenennen"
            .Tag = "Umbenennen"
            '.OnAction = "awinRenameProject"
        End With
        'awinevent = New clsAwinEvent
        'awinevent.Button3Events = myitem
        awinevent = New clsEventsPfCharts
        awinevent.PfChartRightClick = myitem
        awinButtonEvents.Add(awinevent)


        ' Add a menu item
        myitem = myBar.Controls.Add(Type:=MsoControlType.msoControlButton)
        With myitem
            .Caption = "Löschen"
            .Tag = "Loesche aus Portfolio"
            '.OnAction = "awinDeleteChartorProject"
        End With
        'awinevent = New clsAwinEvent
        'awinevent.Button3Events = myitem
        awinevent = New clsEventsPfCharts
        awinevent.PfChartRightClick = myitem
        awinButtonEvents.Add(awinevent)

        ' Add a menu item
        myitem = myBar.Controls.Add(Type:=MsoControlType.msoControlButton)
        With myitem
            .Caption = "Show / Noshow"
            .Tag = "Show / Noshow"
            '.OnAction = "awinShowNoShowProject"
        End With
        'awinevent = New clsAwinEvent
        'awinevent.Button3Events = myitem
        awinevent = New clsEventsPfCharts
        awinevent.PfChartRightClick = myitem
        awinButtonEvents.Add(awinevent)

        ' Add a menu item
        myitem = myBar.Controls.Add(Type:=MsoControlType.msoControlButton)
        With myitem
            .Caption = "Bearbeiten Projekt-Attribute"
            .Tag = "Bearbeiten Projekt-Attribute"
            '.OnAction = "awinEditDataProject"
        End With
        'awinevent = New clsAwinEvent
        'awinevent.Button3Events = myitem
        awinevent = New clsEventsPfCharts
        awinevent.PfChartRightClick = myitem
        awinButtonEvents.Add(awinevent)

        ' Add a menu item
        myitem = myBar.Controls.Add(Type:=MsoControlType.msoControlButton)
        With myitem
            .Caption = "Beauftragen"
            .Tag = "Beauftragen"
            '.OnAction = "awinBeauftrageProject"
        End With
        'awinevent = New clsAwinEvent
        'awinevent.Button3Events = myitem
        awinevent = New clsEventsPfCharts
        awinevent.PfChartRightClick = myitem
        awinButtonEvents.Add(awinevent)

    End Sub

    Sub awinRightClickinPRCCharts()
        Dim myBar As CommandBar
        Dim myitem As CommandBarButton
        Dim i As Integer, endofsearch As Integer
        Dim found As Boolean
        'Dim awinevent As clsAwinEvents
        Dim awinevent As clsEventsPrcCharts

        found = False
        i = 1

        With appInstance.CommandBars
            endofsearch = .Count

            While i <= endofsearch And Not found
                If .Item(i).Name = "awinRightClickinPRCChart" Then
                    found = True
                Else
                    i = i + 1
                End If
            End While
        End With

        If found Then
            Exit Sub
        End If

        'CommandBars.Item.Name
        myBar = appInstance.CommandBars.Add(Name:="awinRightClickinPRCChart", Position:=MsoBarPosition.msoBarPopup, Temporary:=True)


        ' Add a menu item
        myitem = myBar.Controls.Add(Type:=MsoControlType.msoControlButton)
        With myitem
            .Caption = "Löschen"
            .Tag = "Löschen"
            '.OnAction = "awinDeleteChartorProject"
        End With

        'awinevent = New clsAwinEvent
        'awinevent.Button4Events = myitem
        awinevent = New clsEventsPrcCharts
        awinevent.PrcChartRightClick = myitem
        awinButtonEvents.Add(awinevent)

        ' Add a menu item
        myitem = myBar.Controls.Add(Type:=MsoControlType.msoControlButton)
        With myitem
            .Caption = "Röntgenblick ein/aus"
            .Tag = "Bedarf anzeigen"
            '.OnAction = "awinShowNeedsOfProjects"
        End With
        'awinevent = New clsAwinEvent
        'awinevent.Button4Events = myitem
        awinevent = New clsEventsPrcCharts
        awinevent.PrcChartRightClick = myitem
        awinButtonEvents.Add(awinevent)

        ' Add a menu item
        myitem = myBar.Controls.Add(Type:=MsoControlType.msoControlButton)
        With myitem
            .Caption = "Optimieren"
            .Tag = "Optimieren"
            '.OnAction = "awinOptimizeStartOfProjects"
        End With
        'awinevent = New clsAwinEvent
        'awinevent.Button4Events = myitem
        awinevent = New clsEventsPrcCharts
        awinevent.PrcChartRightClick = myitem
        awinButtonEvents.Add(awinevent)


    End Sub

    Sub awinKontextReset()

        Try
            appInstance.CommandBars("awinRightClickinPortfolio").Delete()
        Catch ex As Exception

        End Try

        Try
            appInstance.CommandBars("awinRightClickinPRCChart").Delete()
        Catch ex As Exception

        End Try


        ' die Short Cut Menues aus Excel wieder alle aktivieren ...
        'Dim cbar As CommandBar

        'For Each cbar In appInstance.CommandBars

        '    cbar.Enabled = True
        '    'Try
        '    '    cbar.Reset()
        '    'Catch ex As Exception

        '    'End Try

        'Next


    End Sub

    Function awinRetrieveUniqueId(objekttyp As String) As Long
        Dim nr As Long

        With appInstance.Worksheets(arrWsNames(14))
            nr = .Cells(1, 1).Value
            .Cells(1, 1).Value = nr + 1
            .Cells(2 + nr, 1).Value = nr
            .Cells(2 + nr, 2).Value = objekttyp
        End With

        awinRetrieveUniqueId = nr

    End Function



    '
    ' gibt die Überdeckung zurück zwischen den beiden Zeiträumen definiert durch showRangeLeft /showRangeRight und anfang / ende
    ' anzahl enthält die Breite der Überdeckung
    ' ixzeitraum gibt an , in welchem Monat des Zeitraums die Überdeckung anfängt: 0 = 1. Monat
    ' ix gibt an, in welchem Monat des durch Anfang / ende definierten Zeitraums die Überdeckung anfängt
    '
    Sub awinIntersectZeitraum(anfang As Integer, ende As Integer, _
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
        Dim chtobj As ChartObject

        With appInstance.Worksheets(arrWsNames(3))

            For Each chtobj In .ChartObjects
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
        Dim chtobj As ChartObject
        Dim chtTitle As String

        ' finde alle Charts, die Cockpit Chart sind und vom Typ her diagrammtypen(prctyp)

        With appInstance.Worksheets(arrWsNames(3))
            Dim found As Boolean
            For Each chtobj In .ChartObjects
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
                        If (chtTitle = DiagramList.getDiagramm(i).DiagrammTitel) And _
                                        (DiagramList.getDiagramm(i).isCockpitChart = True) And _
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

        Dim chtobj As ChartObject
        Dim tstLeft As Double
        Dim tmpArray() As String

        ' finde alle Charts, die bei left platziert sind ... 



        With appInstance.Worksheets(arrWsNames(3))

            For Each chtobj In .ChartObjects

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
            If (chtTitle = DiagramList.getDiagramm(i).DiagrammTitel) And _
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
    ''' De-Selektion aller Objekte durch Selektion einer Zelle in Zeile 2 in der Mitte des aktuell gezeigten Fensters 
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Sub awinDeSelect()
        Dim srow As Integer = 1
        Dim ziel As Integer


        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False

        ' Selektierte Projekte auf Null setzen 

        If selectedProjekte.Count > 0 Then
            selectedProjekte.Clear()
            Call awinNeuZeichnenDiagramme(8)
        End If



        '
        ' das folgende selektiert die Zelle in der Mitte des aktuell gezeigten Fensters
        ' das verhindert, daß sich plötzlich der Fenster Ausschnitt verändert
        '
        Try
            With appInstance.ActiveWindow
                ziel = (.VisibleRange.Left + .VisibleRange.Width / 2) / boxWidth
            End With

            With appInstance.ActiveSheet
                .Cells(2, ziel).Select()
            End With
        Catch ex As Exception

            With appInstance.ActiveSheet
                .Cells(2, 20).Select()
            End With

        End Try



        appInstance.EnableEvents = formerEE

    End Sub

    Public Function magicBoardZeileIstFrei(ByVal zeile As Integer) As Boolean
        Dim istfrei = True
        Dim ix As Integer = 1
        Dim anzahlP As Integer = ShowProjekte.Count

        If zeile >= 2 Then

            For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

                With kvp.Value
                    If zeile = .tfZeile Then
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



    Public Function magicBoardIstFrei(ByRef mycollection As Collection, ByVal pname As String, ByVal zeile As Integer, _
                                      ByVal spalte As Integer, ByVal laenge As Integer, ByVal anzahlZeilen As Integer) As Boolean
        Dim istfrei = True
        Dim ix As Integer = 1
        Dim anzahlP As Integer = ShowProjekte.Count


        For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

            If pname <> kvp.Key And Not mycollection.Contains(kvp.Key) And kvp.Value.shpUID <> "" Then
                With kvp.Value
                    If .tfZeile >= zeile And .tfZeile <= zeile + anzahlZeilen Then
                        If spalte <= .tfspalte Then
                            If spalte + laenge - 1 >= .tfspalte Then
                                istfrei = False
                                Exit For
                            End If
                        ElseIf spalte <= .tfspalte + .Dauer - 1 Then
                            istfrei = False
                            Exit For
                        End If
                    End If
                End With
            End If

        Next
        magicBoardIstFrei = istfrei
    End Function

    Public Function findeMagicBoardPosition(ByRef mycollection As Collection, ByVal pname As String, ByVal zeile As Integer, ByVal spalte As Integer, ByVal laenge As Integer) As Integer
        Dim lookDown As Boolean = True
        Dim tryoben As Integer, tryunten As Integer
        Dim anzahlzeilen As Integer

        
        Dim hproj As clsProjekt = ShowProjekte.getProject(pname)
        anzahlzeilen = getNeededSpace(hproj) - 1

        ' Konsistenzbedingung prüfen ... 
        If zeile < 2 Then
            zeile = 2
        End If

        If mycollection.Count = 0 Then
            mycollection.Add(pname, pname)
        End If

        If Not magicBoardIstFrei(mycollection, pname, zeile, spalte, laenge, anzahlzeilen) Then
            tryoben = zeile - 1
            tryunten = zeile + 1

            ' jetzt ggf eine neue Position für das Shape suchen - dabei iterierend unten bzw oben suchen
            zeile = tryunten
            lookDown = True

            While Not magicBoardIstFrei(mycollection, pname, zeile, spalte, laenge, anzahlzeilen)
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

        findeMagicBoardPosition = zeile

    End Function

    Sub awinTestSub()
        Call MsgBox("del gedrückt ...")
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="constellationName"></param>
    ''' <remarks></remarks>
    Public Sub awinStoreConstellation(ByVal constellationName As String)


        ' prüfen, ob diese Constellation bereits existiert ..
        If projectConstellations.Contains(constellationName) Then

            Try
                projectConstellations.Remove(constellationName)
            Catch ex As Exception

            End Try

        End If

        Dim newC As New clsConstellation
        With newC
            .constellationName = constellationName
        End With

        Dim newConstellationItem As clsConstellationItem
        For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste
            newConstellationItem = New clsConstellationItem
            With newConstellationItem
                .projectName = kvp.Key
                .show = True
                .Start = kvp.Value.startDate
                .variantName = kvp.Value.variantName
                .zeile = kvp.Value.tfZeile
            End With
            newC.Add(newConstellationItem)
        Next


        Try
            projectConstellations.Add(newC)
        Catch ex As Exception
            Call MsgBox("Fehler bei Add projectConstellations in awinStoreConstellations")
        End Try

    End Sub


    Public Sub awinLoadConstellation(ByVal constellationName As String)
        Dim activeConstellation As New clsConstellation
        Dim hproj As New clsProjekt


        ' prüfen, ob diese Constellation bereits existiert ..
        Try
            activeConstellation = projectConstellations.getConstellation(constellationName)
        Catch ex As Exception
            Call MsgBox(" Projekt-Konstellation " & constellationName & " existiert nicht ")
            Exit Sub
        End Try

        ' die aktuelle Konstellation in "Last" speichern 
        Call awinStoreConstellation("Last")

        ' jetzt wird die activeConstellation in ShowProjekte bzw. NoShowProjekte umgesetzt 
        ' dazu werden erst mal alle Projekte in Showprojekte in Noshowprojekte verschoben ...

        'For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste
        '    NoShowProjekte.Add(kvp.Value)
        'Next
        ShowProjekte.Clear()
        ' jetzt werden die Start-Values entsprechend gesetzt ..

        For Each kvp As KeyValuePair(Of String, clsConstellationItem) In activeConstellation.Liste
            Try
                hproj = AlleProjekte(kvp.Key)
                With hproj
                    .startDate = kvp.Value.Start
                    .StartOffset = 0
                    .tfZeile = kvp.Value.zeile
                End With

                If kvp.Value.show Then

                    Try
                        ShowProjekte.Add(hproj)

                        Dim pname As String
                        Dim tryzeile As Integer
                        With hproj
                            pname = .name
                            tryzeile = .tfZeile
                        End With
                        ' nicht zeichnen - das wird nachher alles auf einen Schlag erledigt ..
                        'Call ZeichneProjektinPlanTafel(pname, tryzeile)

                        'NoShowProjekte.Remove(hproj.name)
                    Catch ex1 As Exception
                        Call MsgBox("Fehler in awinLoadConstellation aufgetreten: " & ex1.Message)
                    End Try

                End If
            Catch ex As Exception
                ' still-to-do:
                ' hier muss das Projekt aus der Datenbank geholt werden ; 
                ' dazu muss diese Sub in ein anderes Modul transferiert werden 
            End Try

        Next

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
    Public Sub awinDeleteMilestoneShapes(ByVal auswahl As Integer)

        Dim worksheetShapes As Excel.Shapes
        Dim shpElement As Excel.Shape

        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim formereO As Boolean = enableOnUpdate
        appInstance.EnableEvents = False
        enableOnUpdate = False

        Try
            worksheetShapes = appInstance.Worksheets(arrWsNames(3)).shapes

            For Each shpElement In worksheetShapes
                With shpElement

                    Select Case auswahl
                        Case 0
                            If .AutoShapeType = MsoAutoShapeType.msoShapeDiamond Or _
                                .AutoShapeType = MsoAutoShapeType.msoShapeOval Or _
                                (.AutoShapeType = MsoAutoShapeType.msoShapeMixed And .Connector = MsoTriState.msoTrue) Or _
                                (.Connector = MsoTriState.msoTrue And .Title = "Dependency") Then
                                .Delete()
                            End If

                            ' Schließen der Status Anzeige Fenster
                            formMilestone.Visible = False
                            formStatus.Visible = False
                            formPhase.Visible = False

                        Case 1
                            If .AutoShapeType = MsoAutoShapeType.msoShapeDiamond Then
                                .Delete()
                            End If
                            formMilestone.Visible = False
                        Case 2
                            If .AutoShapeType = MsoAutoShapeType.msoShapeOval Then
                                .Delete()
                            End If
                            formStatus.Visible = False
                        Case 3
                            If (.AutoShapeType = MsoAutoShapeType.msoShapeMixed And .Connector = MsoTriState.msoTrue And .Title <> "Dependency") Then
                                .Delete()
                            End If
                            formPhase.Visible = False

                        Case 4
                            If (.Connector = MsoTriState.msoTrue And .Title = "Dependency") Then
                                .Delete()
                            End If

                        Case Else

                    End Select

                End With


            Next
        Catch ex As Exception

        End Try

        appInstance.EnableEvents = formerEE
        enableOnUpdate = formereO

    End Sub


    ''' <summary>
    ''' löscht alle MilestoneShapes des Projektes mit Namen pname
    ''' 
    ''' </summary>
    ''' <param name="pname">
    ''' gibt den Projekt-Namen an </param>
    ''' <param name="auswahl">
    ''' 0=alle
    ''' 1=nur meilensteine
    ''' 2=nur Status</param>
    ''' <remarks></remarks>
    Public Sub awinDeleteMilestoneShapes(ByVal pname As String, ByVal auswahl As Integer)

        Dim worksheetShapes As Excel.Shapes
        Dim shpElement As Excel.Shape
        Dim tmpStr(3) As String

        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim formereO As Boolean = enableOnUpdate
        appInstance.EnableEvents = False
        enableOnUpdate = False

        Try
            worksheetShapes = appInstance.Worksheets(arrWsNames(3)).shapes

            For Each shpElement In worksheetShapes
                With shpElement

                    Try
                        tmpStr = .Name.Split(New Char() {"#"}, 3)
                        If tmpStr(0).Trim = pname.Trim Then

                            Select Case auswahl
                                Case 0
                                    If .AutoShapeType = MsoAutoShapeType.msoShapeDiamond Or _
                                        .AutoShapeType = MsoAutoShapeType.msoShapeOval Then
                                        .Delete()
                                    End If


                                Case 1
                                    If .AutoShapeType = MsoAutoShapeType.msoShapeDiamond Then
                                        .Delete()
                                    End If
                                    formMilestone.Visible = False
                                Case 2
                                    If .AutoShapeType = MsoAutoShapeType.msoShapeOval Then
                                        .Delete()
                                    End If
                                    formStatus.Visible = False
                                Case Else

                            End Select

                        End If
                    Catch ex As Exception

                    End Try
                   
                    
                End With


            Next
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
                    ixi = newLength / 2
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
                        newValues(newLength / 2) = newValues(newLength / 2) + newSum - ix
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
        Dim minvorne As Boolean = False, minhinten As Boolean = False, _
            maxvorne As Boolean = False, maxhinten As Boolean = False


        ' Festsetzungen 
        Try
            min = Bedarf.Min
            max = Bedarf.Max
            avg = Bedarf.Sum / Bedarf.Length
            bereich = Bedarf.Length / 4
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
            calcDauerIndays = DateDiff(DateInterval.Day, startDatum, endeDatum) - 1
        Else
            calcDauerIndays = DateDiff(DateInterval.Day, startDatum, endeDatum) + 1
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
            calcDauerIndays = DateDiff(DateInterval.Day, startDatum, endeDatum) - 1
        Else
            calcDauerIndays = DateDiff(DateInterval.Day, startDatum, endeDatum) + 1
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
    ''' erzeugt die monatlichen Budget Werte für ein Projekt
    ''' berechnet aus dem Wert für Erloes, verteilt nach einem Schlüssel, der sich aus Marge und Kostenbedarf ergibt 
    ''' </summary>
    ''' <param name="hproj"></param>
    ''' <remarks></remarks>

    Public Sub awinCreateBudgetWerte(ByRef hproj As clsProjekt)


        Dim costValues() As Double, budgetValues() As Double
        Dim curBudget As Double, avgbudget As Double


        costValues = hproj.getGesamtKostenBedarf
        ReDim budgetValues(costValues.Length - 1)

        curBudget = hproj.Erloes
        avgbudget = curBudget / costValues.Length

        If curBudget > 0 Then
            If costValues.Sum > 0 Then
                Dim pMarge As Double = hproj.ProjectMarge
                For i = 0 To costValues.Length - 1
                    budgetValues(i) = costValues(i) * (1 + pMarge)
                Next
            Else
                For i = 0 To costValues.Length - 1
                    budgetValues(i) = avgbudget
                Next
            End If
        End If


        hproj.budgetWerte = budgetValues


    End Sub

    ''' <summary>
    ''' aktualisiert die Budget werte , wobei die Charakteristik erhalten bleibt 
    ''' Vorbedingung ist, daß das bisherige Budget > 0 Null ist 
    ''' </summary>
    ''' <param name="hproj"></param>
    ''' <param name="newBudget">Gesamt Wert des neuen Budgets</param>
    ''' <remarks></remarks>
    Public Sub awinUpdateBudgetWerte(ByRef hproj As clsProjekt, ByVal newBudget As Double)



        Dim curValues() As Double, budgetValues() As Double
        Dim oldBudget As Double
        Dim faktor As Double

        curValues = hproj.budgetWerte
        ReDim budgetValues(curValues.Length - 1)
        oldBudget = curValues.Sum

        If oldBudget = 0 Then
            Throw New Exception("altes Budget darf beim Update nicht Null sein")
        Else
            If newBudget <= 0 Then
                ' budgetvalues ist bereits auf Null gesetzt  
            Else
                faktor = newBudget / oldBudget
                For i = 0 To curValues.Length - 1
                    budgetValues(i) = curValues(i) * faktor
                Next
            End If

        End If

        hproj.budgetWerte = budgetValues

    End Sub

End Module
