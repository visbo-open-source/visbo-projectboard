Imports ProjectBoardDefinitions
Imports MongoDbAccess
Imports ClassLibrary1
Imports WpfWindow
Imports WPFPieChart
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Excel
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Security.Principal
Imports System.Diagnostics
Imports System.Drawing



'TODO: Führen Sie diese Schritte aus, um das Element auf dem Menüband (XML) zu aktivieren:

'1: Kopieren Sie folgenden Codeblock in die ThisAddin-, ThisWorkbook- oder ThisDocument-Klasse.

'Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
'    Return New Ribbon1()
'End Function

'2. Erstellen Sie Rückrufmethoden im Abschnitt "Menübandrückrufe" dieser Klasse, um Benutzeraktionen
'   zu behandeln, zum Beispiel das Klicken auf eine Schaltfläche. Hinweis: Wenn Sie dieses Menüband aus dem
'   Menüband-Designer exportiert haben, verschieben Sie den Code aus den Ereignishandlern in die Rückrufmethoden, und
'   ändern Sie den Code für die Verwendung mit dem Programmiermodell für die Menübanderweiterung (RibbonX).

'3. Weisen Sie den Steuerelementtags in der Menüband-XML-Datei Attribute zu, um die entsprechenden Rückrufmethoden im Code anzugeben.

'Weitere Informationen erhalten Sie in der Menüband-XML-Dokumentation in der Hilfe zu Visual Studio-Tools für Office.

<Runtime.InteropServices.ComVisible(True)> _
    Public Class Ribbon1
    Implements Office.IRibbonExtensibility

    Private ribbon As Office.IRibbonUI

    Public Sub New()
    End Sub

    Public Function GetCustomUI(ByVal ribbonID As String) As String Implements Office.IRibbonExtensibility.GetCustomUI
        Return GetResourceText("ExcelWorkbook1.Ribbon1.xml")
    End Function

#Region "Menübandrückrufe"
    'Erstellen Sie hier Rückrufmethoden. Weitere Informationen über das Hinzufügen von Rückrufmethoden erhalten Sie, indem Sie das Menüband-XML-Element im Projektmappen-Explorer markieren und dann F1 drücken.
    Public Sub Ribbon_Load(ByVal ribbonUI As Office.IRibbonUI)
        Me.ribbon = ribbonUI
    End Sub


    Sub PTNeueKonstellation(control As IRibbonControl)

        Dim storeConstellationFrm As New frmStoreConstellation
        Dim returnValue As DialogResult
        Dim constellationName As String
        Dim speichernDatenbank As String = "Pt5G2B1"
        Dim request As New Request(awinSettings.databaseName, dbUsername, dbPasswort)

        Dim newConstellationForm As New frmProjPortfolioAdmin



        Call projektTafelInit()

        'If control.Id = speichernDatenbank Then
        '    ' Wenn das Speichern eines Portfolios aus dem Menu Datenbank aufgerufen wird, so werden erneut alle Portfolios aus der Datenbank geholt

        '    If request.pingMongoDb() Then
        '        projectConstellations = request.retrieveConstellationsFromDB()
        '    Else
        '        Call MsgBox("Datenbank-Verbindung ist unterbrochen !")
        '    End If
        'End If

        'Try

        '    With newConstellationForm
        '        .Text = "Portfolio erstellen bzw. ändern"
        '        .portfolioName.Text = currentConstellation
        '        .portfolioName.Visible = True
        '        .Label1.Visible = True
        '        .aKtionskennung = PTtvactions.definePortfolioSE
        '    End With

        '    returnValue = newConstellationForm.ShowDialog

        '    If returnValue = DialogResult.OK Then
        '        'deletedProj = RemoveSelectedProjectsfromDB(deleteProjects.selectedItems)    ' es werden die selektierten Projekte in der DB gespeichert, die Anzahl gespeicherter Projekte sind das Ergebnis

        '    Else
        '        ' returnValue = DialogResult.Cancel

        '    End If

        'Catch ex As Exception

        '    Call MsgBox(ex.Message)
        'End Try


        '
        ' alte Version ; vor dem 26.10.14
        '
        If AlleProjekte.Count > 0 Then
            returnValue = storeConstellationFrm.ShowDialog  ' Aufruf des Formulars zur Eingabe des Portfolios

            If returnValue = DialogResult.OK Then
                constellationName = storeConstellationFrm.ComboBox1.Text

                Call awinStoreConstellation(constellationName)

                ' setzen der public variable, welche Konstellation denn jetzt gesetzt ist
                currentConstellation = constellationName

            End If
        Else
            Call MsgBox("Es sind keine Projekte in der Projekt-Tafel geladen!")
        End If
        ' 
        ' Ende alte Version; vor dem 26.10.14
        '
        enableOnUpdate = True

    End Sub

    Sub PTLadenKonstellation(control As IRibbonControl)

        Dim loadFromDatenbank As String = "PT5G1B1"
        Dim loadConstellationFrm As New frmLoadConstellation

        Dim constellationName As String
        Dim request As New Request(awinSettings.databaseName, dbUsername, dbPasswort)

        Dim initMessage As String = "Es sind dabei folgende Probleme aufgetreten" & vbLf & vbLf

        Dim successMessage As String = initMessage
        Dim returnValue As DialogResult


        Call projektTafelInit()

        ' Wenn das Laden eines Portfolios aus dem Menu Datenbank aufgerufen wird, so werden erneut alle Portfolios aus der Datenbank geholt

        If control.Id = loadFromDatenbank Then
            If request.pingMongoDb() Then
                projectConstellations = request.retrieveConstellationsFromDB()
            Else
                Call MsgBox("Datenbank-Verbindung ist unterbrochen !")
            End If
        End If

        enableOnUpdate = False

        loadConstellationFrm.addToSession.Checked = True
        returnValue = loadConstellationFrm.ShowDialog

        If returnValue = DialogResult.OK Then

            If loadConstellationFrm.addToSession.Checked = True Then
                constellationName = loadConstellationFrm.ListBox1.Text
                Call awinAddConstellation(constellationName, successMessage)
            Else
                constellationName = loadConstellationFrm.ListBox1.Text
                Call awinLoadConstellation(constellationName, successMessage)

                appInstance.ScreenUpdating = False
                'Call diagramsVisible(False)
                Call awinClearPlanTafel()
                Call awinZeichnePlanTafel(False)
                Call awinNeuZeichnenDiagramme(2)
                'Call diagramsVisible(True)
                appInstance.ScreenUpdating = True

                If successMessage.Length > initMessage.Length Then
                    Call MsgBox(constellationName & " wurde geladen ..." & vbLf & vbLf & successMessage)
                Else
                    'Call MsgBox(constellationName & " wurde geladen ...")
                End If

                ' setzen der public variable, welche Konstellation denn jetzt gesetzt ist
                currentConstellation = constellationName
            End If



        End If
        enableOnUpdate = True

    End Sub
    Sub PTRemoveKonstellation(control As IRibbonControl)

        Dim ButtonId As String = control.Id

        Dim remConstellationFrm As New frmRemoveConstellation
        Dim constellationName As String

        Dim returnValue As DialogResult

        Call projektTafelInit()


        Dim deleteDatenbank As String = "Pt5G3B1"
        Dim request As New Request(awinSettings.databaseName, dbUsername, dbPasswort)

        Dim removeFromDB As Boolean

        If control.Id = deleteDatenbank Then
            removeFromDB = True
            If request.pingMongoDb() Then
                projectConstellations = request.retrieveConstellationsFromDB()
            Else
                Call MsgBox("Datenbank-Verbindung ist unterbrochen !")
                removeFromDB = False
            End If
        Else
            removeFromDB = False
        End If

        enableOnUpdate = False

        returnValue = remConstellationFrm.ShowDialog

        If returnValue = DialogResult.OK Then
            constellationName = remConstellationFrm.ListBox1.Text

            Call awinRemoveConstellation(constellationName, removeFromDB)
            Call MsgBox(constellationName & " wurde gelöscht ...")

            If constellationName = currentConstellation Then

                ' aktuelle Konstellation unter dem Namen 'Last' speichern
                Call storeSessionConstellation(ShowProjekte, "Last")
                currentConstellation = "Last"
            Else
                ' aktuelle Konstellation bleibt unverändert
            End If


        End If
        enableOnUpdate = True

    End Sub


    Sub PT5StoreProjects(control As IRibbonControl)

        Dim storedProj As Integer = 0

        Call projektTafelInit()

        Try
            If AlleProjekte.Count > 0 Then

                storedProj = StoreSelectedProjectsinDB()    ' es werden die selektierten Projekte einschl. der geladenen Varianten 
                ' in der DB gespeichert, die Anzahl gespeicherter Projekte sind das Ergebnis

                If storedProj = 0 Then
                    Call MsgBox("Es wurde kein Projekt selektiert. " & vbLf & "Alle Projekte speichern?", MsgBoxStyle.OkCancel)

                    If MsgBoxResult.Ok = vbOK Then
                        Call StoreAllProjectsinDB()
                    End If
                Else
                    'Call MsgBox("Es wurden " & storedProj & " Projekte gespeichert!")
                End If

            Else
                Call MsgBox("keine Projekte zu speichern ...")
            End If
        Catch ex As Exception

            Call MsgBox(ex.Message)
        End Try

        Call awinDeSelect()

    End Sub
    ''' <summary>
    ''' löscht die ausgewählten Projekte aus der Datenbank 
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub PT5DeleteProjectsInDB(control As IRibbonControl)


        Dim deletedProj As Integer = 0
        Dim returnValue As DialogResult

        'Dim deleteProjects As New frmDeleteProjects
        Dim deleteProjects As New frmProjPortfolioAdmin

        Try

            With deleteProjects
                .Text = "Projekte, Varianten bzw. Snapshots in der Datenbank löschen"
                .aKtionskennung = PTTvActions.delFromDB
                .OKButton.Text = "Löschen"
                .portfolioName.Visible = False
                .Label1.Visible = False
            End With

            returnValue = deleteProjects.ShowDialog

            ' die Operation ist bereits ausgeführt - deswegen muss hier nichts mehr unterschieden werden 

            If returnValue = DialogResult.OK Then
                ' everything is done ... 

            Else
                ' everything is done ... 

            End If

        Catch ex As Exception

            Call MsgBox(ex.Message)
        End Try



    End Sub

    ''' <summary>
    ''' löscht alles, was aktuell in der Session ist 
    ''' Projekte, Charts, Shapes ... 
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub PT6G3ClearSession(control As IRibbonControl)

        Call projektTafelInit()

        ' Bestätigungs-Fenster aufrufen 
        Dim bestaetigeLoeschen As New frmconfirmDeletePrj
        Dim returnValue As DialogResult
        Dim allShapes As Excel.Shapes

        bestaetigeLoeschen.botschaft = "Bitte bestätigen Sie das Löschen der kompletten Session"
        returnValue = bestaetigeLoeschen.ShowDialog

        If returnValue = DialogResult.Cancel Then
            ' nichts tun
        Else
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
            selectedProjekte.Clear()
            ImportProjekte.Clear()
            DiagramList.Clear()
            awinButtonEvents.Clear()

            allDependencies.Clear()
            projectboardShapes.clear()
            ' Session gelöscht

            appInstance.EnableEvents = True
            enableOnUpdate = True
        End If

    End Sub

    Sub PT6DeleteCharts(control As IRibbonControl)

        Dim anzDiagrams As Integer
        Dim chtobj As Excel.ChartObject
        Dim i As Integer = 1

        Call projektTafelInit()

        With CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet)

            anzDiagrams = CInt(CType(.ChartObjects, Excel.ChartObjects).Count)

            While i <= anzDiagrams

                chtobj = CType(.ChartObjects(1), Excel.ChartObject)
                Call awinDeleteChart(chtobj)
                i = i + 1

            End While


        End With


    End Sub
    Sub PT0SaveCockpit(control As IRibbonControl)


        Dim i As Integer = 1
        Dim storeCockpitFrm As New frmStoreCockpit
        Dim returnValue As DialogResult
        Dim cockpitName As String
        Try

            Call projektTafelInit()

            Call awinDeSelect()

            Dim anzDiagrams As Integer = CType(appInstance.Worksheets(arrWsNames(3)).ChartObjects, Excel.ChartObjects).Count

            If anzDiagrams > 0 Then


                ' hier muss die Auswahl des Names für das Cockpit erfolgen

                returnValue = storeCockpitFrm.ShowDialog  ' Aufruf des Formulars zur Eingabe des Cockpitnamens

                If returnValue = DialogResult.OK Then

                    cockpitName = storeCockpitFrm.ComboBox1.Text

                    appInstance.ScreenUpdating = False

                    enableOnUpdate = False

                    Call awinStoreCockpit(cockpitName)

                    enableOnUpdate = True

                    appInstance.ScreenUpdating = True
                Else


                End If
                ' hier muss eventuell ein Neuzeichnen erfolgen
            Else
                Call MsgBox("Es ist kein Chart angezeigt")
            End If

        Catch ex As Exception
            Throw New ArgumentException("PT0SaveCockpit: Fehler:  ", ex.Message)
        End Try

    End Sub

    Sub PT0ShowCockpit(control As IRibbonControl)


        Dim i As Integer = 1
        Dim loadCockpitFrm As New frmLoadCockpit
        Dim returnValue As DialogResult
        Dim cockpitName As String

        Call projektTafelInit()

        If ShowProjekte.Count > 0 Then

            If showRangeRight - showRangeLeft > 5 Then

                Dim awinSelection As Excel.ShapeRange

                Try
                    'awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
                    awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
                Catch ex As Exception
                    awinSelection = Nothing
                End Try

                appInstance.EnableEvents = False
                enableOnUpdate = False

                ' hier muss die Auswahl des Names für das Cockpit erfolgen

                returnValue = loadCockpitFrm.ShowDialog  ' Aufruf des Formulars zur Eingabe des Cockpitnamens

                If returnValue = DialogResult.OK Then

                    cockpitName = loadCockpitFrm.ListBox1.Text

                    appInstance.ScreenUpdating = False

                    Try
                        Call awinLoadCockpit(cockpitName)

                        ' nur wenn ein Projekt selektiert wurde, werden die Projekt-Charts aktualisiert
                        If Not awinSelection Is Nothing Then


                            If awinSelection.Count = 1 Then
                                Dim singleShp As Excel.Shape
                                Dim hproj As clsProjekt

                                ' jetzt die Aktion durchführen ...
                                singleShp = awinSelection.Item(1)

                                Try
                                    hproj = ShowProjekte.getProject(singleShp.Name)
                                Catch ex As Exception
                                    Call MsgBox("Projekt nicht gefunden ..." & singleShp.Name)
                                    Exit Sub
                                End Try

                                Call aktualisiereCharts(hproj, True)

                                Call awinDeSelect()
                            End If

                        End If

                        Call awinNeuZeichnenDiagramme(9)

                    Catch ex As Exception
                        Call MsgBox("Fehler beim Laden ..")
                    End Try




                    appInstance.ScreenUpdating = True

                Else
                    appInstance.ScreenUpdating = True

                End If
            Else
                Call MsgBox("Bitte wählen Sie einen Zeitraum aus!")
            End If
        Else
            Call MsgBox("Es sind noch keine Projekte geladen!")
        End If

        ' hier muss eventuell ein Neuzeichnen erfolgen
        enableOnUpdate = True
        appInstance.EnableEvents = True
    End Sub


    ''' <summary>
    ''' wird aktuell verwendet , um eine Stelle für Testen bestimmter Funktionalitäten zu haben
    ''' ohne dass eine neue Ribbon Erweiterung gemacht werden muss
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub awinTestNewFunctions(control As IRibbonControl)
        'Call MsgBox("Anzahl Aufrufe: " & anzahlCalls)

        Dim awinSelection As Excel.ShapeRange
        Dim i As Integer
        Dim hproj As clsProjekt
        Dim singleShp As Excel.Shape
        Dim ausgabeString As String = ""
        Dim vglWert As Integer
        Dim curCoord() As Double
        Dim key As String

        Call projektTafelInit()


        enableOnUpdate = False



        Try
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)

        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            ' Es muss mindestens 1 Projekt selektiert sein
            For i = 1 To awinSelection.Count

                singleShp = awinSelection.Item(i)
                key = singleShp.Name
                hproj = ShowProjekte.getProject(singleShp.Name)
                vglWert = calcYCoordToZeile(singleShp.Top)
                curCoord = projectboardShapes.getCoord(singleShp.Name)

                ausgabeString = ausgabeString & hproj.name & ": " & hproj.tfZeile.ToString & _
                                 " - " & vglWert.ToString & "; " & _
                                 calcXCoordToDate(singleShp.Left).ToShortDateString & " vs. " & hproj.startDate.ToShortDateString & _
                                 " vs. " & calcXCoordToDate(curCoord(1)).ToShortDateString & singleShp.Left.ToString & vbLf


            Next i


        End If

        Call awinDeSelect()
        Call MsgBox(ausgabeString)

        enableOnUpdate = True

        ' für andere Zwecke ... 

        'For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

        '    If Not kvp.Value.isConsistent Then
        '        Call MsgBox("inkonsistenz: " & kvp.Key)
        '        ok = False
        '    End If

        'Next

        'If ok Then
        '    Call MsgBox("keine Inkonsistenz gefunden ...")
        'End If

        'Dim anzProjekte As Integer = ShowProjekte.Liste.Count
        'Dim anzShapes As Integer = ShowProjekte.shpListe.Count
        'Call MsgBox("Test ---" & vbLf & _
        '            "Projekte: " & anzProjekte & _
        '            "Shapes  : " & anzShapes)

        'Dim hws As Excel.Worksheet
        'hws = appInstance.Worksheets(arrWsNames(11))
        'appInstance.EnableEvents = False

        'With hws
        '    .Unprotect()
        '    .Visible = True
        '    .Protect()
        'End With


        'appInstance.EnableEvents = True
        'hws.Activate()

    End Sub




    ''' <summary>
    ''' Rename Funktion für ein Projekt
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G1Rename(control As IRibbonControl)

        Dim singleShp As Excel.Shape
        Dim request As New Request(awinSettings.databaseName, dbUsername, dbPasswort)
        'Dim pName As String, variantName As String
        'Dim shapeText As String

        Dim tmpshapes As Excel.Shapes
        Dim oldKey As String, newKey As String
        Dim erg As String = ""
        Dim atleastOne As Boolean = False
        Dim hproj As clsProjekt

        Call projektTafelInit()

        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False


        enableOnUpdate = False

        Try
            tmpshapes = CType(CType(appInstance.ActiveSheet, Excel.Worksheet).Shapes, Excel.Shapes)
        Catch ex As Exception
            tmpshapes = Nothing
        End Try

        If Not tmpshapes Is Nothing Then

            ' jetzt die Aktion durchführen ...
            Try
                For Each singleShp In tmpshapes

                    Dim shapeArt As Integer
                    shapeArt = kindOfShape(singleShp)

                    With singleShp



                        If isProjectType(shapeArt) Then

                            ' jetzt muss Pname und Variant-Name ermittel werde 
                            Try
                                hproj = ShowProjekte.getProject(.Name)


                                If hproj.getShapeText <> .TextFrame2.TextRange.Text Then
                                    ' das Shape wurde vom Nutzer umbenannt 
                                    atleastOne = True


                                    Dim oldPname As String = hproj.name
                                    Dim oldVname As String = hproj.variantName
                                    Dim tmpstr(5) As String
                                    Dim newPname As String = ""
                                    Dim newVname As String = ""
                                    tmpstr = .TextFrame2.TextRange.Text.Trim.Split(New Char() {CChar("("), CChar(")")}, 3)

                                    newPname = tmpstr(0)
                                    If tmpstr.Length > 1 Then
                                        newVname = tmpstr(1)
                                    End If

                                    Try


                                        If request.pingMongoDb() Then

                                            If ShowProjekte.contains(newPname) Or request.projectNameAlreadyExists(newPname, hproj.variantName) Or Len(newPname.Trim) = 0 Or IsNumeric(newPname) Then

                                                ' ungültiger Name - alten Namen wiederherstellen 
                                                .TextFrame2.TextRange.Text = hproj.getShapeText
                                                erg = erg & oldPname & " bleibt, " & newPname & " ungültig oder existiert bereits in DB" & vbLf
                                            Else
                                                ' der neue Name ist gültig 
                                                .Name = newPname

                                                oldKey = calcProjektKey(hproj)
                                                newKey = calcProjektKey(newPname, hproj.variantName)
                                                With hproj
                                                    .name = newPname
                                                End With

                                                ShowProjekte.Remove(oldPname)
                                                hproj.timeStamp = Date.Now
                                                ShowProjekte.Add(hproj)
                                                AlleProjekte.Remove(oldKey)
                                                AlleProjekte.Add(newKey, hproj)

                                                erg = erg & oldPname & " -> " & newPname & vbLf

                                            End If
                                        Else
                                            'Call MsgBox("Datenbank-Verbindung ist unterbrochen !")
                                            .TextFrame2.TextRange.Text = oldPname
                                            erg = erg & oldPname & " bleibt, " & newPname & " ungültig, DB ist nicht aktiv" & vbLf
                                        End If

                                    Catch ex1 As Exception
                                        Call MsgBox(ex1.Message)
                                        .TextFrame2.TextRange.Text = oldPname
                                        erg = erg & oldPname & " bleibt, " & newPname & " ungültig" & vbLf
                                    End Try

                                End If
                            Catch ex As Exception
                                Call MsgBox("Fehler : zu Shape mit Namen " & .Name & " gibt es kein Projekt!")
                            End Try



                        End If
                    End With
                Next
            Catch ex As Exception
                Call MsgBox("Aktion im Extended Mode nicht unterstützt ...")
            End Try


        End If

        If atleastOne Then
            Call MsgBox(erg)
        Else
            Call MsgBox("es hat kein Rename stattgefunden")
        End If

        enableOnUpdate = True
        appInstance.EnableEvents = formerEE
    End Sub

    Sub PT2ProjektNeu(control As IRibbonControl)

        Dim ProjektEingabe As New frmProjektEingabe1
        Dim returnValue As DialogResult
        Dim zeile As Integer = 0
        Dim request As New Request(awinSettings.databaseName, dbUsername, dbPasswort)

        Call projektTafelInit()

        enableOnUpdate = False


        returnValue = ProjektEingabe.ShowDialog

        If returnValue = DialogResult.OK Then
            With ProjektEingabe

                If request.pingMongoDb() Then

                    If Not request.projectNameAlreadyExists(projectname:=.projectName.Text, variantname:="") Then

                        ' Projekt existiert noch nicht in der DB, kann also eingetragen werden

                        Call TrageivProjektein(.projectName.Text, .vorlagenDropbox.Text, CDate(.calcProjektStart), _
                                           CDate(.calcProjektEnde), CType(.Erloes.Text, Double), zeile, _
                                           CType(.sFit.Text, Double), CType(.risiko.Text, Double), CDbl(.volume.Text))
                    Else
                        Call MsgBox(" Projekt '" & .projectName.Text & "' existiert bereits in der Datenbank!")
                    End If

                Else

                    Call MsgBox("Datenbank- Verbindung ist unterbrochen !")
                    appInstance.ScreenUpdating = True

                    ' Projekt soll trotzdem angezeigt werden
                    Call TrageivProjektein(.projectName.Text, .vorlagenDropbox.Text, CDate(.calcProjektStart), _
                                           CDate(.calcProjektEnde), CType(.Erloes.Text, Double), zeile, _
                                           CType(.sFit.Text, Double), CType(.risiko.Text, Double), CDbl(.volume.Text))

                End If

            End With
        End If


        enableOnUpdate = True

    End Sub

    ''' <summary>
    ''' eine neue Variante anlegen 
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub PT2VarianteNeu(control As IRibbonControl)

        Dim singleShp As Excel.Shape
        Dim hproj As clsProjekt
        Dim awinSelection As Excel.ShapeRange
        Dim neueVariante As New frmCreateNewVariant
        Dim resultat As DialogResult
        Dim request As New Request(awinSettings.databaseName, dbUsername, dbPasswort)
        Dim newproj As clsProjekt
        Dim key As String
        Dim phaseList As New Collection
        Dim milestoneList As New Collection


        Call projektTafelInit()

        enableOnUpdate = False

        Try
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            If awinSelection.Count = 1 Then
                ' jetzt die Aktion durchführen ...
                singleShp = awinSelection.Item(1)

                Try
                    hproj = ShowProjekte.getProject(singleShp.Name)
                    phaseList = projectboardShapes.getPhaseList(hproj.name)
                    milestoneList = projectboardShapes.getMilestoneList(hproj.name)
                Catch ex As Exception
                    Call MsgBox("Projekt " & singleShp.Name & " nicht gefunden ...")
                    enableOnUpdate = True
                    Exit Sub
                End Try

                ' enableevents wird hier nicht false gesetzt; wenn dann wird das im Formular gemacht 
                ' screenupdating wird hier ebenso nicht auf false gesetzt 

                ' jetzt wird hier das Formular aufgerufen, wo eine neue Variante eingegeben werden kann 
                With neueVariante
                    .projektName.Text = hproj.name
                    .variantenName.Text = hproj.variantName
                    .newVariant.Text = ""
                End With

                resultat = neueVariante.ShowDialog
                If resultat = DialogResult.OK Then

                    newproj = New clsProjekt
                    hproj.copyTo(newproj)

                    With newproj
                        .name = hproj.name
                        .variantName = neueVariante.newVariant.Text
                        .ampelErlaeuterung = hproj.ampelErlaeuterung
                        .ampelStatus = hproj.ampelStatus
                        .timeStamp = Date.Now
                        .shpUID = hproj.shpUID
                        .tfZeile = hproj.tfZeile
                        .Status = ProjektStatus(0)
                        If Not IsNothing(hproj.budgetWerte) Then
                            .budgetWerte = hproj.budgetWerte
                        End If

                    End With

                    ' jetzt muss die bisherige Variante aus Showprojekte rausgenommen werden ..
                    ShowProjekte.Remove(hproj.name)

                    ' die neue Variante wird aufgenommen
                    key = calcProjektKey(newproj)
                    AlleProjekte.Add(key, newproj)
                    ShowProjekte.Add(newproj)

                    ' wenn bestimmte Projekte beim Suchen nach einem Platz nicht berücksichtigt werden sollen,
                    ' dann müssen sie in einer Collection an ZeichneProjektinPlanTafel übergeben werden 
                    Try

                        Dim tmpCollection As New Collection
                        Call ZeichneProjektinPlanTafel(tmpCollection, newproj.name, newproj.tfZeile, phaseList, milestoneList)

                    Catch ex As Exception

                        Call MsgBox("Fehler bei Zeichnen Projekt: " & ex.Message)

                    End Try


                End If

            Else
                Call MsgBox("bitte nur ein Projekt selektieren")

            End If
        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If

        enableOnUpdate = True



    End Sub

    ''' <summary>
    ''' aktiviert die selektierte Variante 
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub PT2VarianteAktiv(control As IRibbonControl)

        Dim deletedProj As Integer = 0
        'Dim returnValue As DialogResult

        'Dim activateVariant As New frmDeleteProjects
        Dim activateVariant As New frmProjPortfolioAdmin

        Try

            With activateVariant
                .Text = "Variante aktivieren"
                .aKtionskennung = PTTvActions.activateV
                .OKButton.Visible = False
                '.OKButton.Text = "Löschen"
                .portfolioName.Visible = False
                .Label1.Visible = False
            End With

            'returnValue = activateVariant.ShowDialog
            activateVariant.Show()

            'If returnValue = DialogResult.OK Then
            '    'deletedProj = RemoveSelectedProjectsfromDB(deleteProjects.selectedItems)    ' es werden die selektierten Projekte in der DB gespeichert, die Anzahl gespeicherter Projekte sind das Ergebnis

            'Else
            '    ' returnValue = DialogResult.Cancel

            'End If

        Catch ex As Exception

            Call MsgBox(ex.Message)
        End Try


    End Sub

    ''' <summary>
    ''' die Variante, die übernommen werden soll, muss bereits in der Showprojekte sein und selektiert sein
    ''' Das Projekt wird zur Standard-Variante 
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub PT2VarianteUebernehmen(control As IRibbonControl)

        Dim awinSelection As Excel.ShapeRange
        Dim i As Integer
        Dim hproj As clsProjekt
        Dim singleShp As Excel.Shape
        'Dim ausgabeString As String = ""
        'Dim vglWert As Integer
        'Dim curCoord() As Double
        Dim key As String

        Call projektTafelInit()


        enableOnUpdate = False



        Try
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)

        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            ' Es muss mindestens 1 Projekt selektiert sein
            For i = 1 To awinSelection.Count

                singleShp = awinSelection.Item(i)
                hproj = ShowProjekte.getProject(singleShp.Name)

                ' das Projekt zur Standard Variante machen 
                If hproj.variantName <> "" Then

                    ' die aktuelle Variante aus der AlleProjekte rausnehmen 
                    key = calcProjektKey(hproj)
                    AlleProjekte.Remove(key)

                    ' das bisherige Standard Projekt aus der AlleProjekte rausnehmen 
                    key = calcProjektKey(hproj.name, "")
                    AlleProjekte.Remove(key)

                    'jetzt die aktuelle Variante zur Standard Variante machen 
                    hproj.variantName = ""
                    hproj.timeStamp = Date.Now
                    If hproj.Status = ProjektStatus(0) Then
                        hproj.Status = ProjektStatus(1)
                    End If

                    ' die "neue" Standard Variante in AlleProjekte aufnehmen 
                    AlleProjekte.Add(key, hproj)

                    ' wenn bestimmte Projekte beim Suchen nach einem Platz nicht berücksichtigt werden sollen,
                    ' dann müssen sie in einer Collection an ZeichneProjektinPlanTafel übergeben werden 
                    Dim tmpCollection As New Collection
                    Call ZeichneProjektinPlanTafel(tmpCollection, hproj.name, hproj.tfZeile, tmpCollection, tmpCollection)

                End If



            Next i


        End If

    End Sub

    ''' <summary>
    ''' Es werden Projekte, die Varianten haben angezeigt in einem TreeView
    ''' Hier können Varianten ausgewählt werden, die gelöscht werden sollen
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub PT2VarianteLoeschen(control As IRibbonControl)

        Dim deletedProj As Integer = 0
        'Dim returnValue As DialogResult

        'Dim activateVariant As New frmDeleteProjects
        Dim deleteVariant As New frmProjPortfolioAdmin

        Try

            With deleteVariant
                .Text = "Variante löschen"
                .aKtionskennung = PTTvActions.deleteV
                .OKButton.Visible = True
                .OKButton.Text = "Löschen"
                .portfolioName.Visible = False
                .Label1.Visible = False
            End With

            'returnValue = activateVariant.ShowDialog
            deleteVariant.Show()

            'If returnValue = DialogResult.OK Then
            '    'deletedProj = RemoveSelectedProjectsfromDB(deleteProjects.selectedItems)    ' es werden die selektierten Projekte in der DB gespeichert, die Anzahl gespeicherter Projekte sind das Ergebnis

            'Else
            '    ' returnValue = DialogResult.Cancel

            'End If

        Catch ex As Exception

            Call MsgBox(ex.Message)
        End Try


    End Sub

    Sub PT5changeTimeSpan(control As IRibbonControl)

        Dim mvTimeSpan As New frmMoveTimeSpan
        'Dim returnValue As DialogResult

        Call projektTafelInit()

        appInstance.EnableEvents = False

        'returnValue = mvTimeSpan.Showdialog
        ' in dieser auskommentierten Variante ist es sehr langsam ... deshalb als modales Fenster
        If showRangeRight <> showRangeLeft Then
            mvTimeSpan.Show()
        Else
            Call MsgBox("bitte zuerst eine Zeitspanne definieren")
        End If


        appInstance.EnableEvents = True


    End Sub

    Sub PTDefineDependencies(control As IRibbonControl)

        Dim defineDependencies As New frmDependencies
        Dim result As DialogResult
        Dim awinSelection As Excel.ShapeRange



        Call projektTafelInit()

        enableOnUpdate = False

        If ShowProjekte.Count > 0 Then

            Try
                awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
            Catch ex As Exception
                awinSelection = Nothing
            End Try

            If Not awinSelection Is Nothing Then

                If awinSelection.Count > 1 Then

                    result = defineDependencies.ShowDialog()
                Else

                    Call MsgBox("Bitte zunächst  mindestens zwei Projekte selektieren!")
                End If
            Else
                Call MsgBox("Bitte zunächst mindestens zwei Projekte selektieren!")
            End If

        Else
            Call MsgBox("Es sind keine Projekte geladen!")
        End If

        enableOnUpdate = True

    End Sub

    ''' <summary>
    ''' Ressourcen und Kosten eines Projektes bearbeiten 
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G1Resources(control As IRibbonControl)

        Dim singleShp As Excel.Shape
        Dim pname As String
        Dim awinSelection As Excel.ShapeRange
        Dim hproj As clsProjekt

        Call projektTafelInit()



        ' es wird vbeim Betreten der Tabelle2 nochmal auf False gesetzt ... und insbesondere bei Activate Tabelle1 (!) auf true gesetzt, nicht vorher wieder
        enableOnUpdate = False


        Try
            'awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            If awinSelection.Count = 1 Then
                ' jetzt die Aktion durchführen ...
                singleShp = awinSelection.Item(1)

                Try
                    hproj = ShowProjekte.getProject(singleShp.Name)
                    pname = hproj.name
                Catch ex As Exception
                    Call MsgBox(" Fehler in EditProject " & singleShp.Name & " , Modul: Tom2G1Resources")
                    enableOnUpdate = True
                    Exit Sub
                End Try

                ' jetzt werden die Daten aus hproj in Edit Ressourcen worksheet geschrieben ... 
                appInstance.ScreenUpdating = False
                Call awinStoreProjForEditRess(hproj)
                Dim oldShpID As Integer = CInt(hproj.shpUID)

                ' hier wird das non-modale Dialog Fenster aufgerufen 
                Dim confirmEdit As New frmConfirmEditRess

                confirmEdit.selectedProject = hproj.name
                confirmEdit.Show()

                With CType(appInstance.Worksheets(arrWsNames(5)), Excel.Worksheet)
                    .Activate()
                End With



            Else
                Call MsgBox("bitte nur ein Projekt selektieren")
                'For Each singleShp In awinSelection
                '    With singleShp
                '        If .AutoShapeType = MsoAutoShapeType.msoShapeRoundedRectangle Then
                '            nrSelPshp = nrSelPshp + 1
                '            SID = .ID.ToString
                '        End If
                '    End With
                'Next
            End If
        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If


        ' das muss hier de-aktiviert werden, weil durch non-modalen Aufruf des Formulars enableonupdate wieder auf true gesetzt wird 
        ' enableOnUpdate = True



    End Sub

    ''' <summary>
    ''' Attribute eines Projektes bearbeiten 
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G1Attribute(control As IRibbonControl)

        Dim ProjektAendern As New frmProjektAendern
        Dim returnValue As DialogResult

        Dim singleShp As Excel.Shape

        Dim awinSelection As Excel.ShapeRange
        Dim hproj As clsProjekt
        Dim databaseName As String = awinSettings.databaseName

        Call projektTafelInit()

        enableOnUpdate = False

        Try
            'awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            If awinSelection.Count = 1 Then
                ' jetzt die Aktion durchführen ...
                singleShp = awinSelection.Item(1)

                Try
                    hproj = ShowProjekte.getProject(singleShp.Name)

                    ' jetzt werden die Werte im Fenster vorbesetzt ...
                    With ProjektAendern
                        .projectName.Text = hproj.name
                        .vorlagenName.Text = hproj.VorlagenName
                        For Each kvp As KeyValuePair(Of Integer, clsBusinessUnit) In businessUnitDefinitions
                            .businessUnit.Items.Add(kvp.Value.name)
                        Next
                        .businessUnit.Text = hproj.businessUnit
                        .Erloes.Text = hproj.Erloes.ToString
                        .risiko.Text = hproj.Risiko.ToString("0.0")
                        .sFit.Text = hproj.StrategicFit.ToString("0.0")
                    End With
                    ' Aufruf Dialog Fenster 
                    returnValue = ProjektAendern.ShowDialog

                    If returnValue = DialogResult.OK Then
                        With hproj
                            .timeStamp = Date.Now

                            If .Erloes <> CType(ProjektAendern.Erloes.Text, Double) Then
                                If .Erloes = 0 Then
                                    .Erloes = CType(ProjektAendern.Erloes.Text, Double)

                                    ' Workaround: 
                                    Dim tmpValue As Integer = hproj.dauerInDays
                                    Call awinCreateBudgetWerte(hproj)
                                Else
                                    Try
                                        Call awinUpdateBudgetWerte(hproj, CType(ProjektAendern.Erloes.Text, Double))
                                        .Erloes = CType(ProjektAendern.Erloes.Text, Double)
                                    Catch ex As Exception
                                        .Erloes = CType(ProjektAendern.Erloes.Text, Double)
                                        ' Workaround: 
                                        Dim tmpValue As Integer = hproj.dauerInDays
                                        Call awinCreateBudgetWerte(hproj)
                                    End Try

                                End If
                            End If

                            .StrategicFit = CType(ProjektAendern.sFit.Text, Double)
                            .Risiko = CType(ProjektAendern.risiko.Text, Double)
                            .businessUnit = CType(ProjektAendern.businessUnit.Text, String)

                        End With

                        Call awinNeuZeichnenDiagramme(5)
                    End If

                Catch ex As Exception
                    Call MsgBox(" Fehler in EditProject " & singleShp.Name & " , Modul: Tom2G1Resources")
                    Exit Sub
                End Try


            Else
                Call MsgBox("bitte nur ein Projekt selektieren")
                'For Each singleShp In awinSelection
                '    With singleShp
                '        If .AutoShapeType = MsoAutoShapeType.msoShapeRoundedRectangle Then
                '            nrSelPshp = nrSelPshp + 1
                '            SID = .ID.ToString
                '        End If
                '    End With
                'Next
            End If
        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If

        enableOnUpdate = True


    End Sub

    ''' <summary>
    ''' earliest und latest Start eines Projektes ändern 
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G1EarliestLatestStart(control As IRibbonControl)

        Dim setStartEnd As New frmEarliestLatestStart

        Dim returnValue As DialogResult
        Dim awinSelection As Excel.ShapeRange
        Dim i As Integer
        Dim hproj As clsProjekt
        Dim singleShp As Excel.Shape
        Dim pname As String
        Dim todoListe As New Collection
        Dim errMessage As String = ""
        Dim initMsg As String = "für folgende Projekte nicht zulässig, da sie nicht mehr Status=geplant haben: "

        Call projektTafelInit()

        ' es wird vbeim Betreten der Tabelle2 nochmal auf False gesetzt ... und insbesondere bei Activate Tabelle1 (!) auf true gesetzt, nicht vorher wieder
        enableOnUpdate = False

        ' Änderung 2.7.14 tk : Vorbedingung sicherstellen: nur Projekte, die noch nicht beauftragt sind, können noch verschoben und 
        ' werden
        '
        Try
            'awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            ' Es muss mindestens 1 Projekt selektiert sein
            For i = 1 To awinSelection.Count

                singleShp = awinSelection.Item(i)

                Try
                    hproj = ShowProjekte.getProject(singleShp.Name)
                    pname = hproj.name
                Catch ex As Exception
                    Call MsgBox(" Fehler! Projekt " & singleShp.Name & " nicht im Hauptspeicher")
                    enableOnUpdate = True
                    Exit Sub
                End Try

                If hproj.Status = ProjektStatus(0) Then
                    ' nur dann macht das Setzen von earliest / latest Sinn ...

                    todoListe.Add(hproj.name)

                    If i = 1 Then

                        ' jetzt die Aktion durchführen ...

                        With setStartEnd

                            .EarliestStart.Value = hproj.earliestStart
                            .LatestStart.Value = hproj.latestStart

                        End With


                    Else

                        With setStartEnd

                            If .EarliestStart.Value <> hproj.earliestStart Or .LatestStart.Value <> hproj.latestStart Then

                                .EarliestStart.Value = 0
                                .LatestStart.Value = 0

                            End If

                        End With


                    End If
                Else
                    errMessage = errMessage & vbLf & hproj.name
                End If

            Next i

            If todoListe.Count > 0 Then

                returnValue = setStartEnd.ShowDialog

                If returnValue = DialogResult.OK Then

                    For i = 1 To todoListe.Count

                        pname = CStr(todoListe.Item(i))

                        ' jetzt die Aktion durchführen ...
                        Try
                            hproj = ShowProjekte.getProject(pname)
                            With setStartEnd

                                hproj.earliestStart = .EarliestStart.Value
                                hproj.latestStart = .LatestStart.Value
                                hproj.earliestStartDate = hproj.startDate.AddMonths(.EarliestStart.Value)
                                hproj.latestStartDate = hproj.startDate.AddMonths(.LatestStart.Value)

                            End With
                        Catch ex As Exception
                            Call MsgBox(" Fehler! Projekt " & pname & " earliest/latest kann nicht gesetzt werden")
                            enableOnUpdate = True
                            Exit Sub
                        End Try

                    Next i

                    Call MsgBox("ok, frühester und spätester Start gesetzt")

                ElseIf returnValue = DialogResult.Cancel Then
                    'Call MsgBox("Default soll gelten")

                End If

            End If

            If errMessage.Length > 0 Then
                Call MsgBox(initMsg & vbLf & errMessage)
            End If

        Else

            Call MsgBox("Es muss mindestens ein Projekt selektiert sein")

        End If

        Call awinDeSelect()

        'appInstance.ScreenUpdating = True
        enableOnUpdate = True


    End Sub

    ''' <summary>
    ''' Projekt ins Noshow stellen  
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    ''' 
    Sub Tom2G1NoShow(control As IRibbonControl)

        Dim singleShp As Excel.Shape
        'Dim SID As String

        Dim awinSelection As Excel.ShapeRange

        Call projektTafelInit()

        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False

        enableOnUpdate = False

        Try
            'awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            ' jetzt die Aktion durchführen ...

            For Each singleShp In awinSelection

                Dim shapeArt As Integer
                shapeArt = kindOfShape(singleShp)

                With singleShp
                    If isProjectType(shapeArt) Then

                        Call awinShowNoShowProject(pname:=.Name)

                    End If
                End With
            Next

        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If

        enableOnUpdate = True
        appInstance.EnableEvents = formerEE
    End Sub

    ''' <summary>
    ''' neues Formular zur Auswahl Phasen/Meilensteine/Rollen/Kosten anzeigen
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub NameHierarchySelAction(control As IRibbonControl)

        Dim nameFormular As New frmNameSelection
        Dim hryFormular As New frmHierarchySelection
        Dim awinSelection As Excel.ShapeRange

        Call projektTafelInit()

        enableOnUpdate = False
        appInstance.EnableEvents = False


        ' gibt es überhaupt Objekte, zu denen man was anzeigen kann ? 
        'If ShowProjekte.Count > 0 And showRangeRight - showRangeLeft > 5 Then
        If ShowProjekte.Count > 0 Then

            If awinSettings.isHryNameFrmActive Then
                Call MsgBox("es kann nur ein Fenster zur Hierarchie- bzw. Namenauswahl geöffnet sein ...")
            ElseIf control.Id = "PTXG1B4" Then
                ' Namen auswählen, Visualisieren
                awinSettings.useHierarchy = False
                With nameFormular
                    .Text = "Plan-Elemente visualisieren"
                    .OKButton.Text = "Anzeigen"
                    .menuOption = PTmenue.visualisieren
                    .statusLabel.Text = ""


                    .rdbBU.Visible = False
                    .pictureBU.Visible = False
                    .rdbTyp.Visible = False
                    .pictureTyp.Visible = False
                    .rdbRoles.Visible = False
                    .pictureRoles.Visible = False
                    .rdbCosts.Visible = False
                    .pictureCosts.Visible = False

                    ' Leistbarkeits-Charts
                    .chkbxOneChart.Checked = False
                    .chkbxOneChart.Visible = False

                    ' Reports 
                    .repVorlagenDropbox.Visible = False
                    .labelPPTVorlage.Visible = False
                    .einstellungen.Visible = False

                    ' Nicht Modal anzeigen
                    .Show()
                    'returnValue = .ShowDialog
                End With

            ElseIf control.Id = "PTXG1B5" Then
                ' Hierarchie auswählen, visualisieren
                awinSettings.useHierarchy = True
                With hryFormular
                    .Text = "Plan-Elemente visualisieren"
                    .OKButton.Text = "Anzeigen"
                    .AbbrButton.Visible = False
                    .AbbrButton.Enabled = False
                    .menuOption = PTmenue.visualisieren
                    .statusLabel.Text = ""

                    .chkbxOneChart.Checked = False
                    .chkbxOneChart.Visible = False

                    ' Reports
                    .repVorlagenDropbox.Visible = False
                    .labelPPTVorlage.Visible = False
                    .einstellungen.Visible = False

                    ' Nicht Modal anzeigen
                    .Show()
                    'returnValue = .ShowDialog
                End With
            ElseIf control.Id = "PTXG1B6" Then
                ' Namen auswählen, Leistbarkeit
                awinSettings.useHierarchy = False
                With nameFormular
                    .Text = "Leistbarkeits-Charts erstellen"
                    .OKButton.Text = "Charts erstellen"
                    .menuOption = PTmenue.leistbarkeitsAnalyse
                    .statusLabel.Text = ""

                    .rdbBU.Visible = False
                    .pictureBU.Visible = False
                    .rdbTyp.Visible = False
                    .pictureTyp.Visible = False

                    .rdbRoles.Visible = True
                    .pictureRoles.Visible = True
                    .rdbCosts.Visible = True
                    .pictureCosts.Visible = True

                    ' Leistbarkeits-Charts
                    .chkbxOneChart.Checked = False
                    .chkbxOneChart.Visible = True

                    ' Reports 
                    .repVorlagenDropbox.Visible = False
                    .labelPPTVorlage.Visible = False

                    ' Nicht Modal anzeigen
                    .Show()
                    'returnValue = .ShowDialog
                End With
            ElseIf control.Id = "PTXG1B7" Then
                ' Hierarchie auswählen, Leistbarkeit
                awinSettings.useHierarchy = True
                With hryFormular
                    .Text = "Leistbarkeits-Charts erstellen"
                    .OKButton.Text = "Charts erstellen"
                    .AbbrButton.Visible = False
                    .AbbrButton.Enabled = False
                    .menuOption = PTmenue.leistbarkeitsAnalyse
                    .statusLabel.Text = ""

                    .chkbxOneChart.Checked = False
                    .chkbxOneChart.Visible = True

                    ' Reports
                    .repVorlagenDropbox.Visible = False
                    .labelPPTVorlage.Visible = False
                    .einstellungen.Visible = False

                    ' Nicht Modal anzeigen
                    .Show()
                    'returnValue = .ShowDialog
                End With

            ElseIf control.Id = "PT1G1M1B1" Then
                ' Namen auswählen, Einzelprojekt Berichte 

                Try
                    awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
                Catch ex As Exception
                    awinSelection = Nothing
                End Try

                If awinSelection Is Nothing Then
                    Call MsgBox("vorher Projekt/e selektieren ...")
                Else

                    appInstance.ScreenUpdating = False

                    With nameFormular

                        .Text = "Projekt-Varianten Report erzeugen"
                        .OKButton.Text = "Bericht erstellen"
                        .menuOption = PTmenue.einzelprojektReport
                        .statusLabel.Text = ""

                        .rdbRoles.Enabled = False
                        .rdbCosts.Enabled = False

                        .rdbBU.Visible = True
                        .pictureBU.Visible = True

                        .rdbTyp.Visible = True
                        .pictureTyp.Visible = True


                        .einstellungen.Visible = True

                        .chkbxOneChart.Checked = False
                        .chkbxOneChart.Visible = False

                        .repVorlagenDropbox.Visible = True
                        .labelPPTVorlage.Visible = True

                        .Show()
                        'returnValue = .ShowDialog
                    End With

                    appInstance.ScreenUpdating = True

                End If

            ElseIf control.Id = "PT1G1M1B2" Then
                ' Hierarchie auswählen, Einzelprojekt Berichte 
                awinSettings.useHierarchy = True
                With hryFormular
                    .Text = "Projekt-Varianten Report erzeugen"
                    .OKButton.Text = "Bericht erstellen"
                    .menuOption = PTmenue.einzelprojektReport
                    .statusLabel.Text = ""

                    .AbbrButton.Visible = False
                    .AbbrButton.Enabled = False

                    .chkbxOneChart.Checked = False
                    .chkbxOneChart.Visible = True


                    ' Reports
                    .repVorlagenDropbox.Visible = True
                    .labelPPTVorlage.Visible = True
                    .einstellungen.Visible = True

                    ' Nicht Modal anzeigen
                    .Show()
                    'returnValue = .ShowDialog
                End With

            ElseIf control.Id = "PT1G1M2B1" Then
                ' Namen Auswahl, Multiprojekt Report
                appInstance.ScreenUpdating = False

                With nameFormular

                    .Text = "Multiprojekt Reports erzeugen"
                    .OKButton.Text = "Bericht erstellen"
                    .menuOption = PTmenue.multiprojektReport
                    .statusLabel.Text = ""

                    .rdbRoles.Enabled = False
                    .rdbCosts.Enabled = False

                    .rdbBU.Visible = True
                    .pictureBU.Visible = True

                    .rdbTyp.Visible = True
                    .pictureTyp.Visible = True


                    .einstellungen.Visible = True

                    .chkbxOneChart.Checked = False
                    .chkbxOneChart.Visible = False

                    .repVorlagenDropbox.Visible = True
                    .labelPPTVorlage.Visible = True

                    .Show()
                    'returnValue = .ShowDialog
                End With

                appInstance.ScreenUpdating = True


            ElseIf control.Id = "PT1G1M2B2" Then
                ' Hierarchie Auswahl, Multiprojekt Report
                appInstance.ScreenUpdating = False

                awinSettings.useHierarchy = True
                With hryFormular

                    .Text = "Multiprojekt Reports erzeugen"
                    .OKButton.Text = "Bericht erstellen"
                    .menuOption = PTmenue.multiprojektReport
                    .statusLabel.Text = ""

                    .AbbrButton.Visible = False
                    .AbbrButton.Enabled = False

                    .chkbxOneChart.Checked = False
                    .chkbxOneChart.Visible = True


                    ' Reports
                    .repVorlagenDropbox.Visible = True
                    .labelPPTVorlage.Visible = True
                    .einstellungen.Visible = True

                    ' Nicht Modal anzeigen
                    .Show()
                    'returnValue = .ShowDialog
                End With

            ElseIf control.Id = "PT4G1M0B1" Then
                ' Auswahl über Namen, Typ II Export
                appInstance.ScreenUpdating = False

                With nameFormular

                    .Text = "Excel Report erzeugen"
                    .OKButton.Text = "Report erstellen"
                    .menuOption = PTmenue.excelExport
                    .statusLabel.Text = ""

                    .rdbRoles.Enabled = False
                    .rdbCosts.Enabled = False

                    .rdbBU.Visible = True
                    .pictureBU.Visible = True

                    .rdbTyp.Visible = True
                    .pictureTyp.Visible = True

                    .einstellungen.Visible = False

                    .chkbxOneChart.Checked = False
                    .chkbxOneChart.Visible = False

                    .repVorlagenDropbox.Visible = False
                    .labelPPTVorlage.Visible = False

                    .Show()
                    'returnValue = .ShowDialog
                End With

                appInstance.ScreenUpdating = True

            ElseIf control.Id = "PT4G1M0B2" Then
                ' Auswahl über Hierarchie, Typ II Export
                appInstance.ScreenUpdating = False

                awinSettings.useHierarchy = True
                With hryFormular

                    .Text = "Excel Report erzeugen"
                    .OKButton.Text = "Report erstellen"
                    .menuOption = PTmenue.excelExport
                    .statusLabel.Text = ""

                    .AbbrButton.Visible = False
                    .AbbrButton.Enabled = False

                    .chkbxOneChart.Checked = False
                    .chkbxOneChart.Visible = False

                    ' Reports
                    .repVorlagenDropbox.Visible = False
                    .labelPPTVorlage.Visible = False
                    .einstellungen.Visible = False

                    ' Nicht Modal anzeigen
                    .Show()
                    'returnValue = .ShowDialog
                End With

            ElseIf control.Id = "Pt6G3M1B1" Then
                ' normale, volle Auswahl des filters ; Namens-Definition

                With nameFormular

                    .Text = "Datenbank Filter definieren"
                    .OKButton.Text = "Speichern"
                    .menuOption = PTmenue.filterdefinieren
                    .statusLabel.Text = ""

                    .rdbRoles.Enabled = True
                    .rdbCosts.Enabled = True

                    .rdbBU.Visible = True
                    .pictureBU.Visible = True

                    .rdbTyp.Visible = True
                    .pictureTyp.Visible = True

                    .einstellungen.Visible = False

                    .chkbxOneChart.Checked = False
                    .chkbxOneChart.Visible = False

                    .repVorlagenDropbox.Visible = False
                    .labelPPTVorlage.Visible = False

                    .Show()
                    'returnValue = .ShowDialog

                End With

            ElseIf control.Id = "Pt6G3M1B2" Then
                ' Auswahl über Hierarchie, Datenbank Filter nur Filter für Vorgänge und Meilensteine 
                awinSettings.useHierarchy = True
                With hryFormular

                    .Text = "Datenbank Filter definieren"
                    .OKButton.Text = "Speichern"
                    .menuOption = PTmenue.filterdefinieren
                    .statusLabel.Text = ""

                    .AbbrButton.Visible = False
                    .AbbrButton.Enabled = False

                    .chkbxOneChart.Checked = False
                    .chkbxOneChart.Visible = False

                    ' Reports
                    .repVorlagenDropbox.Visible = False
                    .labelPPTVorlage.Visible = False
                    .einstellungen.Visible = False

                    ' Nicht Modal anzeigen
                    .Show()
                    'returnValue = .ShowDialog
                End With

            Else


            End If

        Else
            Call MsgBox("Es sind keine Projekte sichtbar!  ")
        End If
        



        appInstance.EnableEvents = True
        enableOnUpdate = True


    End Sub


    Sub AnalyseLeistbarkeit001(ByVal control As IRibbonControl)

        Dim namensFormular As New frmNameSelection
        Dim hierarchieFormular As New frmHierarchySelection
        Dim returnValue As DialogResult

        Call projektTafelInit()

        enableOnUpdate = False
        appInstance.EnableEvents = False

        ' gibt es überhaupt Objekte, zu denen man was anzeigen kann ? 
        If ShowProjekte.Count > 0 And showRangeRight - showRangeLeft > 5 Then

            If control.Id = "PTXG1B6" Then
                ' Auswahl über Namen

                With namensFormular
                    .Text = "Leistbarkeit analysieren"

                    .rdbBU.Visible = False
                    .pictureBU.Visible = False

                    .rdbTyp.Visible = False
                    .pictureTyp.Visible = False

                    .rdbRoles.Visible = True
                    .pictureRoles.Visible = True

                    .rdbCosts.Visible = True
                    .pictureCosts.Visible = True

                    '.chkbxShowObjects = False

                    .chkbxOneChart.Checked = False
                    .chkbxOneChart.Visible = True

                    '.chkbxCreateCharts = True


                    .repVorlagenDropbox.Visible = False
                    .labelPPTVorlage.Visible = False

                    '.showModePortfolio = True

                    .menuOption = PTmenue.leistbarkeitsAnalyse
                    .OKButton.Text = "Charts erstellen"

                    '.Show()
                    returnValue = .ShowDialog
                End With


            Else
                ' Auswahl über Hierarchie
                ' Hierarchie
                awinSettings.useHierarchy = True
                With hierarchieFormular
                    .Text = "Leistbarkeit analysieren"

                    .chkbxOneChart.Checked = False
                    .chkbxOneChart.Visible = True

                    '.chkbxCreateCharts = False


                    .repVorlagenDropbox.Visible = False
                    .labelPPTVorlage.Visible = False

                    '.showModePortfolio = True
                    .menuOption = PTmenue.leistbarkeitsAnalyse

                    .OKButton.Text = "Charts erstellen"

                    '.Show()
                    returnValue = .ShowDialog
                End With

            End If

        ElseIf ShowProjekte.Count = 0 Then

            Call MsgBox("Es sind keine Projekte geladen!  ")

        ElseIf showRangeRight - showRangeLeft <= 5 Then

            Call MsgBox("bitte zuerst einen Zeitraum markieren! ")

        End If



        appInstance.EnableEvents = True
        enableOnUpdate = True



    End Sub


    ''' <summary>
    ''' Projekt ins Show zurückholen 
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G1Show(control As IRibbonControl)

        Dim getBackToShow As New frmGetProjectbackFromNoshow

        Dim returnValue As DialogResult

        Call projektTafelInit()

        enableOnUpdate = False
        appInstance.ScreenUpdating = False

        If AlleProjekte.Count > 0 And ShowProjekte.Count <> AlleProjekte.Count Then

            returnValue = getBackToShow.ShowDialog
        Else
            If AlleProjekte.Count = 0 Then
                Call MsgBox("Es sind keine Projekte geladen!  ")
            Else
                Call MsgBox("Es gibt keine Projekte in der Warteschlange !")
            End If
        End If



        appInstance.ScreenUpdating = True
        enableOnUpdate = True
    End Sub
    ''' <summary>
    ''' Änderungen akzeptieren 
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G1Accept(control As IRibbonControl)

        Dim singleShp As Excel.Shape


        Dim awinSelection As Excel.ShapeRange

        Call projektTafelInit()

        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False

        enableOnUpdate = False

        Try
            'awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            ' jetzt die Aktion durchführen ...

            For Each singleShp In awinSelection

                Dim shapeArt As Integer
                shapeArt = kindOfShape(singleShp)

                With singleShp
                    If isProjectType(shapeArt) Then
                        Call awinBeauftragung(pname:=.Name, type:=0)
                    End If
                End With
            Next

        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If

        enableOnUpdate = True
        appInstance.EnableEvents = formerEE

    End Sub



    ''' <summary>
    ''' Projekt beauftragen
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G1Beauftragen(control As IRibbonControl)

        Dim singleShp As Excel.Shape
        Dim awinSelection As Excel.ShapeRange

        Call projektTafelInit()

        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False

        enableOnUpdate = False

        Try
            'awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            ' jetzt die Aktion durchführen ...

            For Each singleShp In awinSelection

                Dim shapeArt As Integer
                shapeArt = kindOfShape(singleShp)

                With singleShp
                    If isProjectType(shapeArt) Then
                        Call awinBeauftragung(pname:=.Name, type:=1)
                    End If
                End With
            Next

        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If

        enableOnUpdate = True
        appInstance.EnableEvents = formerEE

    End Sub

    ''' <summary>
    ''' Beauftragung zurücknehmen 
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2GXBeauftragen(control As IRibbonControl)

        Dim singleShp As Excel.Shape
        Dim awinSelection As Excel.ShapeRange

        Call projektTafelInit()

        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False

        enableOnUpdate = False

        Try
            'awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            ' jetzt die Aktion durchführen ...

            For Each singleShp In awinSelection
                Dim shapeArt As Integer
                shapeArt = kindOfShape(singleShp)

                With singleShp
                    If isProjectType(shapeArt) Then
                        Call awinCancelBeauftragung(pname:=.Name)
                    End If
                End With
            Next

        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If

        enableOnUpdate = True
        appInstance.EnableEvents = formerEE

    End Sub



    ''' <summary>
    ''' Projekt löschen
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G1Loeschen(control As IRibbonControl)

        Dim bestaetigeLoeschen As New frmconfirmDeletePrj
        Dim singleShp As Excel.Shape
        Dim awinSelection As Excel.ShapeRange
        Dim returnValue As DialogResult

        Call projektTafelInit()

        appInstance.EnableEvents = False
        enableOnUpdate = False

        Try
            'awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            bestaetigeLoeschen.botschaft = "Bitte bestätigen Sie das Löschen" & vbLf & _
                                            "Vorsicht: alle Varianten werden mitgelöscht ..."
            returnValue = bestaetigeLoeschen.ShowDialog

            If returnValue = DialogResult.Cancel Then

                appInstance.EnableEvents = True
                enableOnUpdate = True
                Exit Sub

            End If



            ' jetzt die Aktion durchführen ...


            For Each singleShp In awinSelection


                Dim shapeArt As Integer
                shapeArt = kindOfShape(singleShp)

                With singleShp
                    If isProjectType(shapeArt) Then

                        Try
                            Call awinDeleteProjectInSession(pName:=.Name)

                        Catch ex As Exception
                            Exit For
                        End Try

                    End If
                End With


            Next

            ' ein oder mehrere Projekte wurden gelöscht  - typus = 3
            Call awinNeuZeichnenDiagramme(3)

        Else

            Dim deletedProj As Integer = 0

            If AlleProjekte.Count = 0 Then
                Call MsgBox("es sind keine Projekte geladen !")
            Else

                'Dim deleteProjects As New frmDeleteProjects
                Dim deleteProjects As New frmProjPortfolioAdmin
                Try

                    With deleteProjects
                        .Text = "Projekte, Varianten aus der Session löschen"
                        .aKtionskennung = PTTvActions.delFromSession
                        .OKButton.Text = "Löschen"
                        .portfolioName.Visible = False
                        .Label1.Visible = False
                    End With

                    returnValue = deleteProjects.ShowDialog

                    If returnValue = DialogResult.OK Then

                        'Call MsgBox("ok, aus Session gelöscht  !")

                    Else
                        ' returnValue = DialogResult.Cancel

                    End If

                Catch ex As Exception

                    Call MsgBox(ex.Message)
                End Try

            End If



        End If

        Call awinDeSelect()

        enableOnUpdate = True
        appInstance.EnableEvents = True

    End Sub



    ''' <summary>
    ''' EinzelProjekt Report mit selektierter Vorlage erstellen
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G1Doku(control As IRibbonControl)

        Dim awinSelection As Excel.ShapeRange
        Dim returnValue As DialogResult
        Dim getReportVorlage As New frmSelectPPTTempl

        Call projektTafelInit()

        Try
            'awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If awinSelection Is Nothing Then
            Call MsgBox("vorher Projekt/e selektieren ...")
        Else
            enableOnUpdate = False
            appInstance.ScreenUpdating = False
            appInstance.EnableEvents = False

            getReportVorlage.calledfrom = "Projekt"


            ' sichern der awinSettings.mpp... Einstellungen
            ' Settings für Multiprojekt-Sichten
            ' '' '' ''Dim sav_mppShowAllIfOne As Boolean = awinSettings.mppShowAllIfOne
            ' '' '' ''Dim sav_mppShowMsDate As Boolean = awinSettings.mppShowMsDate
            ' '' '' ''Dim sav_mppShowMsName As Boolean = awinSettings.mppShowMsName
            ' '' '' ''Dim sav_mppShowPhDate As Boolean = awinSettings.mppShowPhDate
            ' '' '' ''Dim sav_mppShowPhName As Boolean = awinSettings.mppShowPhName
            ' '' '' ''Dim sav_mppShowAmpel As Boolean = awinSettings.mppShowAmpel
            ' '' '' ''Dim sav_mppShowProjectLine As Boolean = awinSettings.mppShowProjectLine
            ' '' '' ''Dim sav_mppVertikalesRaster As Boolean = awinSettings.mppVertikalesRaster
            ' '' '' ''Dim sav_mppShowLegend As Boolean = awinSettings.mppShowLegend
            ' '' '' ''Dim sav_mppFullyContained As Boolean = awinSettings.mppFullyContained
            ' '' '' ''Dim sav_mppSortiertDauer As Boolean = awinSettings.mppSortiertDauer
            ' '' '' ''Dim sav_mppOnePage As Boolean = awinSettings.mppOnePage
            Dim sav_mppExtendedMode As Boolean = awinSettings.mppExtendedMode
            awinSettings.mppExtendedMode = True
            ' Settings für Einzelprojekt-Reports
            awinSettings.eppExtendedMode = True


            ' Formular zum Auswählen der Report-Vorlage wird aufgerufen

            returnValue = getReportVorlage.ShowDialog

            awinSettings.eppExtendedMode = False

            ' wieder setzen der awinSettings.mpp... einstellungen
            ' '' ''awinSettings.mppShowAllIfOne = sav_mppShowAllIfOne
            ' '' ''awinSettings.mppShowMsDate = sav_mppShowMsDate
            ' '' ''awinSettings.mppShowMsName = sav_mppShowMsName
            ' '' ''awinSettings.mppShowPhDate = sav_mppShowPhDate
            ' '' ''awinSettings.mppShowPhName = sav_mppShowPhName
            ' '' ''awinSettings.mppShowAmpel = sav_mppShowAmpel
            ' '' ''awinSettings.mppShowProjectLine = sav_mppShowProjectLine
            ' '' ''awinSettings.mppVertikalesRaster = sav_mppVertikalesRaster
            ' '' ''awinSettings.mppShowLegend = sav_mppShowLegend
            ' '' ''awinSettings.mppFullyContained = sav_mppFullyContained
            ' '' ''awinSettings.mppSortiertDauer = sav_mppSortiertDauer
            ' '' ''awinSettings.mppOnePage = sav_mppOnePage

            awinSettings.mppExtendedMode = sav_mppExtendedMode

            appInstance.EnableEvents = True
            appInstance.ScreenUpdating = True
            enableOnUpdate = True
        End If

    End Sub

    Public Sub Tom2G4B1InventurImport(control As IRibbonControl)


        Dim projektInventurFile As String = requirementsOrdner & "Projekt-Inventur.xlsx"
        Dim dateiName As String
        Dim myCollection As New Collection
        Dim importDate As Date = Date.Now

        Call projektTafelInit()

        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False
        enableOnUpdate = False

        dateiName = awinPath & projektInventurFile

        Try

            appInstance.Workbooks.Open(dateiName)
            ' alle Import Projekte erstmal löschen
            ImportProjekte.Clear()
            Call awinImportProjektInventur(myCollection)
            'Call bmwImportProjektInventur(myCollection)

            appInstance.ActiveWorkbook.Save()
            appInstance.ActiveWorkbook.Close(SaveChanges:=False)


        Catch ex As Exception
            Call MsgBox("Fehler bei Import " & vbLf & dateiName & vbLf & ex.Message)
            Exit Sub
        End Try

        Try
            Call importProjekteEintragen(myCollection, importDate, ProjektStatus(0))
        Catch ex As Exception
            Call MsgBox("Fehler bei Import : " & vbLf & ex.Message)
        End Try


        enableOnUpdate = True
        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True

    End Sub

    Public Sub Tom2G4B1RPLANImport(control As IRibbonControl)


        'Dim projektInventurFile As String = requirementsOrdner & "RPLAN Projekte.xlsx"
        Dim dateiName As String
        Dim myCollection As New Collection
        Dim importDate As Date = Date.Now
        Dim returnValue As DialogResult
        Dim getRPLANImport As New frmSelectRPlanImport

        Call projektTafelInit()

        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False
        enableOnUpdate = False

        'dateiName = awinPath & projektInventurFile


        returnValue = getRPLANImport.ShowDialog

        If returnValue = DialogResult.OK Then
            dateiName = getRPLANImport.RPLANdateiName

            Try
                appInstance.Workbooks.Open(dateiName)

                ' alle Import Projekte erstmal löschen
                ImportProjekte.Clear()
                'Call bmwImportProjektInventur(myCollection)
                Call bmwImportProjekteITO15(myCollection, False)
                appInstance.ActiveWorkbook.Close(SaveChanges:=True)
                Call importProjekteEintragen(myCollection, importDate, ProjektStatus(1))

            Catch ex As Exception
                appInstance.ActiveWorkbook.Close(SaveChanges:=False)
                Call MsgBox("Fehler bei Import " & vbLf & dateiName & vbLf & ex.Message)
            End Try
        Else
            Call MsgBox(" Import RPLAN-Projekte wurde abgebrochen")
        End If



        enableOnUpdate = True
        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True

    End Sub

    Public Sub Tom2G4M1Import(control As IRibbonControl)

        Dim request As New Request(awinSettings.databaseName, dbUsername, dbPasswort)
        Dim hproj As New clsProjekt
        Dim cproj As New clsProjekt
        Dim vglName As String = " "
        Dim outputString As String = ""
        Dim dirName As String
        Dim dateiName As String
        Dim pname As String
        Dim importDate As Date = Date.Now
        'Dim importDate As Date = "31.10.2013"
        Dim listOfVorlagen As Collections.ObjectModel.ReadOnlyCollection(Of String)
        Dim projektInventurFile As String = "ProjektInventur.xlsm"

        Call projektTafelInit()

        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False
        enableOnUpdate = False

        Dim myCollection As New Collection




        dirName = awinPath & projektFilesOrdner
        listOfVorlagen = My.Computer.FileSystem.GetFiles(dirName, FileIO.SearchOption.SearchTopLevelOnly, "*.xlsx")

        ' alle Import Projekte erstmal löschen
        ImportProjekte.Clear()


        ' jetzt müssen die Projekte ausgelesen werden, die in dateiListe stehen 
        Dim i As Integer
        For i = 1 To listOfVorlagen.Count
            dateiName = listOfVorlagen.Item(i - 1)

            If dateiName = projektInventurFile Then

                ' nichts machen 

            Else
                Dim skip As Boolean = False


                Try
                    appInstance.Workbooks.Open(dateiName)
                Catch ex1 As Exception
                    'Call MsgBox("Fehler bei Öffnen der Datei " & dateiName)
                    skip = True
                End Try

                If Not skip Then
                    pname = ""
                    hproj = New clsProjekt
                    Try
                        Call awinImportProject(hproj, Nothing, False, importDate)

                        Try
                            Dim keyStr As String = calcProjektKey(hproj)
                            ImportProjekte.Add(calcProjektKey(hproj), hproj)
                            myCollection.Add(calcProjektKey(hproj))
                        Catch ex2 As Exception
                            Call MsgBox("Projekt kann nicht zweimal importiert werden ...")
                        End Try

                        appInstance.ActiveWorkbook.Close(SaveChanges:=False)

                    Catch ex1 As Exception
                        appInstance.ActiveWorkbook.Close(SaveChanges:=False)
                        Call MsgBox(ex1.Message)
                        Call MsgBox("Fehler bei Import von Projekt " & hproj.name)
                    End Try



                End If



            End If


        Next i



        Try
            Call importProjekteEintragen(myCollection, importDate, ProjektStatus(1))
        Catch ex As Exception
            Call MsgBox("Fehler bei Import : " & vbLf & ex.Message)
        End Try




        enableOnUpdate = True
        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True




    End Sub

    ''' <summary>
    ''' exportiert selektierte/alle Files in eine Excel Datei, die genauso aufgebaut ist , wie die BMW Import Datei  
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Public Sub bmwExcelExport(control As IRibbonControl)

        Dim singleShp As Excel.Shape
        Dim hproj As clsProjekt
        Dim outputString As String = ""
        Dim fileListe As New SortedList(Of String, String)
        Dim exportFileName As String = "Export_" & Date.Now.ToString.Replace(":", ".") & ".xlsx"
        Dim ok As Boolean

        Dim awinSelection As Excel.ShapeRange

        Call projektTafelInit()

        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False

        enableOnUpdate = False

        Try
            'awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            ' jetzt die Aktion durchführen ...

            For Each singleShp In awinSelection

                ' hier muss jetzt die todo Liste aufgebaut werden 

                Dim shapeArt As Integer
                shapeArt = kindOfShape(singleShp)

                With singleShp
                    If isProjectType(shapeArt) Then

                        Try

                            hproj = ShowProjekte.getProject(singleShp.Name)
                            fileListe.Add(hproj.name, hproj.name)

                        Catch ex As Exception

                            Call MsgBox(singleShp.Name & ": Fehler bei Aufbau todo Liste für Export ...")

                        End Try

                    End If
                End With

            Next

        Else
            For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste
                fileListe.Add(kvp.Key, kvp.Key)
            Next
        End If

        ' hier muss jetzt das File Projekt Detail aufgemacht werden ...
        Try
            appInstance.Workbooks.Open(awinPath & requirementsOrdner & excelExportVorlage)
            ok = True
        Catch ex As Exception
            ok = False
        End Try

        If ok Then
            Dim zeile As Integer = 2
            For Each kvp As KeyValuePair(Of String, String) In fileListe

                Try
                    hproj = ShowProjekte.getProject(kvp.Key)

                    ' jetzt wird dieses Projekt exportiert ... 
                    Try
                        Call bmwExportProject(hproj, zeile)
                        outputString = outputString & hproj.name & " erfolgreich .." & vbLf
                    Catch ex As Exception
                        outputString = outputString & hproj.name & " nicht erfolgreich .." & vbLf & _
                                        ex.Message & vbLf & vbLf
                    End Try



                Catch ex As Exception

                    Call MsgBox(ex.Message)

                End Try

            Next

            Try
                ' Schließen der Export Datei unter neuem Namen, original Zustand bleibt erhalten
                appInstance.ActiveWorkbook.Close(SaveChanges:=True, Filename:=awinPath & exportFilesOrdner & "\" & _
                                                 exportFileName)
                Call MsgBox(outputString & "exportiert !")
            Catch ex As Exception

                Call MsgBox("Fehler beim Speichern der Export Datei")

            End Try

        End If


        Call awinDeSelect()
        enableOnUpdate = True
        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True



    End Sub

    ''' <summary>
    ''' exportiert selektierte / alle Files in eine Excel Datei; 
    ''' verwendet dabei die Vorlage in Requirements bmwFC52Vorlage.xlsx
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Public Sub bmwFC52Export(control As IRibbonControl)

        Call projektTafelInit()

        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False
        enableOnUpdate = False

        Call awinWriteFC52()

        Call awinDeSelect()
        enableOnUpdate = True
        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True

    End Sub


    Public Sub Tom2G4M1Export(control As IRibbonControl)


        Dim singleShp As Excel.Shape
        Dim hproj As clsProjekt
        Dim outputString As String = ""


        Dim awinSelection As Excel.ShapeRange

        Call projektTafelInit()

        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False

        enableOnUpdate = False

        Try
            'awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            ' jetzt die Aktion durchführen ...

            For Each singleShp In awinSelection
                Try
                    ' hier muss jetzt das File Projekt Detail aufgemacht werden ...
                    appInstance.Workbooks.Open(awinPath & projektAustausch)

                    Dim shapeArt As Integer
                    shapeArt = kindOfShape(singleShp)

                    With singleShp
                        If isProjectType(shapeArt) Then

                            Try
                                hproj = ShowProjekte.getProject(singleShp.Name)

                                ' jetzt wird dieses Projekt exportiert ... 
                                Try
                                    Call awinExportProject(hproj)
                                    outputString = outputString & hproj.getShapeText & " erfolgreich .." & vbLf
                                Catch ex As Exception
                                    outputString = outputString & hproj.getShapeText & " nicht erfolgreich .." & vbLf & _
                                                    ex.Message & vbLf & vbLf
                                End Try



                            Catch ex As Exception
                                Call MsgBox(singleShp.Name & " nicht gefunden ...")

                            End Try

                        End If
                    End With
                    Try
                        ' Schließen der Datei ProjektSteckbrief ohne abspeichern der Änderungen, original Zustand bleibt erhalten
                        appInstance.ActiveWorkbook.Close(SaveChanges:=False, Filename:=awinPath & projektAustausch)
                    Catch ex As Exception

                        Call MsgBox("Fehler beim Schließen der Projektaustausch Vorlage")

                    End Try
                Catch ex As Exception

                    Call MsgBox("Fehler beim Öffnen der Projektaustausch Vorlage")

                End Try


            Next

            Call MsgBox(outputString & "exportiert !")

        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If


        Call awinDeSelect()
        enableOnUpdate = True
        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True




    End Sub

    ''' <summary>
    ''' erstellt die Summary Zuordnungs-Datei 
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G4M2B1ZuordnungRP(control As IRibbonControl)


        Dim fileName As String
        Dim zeile As Integer = 2
        Dim ok As Boolean

        Call projektTafelInit()

        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False
        enableOnUpdate = False


        fileName = "Vorlage Zuordnung.xlsx"

        ' öffnen der Excel Datei 
        Try
            appInstance.Workbooks.Open(awinPath & projektRessOrdner & "\" & fileName)
        Catch ex As Exception
            Call MsgBox("File " & fileName & " nicht gefunden ... Abbruch")
            appInstance.EnableEvents = True
            appInstance.ScreenUpdating = True
            enableOnUpdate = True
            Exit Sub
        End Try




        Call awinExportRessZuordnung(0, " ")


        Try

            appInstance.ActiveWorkbook.SaveAs(awinPath & projektRessOrdner & "\Summary.xlsx", _
                                      ConflictResolution:=XlSaveConflictResolution.xlLocalSessionChanges)
            ok = True
            appInstance.ActiveWorkbook.Close()

        Catch ex As Exception
            ok = False
            appInstance.ActiveWorkbook.Close()
        End Try


        If ok Then
            Call MsgBox("ok, Datei erstellt ...")
        Else
            Call MsgBox("Fehler bei Save as ..\summary.xlsx")
        End If


        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True
        enableOnUpdate = True


    End Sub

    ''' <summary>
    ''' erstellt die Zuordnungs-Datei Ressourcen -> Projekt
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G4M2B2ZuordnungRP(control As IRibbonControl)

        Dim initialeVorlageName As String, kapaFileName As String
        Dim zeile As Integer = 2
        Dim anzRollen As Integer
        Dim i As Integer
        Dim initMessage As String = "bitte die Kapazitäten eintragen zu folgenden Rollen" & vbLf
        Dim infoMessage As String = initMessage
        Dim zuordnungsOrdner As String = projektRessOrdner & "\" & "Projekt Zuordnungen"

        Call projektTafelInit()

        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False
        enableOnUpdate = False



        ' für jede Ressource eine eigene Datei machen
        anzRollen = RoleDefinitions.Count

        Dim ok As Boolean = True
        Dim roleName As String

        For i = 1 To anzRollen

            roleName = RoleDefinitions.getRoledef(i).name.Trim
            kapaFileName = roleName & " Kapazität.xlsx"

            ' öffnen der Excel Datei 
            Try

                appInstance.Workbooks.Open(awinPath & projektRessOrdner & "\" & kapaFileName)
                ok = True

            Catch ex As Exception

                initialeVorlageName = "template Kapazität.xlsx"
                ok = False

                Try
                    appInstance.Workbooks.Open(awinPath & projektRessOrdner & "\" & initialeVorlageName)
                    Try
                        appInstance.ActiveWorkbook.SaveAs(awinPath & projektRessOrdner & "\" & kapaFileName, _
                                      ConflictResolution:=XlSaveConflictResolution.xlLocalSessionChanges)

                        infoMessage = infoMessage & kapaFileName & vbLf
                    Catch ex2 As Exception

                    End Try



                Catch ex1 As Exception
                    Call MsgBox("File " & initialeVorlageName & " nicht gefunden ... Abbruch" & vbLf & vbLf & _
                                "dieses File muss im Ordner " & awinPath & projektRessOrdner & "abgelegt werden")
                    appInstance.EnableEvents = True
                    appInstance.ScreenUpdating = True
                    enableOnUpdate = True
                    Exit Sub
                End Try

            End Try


            If ok Then

                Dim curFilename As String = roleName & " Projekt-Zuordnung" & " " & Date.Now.ToString("MMM yy")


                Try
                    Call awinExportRessZuordnung(1, roleName)
                    'appInstance.ActiveWorkbook.Save()

                    appInstance.ActiveWorkbook.SaveAs(Filename:=awinPath & zuordnungsOrdner & "\" & curFilename, _
                                                      ConflictResolution:=XlSaveConflictResolution.xlLocalSessionChanges)


                Catch ex As Exception

                    Call MsgBox("Fehler bei Zuordnung " & roleName)
                End Try

            End If


            appInstance.ActiveWorkbook.Close(SaveChanges:=False)



        Next

        If infoMessage.Length > initMessage.Length Then
            ' in diesem Fall wurden  nur die Kapazität-Zuordnungs-Files erstellt 
            infoMessage = infoMessage & vbLf & vbLf & "es wurden noch keine Zuordnungs-Dateien erstellt!"
            Call MsgBox(infoMessage)
        Else
            Call MsgBox("ok, Dateien erstellt ...")
        End If



        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True
        enableOnUpdate = True


    End Sub

    Sub PTDemoModusHistory(control As IRibbonControl, ByRef pressed As Boolean)

        demoModusHistory = Not demoModusHistory
        pressed = demoModusHistory

    End Sub

    Sub awinSetModusHistory(control As IRibbonControl, ByRef pressed As Boolean)

        Dim demoModusDate As New frmdemoModusDate
        Dim returnValue As DialogResult

        Call projektTafelInit()

        If pressed Then

            demoModusHistory = True

            returnValue = demoModusDate.ShowDialog

            If returnValue = DialogResult.OK Then

                If demoModusHistory Then
                    Call MsgBox("Demo Modus History: Ein" & vbLf & "neues Datum: " & historicDate)
                Else
                    Call MsgBox("Demo Modus History: Aus")
                End If

            Else
                If demoModusHistory Then
                    Call MsgBox("Demo Modus History: Ein" & vbLf & "altes Datum: " & historicDate)
                Else
                    Call MsgBox("Demo Modus History: Aus")
                End If

            End If
        Else
            demoModusHistory = False
            'Call MsgBox("Demo Modus History: Aus")
        End If


    End Sub


    Public Sub PT5phasenZeichnenInit(control As IRibbonControl, ByRef pressed As Boolean)

        pressed = awinSettings.drawphases

    End Sub

    Public Sub PT5phasenZeichnen(control As IRibbonControl, ByRef pressed As Boolean)

        Call projektTafelInit()

        If pressed Then
            ' jetzt werden die Projekt-Symbole inkl Phasen Darstellung gezeichnet
            awinSettings.drawphases = True
            Call awinClearPlanTafel()
            Call awinZeichnePlanTafel(False)
        Else
            ' jetzt werden die Projekt-Symbole ohne Phasen Darstellung gezeichnet 
            awinSettings.drawphases = False
            'Call awinLoadConstellation("Last")
            Call awinClearPlanTafel()
            Call awinZeichnePlanTafel(False)
        End If

    End Sub

    Sub PTShowSelectedObjects(control As IRibbonControl, ByRef pressed As Boolean)

        awinSettings.showValuesOfSelected = Not awinSettings.showValuesOfSelected
        pressed = awinSettings.showValuesOfSelected

    End Sub

    Sub awinSetShowSelObj(control As IRibbonControl, ByRef pressed As Boolean)

        If pressed Then
            awinSettings.showValuesOfSelected = True
        Else
            awinSettings.showValuesOfSelected = False
        End If

    End Sub


    Sub PTPropAnpassen(control As IRibbonControl, ByRef pressed As Boolean)

        awinSettings.propAnpassRess = Not awinSettings.propAnpassRess
        pressed = awinSettings.propAnpassRess

    End Sub

    Sub awinSetPropAnpass(control As IRibbonControl, ByRef pressed As Boolean)

        If pressed Then
            awinSettings.propAnpassRess = True
        Else
            awinSettings.propAnpassRess = False
        End If

    End Sub

    Sub PTPhaseAnteilig(control As IRibbonControl, ByRef pressed As Boolean)

        awinSettings.phasesProzentual = Not awinSettings.phasesProzentual
        pressed = awinSettings.phasesProzentual

    End Sub

    Sub awinSetPhaseAnteilig(control As IRibbonControl, ByRef pressed As Boolean)

        If pressed Then
            awinSettings.phasesProzentual = True
        Else
            awinSettings.phasesProzentual = False
        End If

    End Sub


    'Public Sub PT6StriktPressed(control As IRibbonControl, ByRef pressed As Boolean)

    '    pressed = awinSettings.mppStrict

    'End Sub

    'Public Sub PT6SetStrict(control As IRibbonControl, ByRef pressed As Boolean)

    '    If pressed Then
    '        awinSettings.mppStrict = True
    '    Else
    '        awinSettings.mppStrict = False
    '    End If

    'End Sub

    'Public Sub PT6fullyContainedPressed(control As IRibbonControl, ByRef pressed As Boolean)

    '    pressed = awinSettings.mppFullyContained

    'End Sub

    'Public Sub PT6SetfullyContained(control As IRibbonControl, ByRef pressed As Boolean)

    '    If pressed Then
    '        awinSettings.mppFullyContained = True
    '    Else
    '        awinSettings.mppFullyContained = False
    '    End If

    'End Sub


    'Public Sub PT6DateTextPressed(control As IRibbonControl, ByRef pressed As Boolean)
    '    pressed = awinSettings.mppShowMsDate
    'End Sub


    'Public Sub PT6SetShowDate(Control As IRibbonControl, ByRef pressed As Boolean)

    '    If pressed Then
    '        awinSettings.mppShowMsDate = True
    '    Else
    '        awinSettings.mppShowMsDate = False
    '    End If

    'End Sub


    'Public Sub PT6NameTextPressed(control As IRibbonControl, ByRef pressed As Boolean)
    '    pressed = awinSettings.mppShowMsName
    'End Sub


    'Public Sub PT6SetShowName(Control As IRibbonControl, ByRef pressed As Boolean)

    '    If pressed Then
    '        awinSettings.mppShowMsName = True
    '    Else
    '        awinSettings.mppShowMsName = False
    '    End If

    'End Sub

    'Public Sub PT6ProjectLinePressed(control As IRibbonControl, ByRef pressed As Boolean)
    '    pressed = awinSettings.mppShowProjectLine
    'End Sub


    'Public Sub PT6SetShowProjectLine(Control As IRibbonControl, ByRef pressed As Boolean)

    '    If pressed Then
    '        awinSettings.mppShowProjectLine = True
    '    Else
    '        awinSettings.mppShowProjectLine = False
    '    End If

    'End Sub

    Public Sub PT6AmpelnPressed(control As IRibbonControl, ByRef pressed As Boolean)
        pressed = awinSettings.mppShowAmpel
    End Sub


    Public Sub PT6SetShowAmpeln(Control As IRibbonControl, ByRef pressed As Boolean)

        If pressed Then
            awinSettings.mppShowAmpel = True
        Else
            awinSettings.mppShowAmpel = False
        End If

    End Sub

    'Public Sub PT6RasterPressed(control As IRibbonControl, ByRef pressed As Boolean)
    '    pressed = awinSettings.mppVertikalesRaster
    'End Sub


    'Public Sub PT6SetRaster(Control As IRibbonControl, ByRef pressed As Boolean)

    '    If pressed Then
    '        awinSettings.mppVertikalesRaster = True
    '    Else
    '        awinSettings.mppVertikalesRaster = False
    '    End If

    'End Sub


    ''' <summary>
    ''' lädt die gewählten Projekte und gewählten Varianten in die Session
    ''' </summary>
    ''' <param name="Control"></param>
    ''' <remarks></remarks>
    Public Sub PT5DatenbankLoadProjekte(Control As IRibbonControl)

        Dim deletedProj As Integer = 0
        Dim returnValue As DialogResult

        'Dim deleteProjects As New frmDeleteProjects
        Dim loadProjectsForm As New frmProjPortfolioAdmin

        Try

            With loadProjectsForm
                .Text = "Projekte und Varianten in die Session laden "
                .aKtionskennung = PTTvActions.loadPV
                .OKButton.Text = "Laden"
                .portfolioName.Visible = False
                .Label1.Visible = False
            End With

            returnValue = loadProjectsForm.ShowDialog

            If returnValue = DialogResult.OK Then
                'deletedProj = RemoveSelectedProjectsfromDB(deleteProjects.selectedItems)    ' es werden die selektierten Projekte in der DB gespeichert, die Anzahl gespeicherter Projekte sind das Ergebnis

            Else
                ' returnValue = DialogResult.Cancel

            End If

        Catch ex As Exception

            Call MsgBox(ex.Message)
        End Try




    End Sub


    Public Sub PT5loadprojectsInit(control As IRibbonControl, ByRef pressed As Boolean)

        pressed = awinSettings.applyFilter


    End Sub

    Public Sub PT5loadProjectsOnChange(control As IRibbonControl, ByRef pressed As Boolean)

        Call projektTafelInit()

        If pressed Then
            ' jetzt sollen die Projekte gemäß Time Span geladen werden - auch bei Veränderung des TimeSpan 
            awinSettings.applyFilter = True
            ' noch zu tun 
            ' Call awinloadProjectsFromDB()
        Else

            ' jetzt werden bei TimeSpan Änderung keine Projekte nachgeladen 
            awinSettings.applyFilter = False


        End If


    End Sub
    ''' <summary>
    ''' Charakteristik Phasen
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G2M1B1Phasen(control As IRibbonControl)

        Dim singleShp As Excel.Shape
        Dim top As Double, left As Double, width As Double, height As Double
        Dim hproj As clsProjekt
        Dim scale As Double
        'Dim SID As String

        Call projektTafelInit()

        enableOnUpdate = False

        Dim awinSelection As Excel.ShapeRange

        Try
            'awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            If awinSelection.Count = 1 Then
                ' jetzt die Aktion durchführen ...
                singleShp = awinSelection.Item(1)

                Try
                    hproj = ShowProjekte.getProject(singleShp.Name)
                Catch ex As Exception
                    Call MsgBox("Projekt nicht gefunden ..." & singleShp.Name)
                    Exit Sub
                End Try

                top = singleShp.Top + boxHeight + 2
                left = singleShp.Left - 5
                If left <= 0 Then
                    left = 5
                End If

                height = 380
                width = hproj.dauerInDays / 365 * 12 * boxWidth + 7
                scale = hproj.dauerInDays


                Dim repObj As Excel.ChartObject
                appInstance.EnableEvents = False
                appInstance.ScreenUpdating = False

                repObj = Nothing
                Dim noColorCollection As New Collection
                Call createPhasesBalken(noColorCollection, hproj, repObj, scale, top, left, height, width, PThis.current)

                appInstance.EnableEvents = True
                appInstance.ScreenUpdating = True
            Else
                Call MsgBox("bitte nur ein Projekt selektieren")
                'For Each singleShp In awinSelection
                '    With singleShp
                '        If .AutoShapeType = MsoAutoShapeType.msoShapeRoundedRectangle Then
                '            nrSelPshp = nrSelPshp + 1
                '            SID = .ID.ToString
                '        End If
                '    End With
                'Next
            End If
        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If

        enableOnUpdate = True

    End Sub

    ''' <summary>
    ''' für BMW Akquise erzeugt 
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G2M1B1Phasen2(control As IRibbonControl)

        Dim singleShp As Excel.Shape
        Dim hproj As clsProjekt
        Dim pname As String

        Call projektTafelInit()

        enableOnUpdate = False

        Dim awinSelection As Excel.ShapeRange

        Try
            'awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            If awinSelection.Count = 1 Then
                ' jetzt die Aktion durchführen ...
                singleShp = awinSelection.Item(1)
                pname = singleShp.Name

                Try
                    hproj = ShowProjekte.getProject(pname)
                Catch ex As Exception
                    Call MsgBox("Projekt nicht gefunden ..." & pname)
                    Exit Sub
                End Try

                appInstance.EnableEvents = False
                appInstance.ScreenUpdating = False


                Dim tmpCollection As New Collection
                ' bestimme die Anzahl Zeilen, die benötigt wird 
                Dim anzahlZeilen As Integer = hproj.calcNeededLines(tmpCollection, awinSettings.drawphases, False)
                Call moveShapesDown(tmpCollection, hproj.tfZeile + 1, anzahlZeilen, 0)
                'Call ZeichneProjektinPlanTafel2(pname, hproj.tfZeile)
                Dim listCollection As New Collection
                Call ZeichneProjektinPlanTafel(tmpCollection, pname, hproj.tfZeile, listCollection, listCollection)


                appInstance.EnableEvents = True
                appInstance.ScreenUpdating = True
            Else
                Call MsgBox("bitte nur ein Projekt selektieren")
                'For Each singleShp In awinSelection
                '    With singleShp
                '        If .AutoShapeType = MsoAutoShapeType.msoShapeRoundedRectangle Then
                '            nrSelPshp = nrSelPshp + 1
                '            SID = .ID.ToString
                '        End If
                '    End With
                'Next
            End If
        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If

        enableOnUpdate = True

    End Sub

    ''' <summary>
    ''' Charakteristik Personal Bedarfe
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G2M1B2Resources(control As IRibbonControl)

        Dim singleShp As Excel.Shape
        'Dim SID As String
        Dim hproj As clsProjekt
        Dim awinSelection As Excel.ShapeRange
        Dim auswahl As Integer = 1
        Dim top As Double, left As Double, width As Double, height As Double

        Call projektTafelInit()

        enableOnUpdate = False

        Try
            'awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            If awinSelection.Count = 1 Then
                ' jetzt die Aktion durchführen ...
                singleShp = awinSelection.Item(1)
                With singleShp
                    top = .Top + boxHeight + 5
                    left = .Left - 5
                End With
                height = 180

                Try
                    hproj = ShowProjekte.getProject(singleShp.Name)
                Catch ex As Exception
                    Call MsgBox("Projekt " & singleShp.Name & " nicht gefunden ...")
                    Exit Sub
                End Try

                Dim repObj As Excel.ChartObject
                appInstance.EnableEvents = False
                appInstance.ScreenUpdating = False

                repObj = Nothing

                width = System.Math.Max(hproj.anzahlRasterElemente * boxWidth + 10, 6 * boxWidth + 10)

                Try
                    Call createRessBalkenOfProject(hproj, repObj, auswahl, top, left, height, width)

                    ' jetzt wird das Pie-Diagramm gezeichnet 
                    left = left + width + 10
                    width = boxWidth * 14
                    height = boxHeight * 10
                    repObj = Nothing
                    Call createRessPieOfProject(hproj, repObj, auswahl, top, left, height, width)
                Catch ex As Exception
                    Call MsgBox(ex.Message)
                End Try


                appInstance.EnableEvents = True
                appInstance.ScreenUpdating = True
            Else
                Call MsgBox("bitte nur ein Projekt selektieren")

            End If
        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If

        enableOnUpdate = True



    End Sub

    ''' <summary>
    ''' Charakteristik Personalkosten
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G2M1B3PKosten(control As IRibbonControl)

        Dim singleShp As Excel.Shape
        'Dim SID As String
        Dim hproj As clsProjekt
        Dim awinSelection As Excel.ShapeRange
        Dim auswahl As Integer = 2 ' steuert die Auswahl als Personalkosten
        Dim top As Double, left As Double, width As Double, height As Double

        Call projektTafelInit()

        enableOnUpdate = False

        Try
            'awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            If awinSelection.Count = 1 Then
                ' jetzt die Aktion durchführen ...
                singleShp = awinSelection.Item(1)
                With singleShp
                    top = .Top + boxHeight + 5
                    left = .Left - 5
                End With
                height = 180

                Try
                    hproj = ShowProjekte.getProject(singleShp.Name)
                Catch ex As Exception
                    Call MsgBox("Projekt " & singleShp.Name & " nicht gefunden ...")
                    Exit Sub
                End Try



                width = hproj.anzahlRasterElemente * boxWidth + 10

                appInstance.EnableEvents = False
                appInstance.ScreenUpdating = False
                Dim repObj As Excel.ChartObject = Nothing

                Try
                    Call createRessBalkenOfProject(hproj, repObj, auswahl, top, left, height, width)

                    ' jetzt wird das Pie-Diagramm gezeichnet 
                    left = left + width + 10
                    width = boxWidth * 14
                    height = boxHeight * 10
                    repObj = Nothing
                    Call createRessPieOfProject(hproj, repObj, auswahl, top, left, height, width)
                Catch ex As Exception
                    Call MsgBox(ex.Message)
                End Try


                appInstance.EnableEvents = True
                appInstance.ScreenUpdating = True

            Else
                Call MsgBox("bitte nur ein Projekt selektieren")

            End If
        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If

        enableOnUpdate = True



    End Sub

    ''' <summary>
    ''' Charakteristik Andere Kosten
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G2M1B4AKosten(control As IRibbonControl)

        Dim singleShp As Excel.Shape
        'Dim SID As String
        Dim hproj As clsProjekt
        Dim awinSelection As Excel.ShapeRange
        Dim auswahl As Integer = 1
        Dim top As Double, left As Double, width As Double, height As Double

        Call projektTafelInit()

        enableOnUpdate = False

        Try
            'awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            If awinSelection.Count = 1 Then
                ' jetzt die Aktion durchführen ...
                singleShp = awinSelection.Item(1)
                With singleShp
                    top = .Top + boxHeight + 5
                    left = .Left - 5
                End With
                height = 180

                Try
                    hproj = ShowProjekte.getProject(singleShp.Name)
                Catch ex As Exception
                    Call MsgBox("Projekt " & singleShp.Name & " nicht gefunden ...")
                    Exit Sub
                End Try

                width = hproj.anzahlRasterElemente * boxWidth + 10
                appInstance.EnableEvents = False
                appInstance.ScreenUpdating = False
                Dim repObj As Excel.ChartObject = Nothing

                Call createCostBalkenOfProject(hproj, repObj, auswahl, top, left, height, width)

                ' jetzt wird das Pie-Diagramm gezeichnet 
                left = left + width + 10
                width = boxWidth * 14
                height = boxHeight * 10
                repObj = Nothing

                Try
                    Call createCostPieOfProject(hproj, repObj, auswahl, top, left, height, width)
                Catch ex As Exception
                    Call MsgBox(ex.Message)
                End Try


                appInstance.EnableEvents = True
                appInstance.ScreenUpdating = True

            Else
                Call MsgBox("bitte nur ein Projekt selektieren")

            End If
        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If

        enableOnUpdate = True


    End Sub

    ''' <summary>
    ''' Charakteristik Gesamtkosten
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G2M1B5GKosten(control As IRibbonControl)

        Dim singleShp As Excel.Shape
        Dim hproj As clsProjekt
        Dim awinSelection As Excel.ShapeRange
        Dim auswahl As Integer = 2 ' das steuert , dass die Gesamtkosten angezeigt werden 
        Dim top As Double, left As Double, width As Double, height As Double

        Call projektTafelInit()

        enableOnUpdate = False

        Try
            'awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            If awinSelection.Count = 1 Then
                ' jetzt die Aktion durchführen ...
                singleShp = awinSelection.Item(1)
                With singleShp
                    top = .Top + boxHeight + 5
                    left = .Left - 5
                End With
                height = 180

                Try
                    hproj = ShowProjekte.getProject(singleShp.Name)
                Catch ex As Exception
                    Call MsgBox("Projekt " & singleShp.Name & " nicht gefunden ...")
                    Exit Sub
                End Try

                width = hproj.anzahlRasterElemente * boxWidth + 10

                appInstance.EnableEvents = False
                appInstance.ScreenUpdating = False
                Dim repObj As Excel.ChartObject = Nothing


                Try
                    Call createCostBalkenOfProject(hproj, repObj, auswahl, top, left, height, width)
                    ' jetzt wird das Pie-Diagramm gezeichnet 
                    left = left + width + 10
                    width = boxWidth * 14
                    height = boxHeight * 10
                    repObj = Nothing
                    Call createCostPieOfProject(hproj, repObj, auswahl, top, left, height, width)
                Catch ex As Exception
                    Call MsgBox(ex.Message)
                End Try


                appInstance.EnableEvents = True
                appInstance.ScreenUpdating = True

            Else
                Call MsgBox("bitte nur ein Projekt selektieren")

            End If
        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If

        enableOnUpdate = True


    End Sub



    ''' <summary>
    ''' Charakteristik Strategie / Risiko 
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G2M1B6SFIT(control As IRibbonControl)


        Dim top As Double, left As Double, width As Double, height As Double
        Dim singleShp As Excel.Shape
        Dim myCollection As New Collection
        Dim awinSelection As Excel.ShapeRange

        Call projektTafelInit()

        Dim formerSU As Boolean = appInstance.ScreenUpdating
        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False

        enableOnUpdate = False

        Try
            'awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            ' jetzt die Aktion durchführen ...

            For Each singleShp In awinSelection
                Dim shapeArt As Integer
                shapeArt = kindOfShape(singleShp)

                With singleShp
                    If isProjectType(shapeArt) Then

                        myCollection.Add(.Name)
                        top = .Top + boxHeight + 2
                        left = .Left - 3
                        width = 12 * boxWidth
                        height = 8 * boxHeight

                    End If
                End With
            Next
            Dim obj As Excel.ChartObject = Nothing
            Call awinCreatePortfolioDiagrams(myCollection, obj, True, PTpfdk.FitRisiko, 0, False, True, True, top, left, width, height)
        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If

        enableOnUpdate = True
        appInstance.EnableEvents = formerEE
        appInstance.ScreenUpdating = formerSU


    End Sub

    Sub Tom2G2M1B6SFITVOl(control As IRibbonControl)

        Dim top As Double, left As Double, width As Double, height As Double
        Dim singleShp As Excel.Shape
        Dim myCollection As New Collection
        Dim awinSelection As Excel.ShapeRange

        Call projektTafelInit()

        Dim formerSU As Boolean = appInstance.ScreenUpdating
        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False

        enableOnUpdate = False

        Try
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            ' jetzt die Aktion durchführen ...

            For Each singleShp In awinSelection
                Dim shapeArt As Integer
                shapeArt = kindOfShape(singleShp)

                With singleShp
                    If isProjectType(shapeArt) Then

                        myCollection.Add(.Name)
                        top = .Top + boxHeight + 2
                        left = .Left - 3
                        width = 12 * boxWidth
                        height = 8 * boxHeight

                    End If
                End With
            Next
            Dim obj As Excel.ChartObject = Nothing

            Call awinCreatePortfolioDiagrams(myCollection, obj, True, PTpfdk.FitRisikoVol, 0, False, True, True, top, left, width, height)
            'Call awinCreateStratRiskVolumeDiagramm(myCollection, obj, True, False, True, True, top, left, width, height)
        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If

        enableOnUpdate = True
        appInstance.EnableEvents = formerEE
        appInstance.ScreenUpdating = formerSU

    End Sub

    Sub Tom2G2M1B6Abhaengigkeit(control As IRibbonControl)


        Dim top As Double, left As Double, width As Double, height As Double
        Dim singleShp As Excel.Shape
        Dim myCollection As New Collection
        Dim deleteList As New Collection
        Dim hproj As clsProjekt
        Dim pname As String

        Dim activeNumber As Integer             ' Kennzahl: auf wieviele Projekte strahlt es aus ?
        Dim passiveNumber As Integer            ' Kennzahl: von wievielen Projekten abhängig 
        Dim awinSelection As Excel.ShapeRange

        Call projektTafelInit()

        Dim formerSU As Boolean = appInstance.ScreenUpdating
        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False

        enableOnUpdate = False

        Try
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            ' jetzt die Aktion durchführen ...

            For Each singleShp In awinSelection
                Dim shapeArt As Integer
                shapeArt = kindOfShape(singleShp)

                With singleShp
                    If isProjectType(shapeArt) Then

                        myCollection.Add(.Name, .Name)
                        top = .Top + boxHeight + 2
                        left = .Left - 3
                        width = 12 * boxWidth
                        height = 8 * boxHeight

                    End If
                End With
            Next

            Dim i As Integer
            For i = 1 To myCollection.Count
                pname = CStr(myCollection.Item(i))
                Try
                    hproj = ShowProjekte.getProject(pname)
                    activeNumber = allDependencies.activeNumber(pname, PTdpndncyType.inhalt)
                    passiveNumber = allDependencies.passiveNumber(pname, PTdpndncyType.inhalt)
                    If activeNumber = 0 And passiveNumber = 0 Then
                        deleteList.Add(pname)
                    End If
                Catch ex As Exception

                End Try
            Next

            ' jetzt müssen die Projekte rausgenommen werden, die keine Abhängigkeiten haben 
            For i = 1 To deleteList.Count
                pname = CStr(deleteList.Item(i))
                Try
                    myCollection.Remove(pname)
                Catch ex As Exception

                End Try
            Next

            If myCollection.Count > 0 Then
                Dim obj As Excel.ChartObject = Nothing
                Call awinCreatePortfolioDiagrams(myCollection, obj, True, PTpfdk.Dependencies, 0, False, True, True, top, left, width, height)
            Else
                Call MsgBox("diese Projekte haben keine Abhängigkeiten")
            End If



        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If

        enableOnUpdate = True
        appInstance.EnableEvents = formerEE
        appInstance.ScreenUpdating = formerSU


    End Sub

    Sub Tom2G2M1B6CRisk(control As IRibbonControl)

        Dim top As Double, left As Double, width As Double, height As Double
        Dim singleShp As Excel.Shape
        Dim myCollection As New Collection
        Dim sichtbarerBereich As Excel.Range
        Dim awinSelection As Excel.ShapeRange

        Call projektTafelInit()

        Dim formerSU As Boolean = appInstance.ScreenUpdating
        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False

        enableOnUpdate = False

        Try
            'awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            ' jetzt die Aktion durchführen ...

            For Each singleShp In awinSelection
                Dim shapeArt As Integer
                shapeArt = kindOfShape(singleShp)

                With singleShp
                    If isProjectType(shapeArt) Then

                        myCollection.Add(.Name)
                        top = .Top + boxHeight + 2
                        left = .Left - 3
                        width = 300
                        height = 280

                    End If
                End With
            Next

            If myCollection.Count > 1 Then

                With appInstance.ActiveWindow
                    sichtbarerBereich = .VisibleRange
                    left = CDbl(sichtbarerBereich.Left) + (CDbl(sichtbarerBereich.Width) - 500) / 2
                    top = CDbl(sichtbarerBereich.Top) + (CDbl(sichtbarerBereich.Height) - 450) / 2
                End With

                width = 500
                height = 450
            End If

            Dim obj As Excel.ChartObject = Nothing

            Try
                Call awinCreatePortfolioDiagrams(myCollection, obj, True, PTpfdk.ComplexRisiko, 0, False, True, True, top, left, width, height)
            Catch ex As Exception

            End Try

        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If

        enableOnUpdate = True
        appInstance.EnableEvents = formerEE
        appInstance.ScreenUpdating = formerSU

        ' die Projekte sollen hier doch nicht deselektiert werden, weil dadurch die awinNeuZeichnenDiagramm aufgerufen wird und damit auch die awinUpdatePortfolioDiagrams
        ' was dazu führt, dass alle Projekt in der Projektliste wieder in das Diagramm eingezeichnet werden.
        'Call awinDeSelect()



    End Sub


    ''' <summary>
    ''' zeigt den Soll-Ist Vergleich für das gewählte Projekt an 
    ''' Beauftragung / letzter Plan-Stand / aktueller Plan-Stand
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G2M2M1B2SollIstPKosten(control As IRibbonControl)

        Call projektTafelInit()
        ' auswahl steuert , dass die Personal-Kosten angezeigt werden 
        Dim auswahl As Integer = 1

        Dim vglBaseline As Boolean = True

        ' typ steuert, ob Summenbetrachtung oder Curve angezeigt wird
        Dim typ As String = " "

        Call awinSollIstVergleich(auswahl, typ, vglBaseline)

    End Sub

    ''' <summary>
    ''' zeigt den Soll-Ist Vergleich für das gewählte Projekt an 
    ''' Beauftragung / letzter Plan-Stand / aktueller Plan-Stand
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G2M2M2B2SollIstAKosten(control As IRibbonControl)
        ' auswahl steuert , welche Kosten angezeigt werden
        Dim auswahl As Integer = 2
        Dim vglBaseline As Boolean = True
        ' typ steuert, ob Summenbetrachtung oder Curve angezeigt wird
        Dim typ As String = " "

        Call projektTafelInit()

        Call awinSollIstVergleich(auswahl, typ, vglBaseline)

    End Sub


    Sub Tom2G2M2M3B2SollIstGKosten(control As IRibbonControl)

        ' auswahl steuert , welche Kosten angezeigt werden
        Dim auswahl As Integer = 3
        Dim vglBaseline As Boolean = True

        ' typ steuert, ob Summenbetrachtung oder Curve angezeigt wird
        Dim typ As String = " "

        Call projektTafelInit()

        Call awinSollIstVergleich(auswahl, typ, vglBaseline)

    End Sub

    ''' <summary>
    ''' Fortschritts-Chart im Vergleich zur Beauftragung
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G2M4Fortschritt1(control As IRibbonControl)

        Call projektTafelInit()

        Call awinStatusAnzeige(1, 1, " ")

    End Sub

    ''' <summary>
    ''' Fortschritts-Chart im Vergleich zur letzten Planungs-Freigabe
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G2M4Fortschritt2(control As IRibbonControl)

        Call projektTafelInit()
        Call awinStatusAnzeige(2, 1, " ")

    End Sub



    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="auswahl"></param>
    ''' <param name="typ"></param>
    ''' <remarks></remarks>
    Private Sub awinSollIstVergleich(ByVal auswahl As Integer, ByVal typ As String, ByVal vglBaseline As Boolean)
        Dim request As New Request(awinSettings.databaseName, dbUsername, dbPasswort)
        Dim singleShp As Excel.Shape
        Dim hproj As clsProjekt
        Dim awinSelection As Excel.ShapeRange
        Dim top As Double, left As Double, width As Double, height As Double
        Dim reportObj As Excel.ChartObject
        Dim heute As Date = Date.Now
        Dim vglName As String = " "
        Dim pName As String = ";"
        Dim variantName As String = ""

        Call projektTafelInit()

        enableOnUpdate = False

        Try
            'awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            If awinSelection.Count = 1 Then
                ' jetzt die Aktion durchführen ...
                singleShp = awinSelection.Item(1)
                With singleShp
                    top = .Top + boxHeight + 5
                    left = .Left - 5
                End With
                height = 300
                width = 400

                Try
                    hproj = ShowProjekte.getProject(singleShp.Name)

                Catch ex As Exception
                    Call MsgBox("Projekt " & singleShp.Name & " nicht gefunden ...")
                    Exit Sub
                End Try

                If Not projekthistorie Is Nothing Then
                    If projekthistorie.Count > 0 Then
                        vglName = projekthistorie.First.getShapeText
                    End If
                Else
                    projekthistorie = New clsProjektHistorie
                End If

                With hproj
                    pName = .name
                    variantName = .variantName
                End With

                If vglName <> hproj.getShapeText Then
                    If request.pingMongoDb() Then
                        ' projekthistorie muss nur dann neu bestimmt werden, wenn sie nicht bereits für dieses Projekt geholt wurde
                        projekthistorie.liste = request.retrieveProjectHistoryFromDB(projectname:=pName, variantName:=variantName, _
                                                                            storedEarliest:=StartofCalendar, storedLatest:=Date.Now)
                        projekthistorie.Add(Date.Now, hproj)
                    Else
                        Call MsgBox(" Datenbank-Verbindung ist unterbrochen!" & vbLf & " Projekthistorie kann nicht geladen werden")
                        projekthistorie.clear()
                    End If

                Else
                    ' der aktuelle Stand hproj muss hinzugefügt werden 
                    Dim lastElem As Integer = projekthistorie.Count - 1
                    projekthistorie.RemoveAt(lastElem)
                    projekthistorie.Add(Date.Now, hproj)
                End If

                Dim nrSnapshots As Integer = projekthistorie.Count

                appInstance.EnableEvents = False
                appInstance.ScreenUpdating = False
                reportObj = Nothing

                Dim qualifier As String = " "

                Try
                    If typ = "Curve" Then
                        Call createSollIstCurveOfProject(hproj, reportObj, heute, auswahl, qualifier, vglBaseline, top, left, height, width)
                    Else
                        Call createSollIstOfProject(hproj, reportObj, heute, auswahl, qualifier, vglBaseline, top, left, height, width)
                    End If
                Catch ex As Exception

                End Try

                appInstance.EnableEvents = True
                appInstance.ScreenUpdating = True

            Else
                Call MsgBox("bitte nur ein Projekt selektieren")

            End If
        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If

        enableOnUpdate = True


    End Sub


    ''' <summary>
    ''' zeigt die Fortschrittsanzeige an 
    ''' </summary>
    ''' <param name="compareTyp">
    ''' 0=erste Eintrag in Projekt-Historie 
    ''' 1=Beauftragung
    ''' 2=letzte Freigabe
    ''' 3=letzter DB-Eintrag in Projekthistorie
    ''' </param>
    ''' <param name="auswahl">
    ''' 1=Personalkosten
    ''' 2=Sonstige Kosten
    ''' 3=Gesamtkosten
    ''' 3=rolle + Qualifier
    ''' 5=kostenart + qualifier
    ''' </param>
    ''' <param name="qualifier">
    ''' gibt an, um welche Rolle / Kostenart es sich handelt - falls auswahl = 4 oder 5
    ''' </param>
    ''' <remarks></remarks>
    Private Sub awinStatusAnzeige(ByVal compareTyp As Integer, ByVal auswahl As Integer, ByVal qualifier As String)
        Dim request As New Request(awinSettings.databaseName, dbUsername, dbPasswort)
        Dim singleShp As Excel.Shape
        Dim hproj As clsProjekt
        Dim awinSelection As Excel.ShapeRange
        Dim top As Double, left As Double, width As Double, height As Double
        Dim reportObj As Excel.ChartObject
        Dim heute As Date = Date.Now
        Dim vglName As String = " "
        Dim pName As String = ";"
        Dim variantName As String = ""
        Dim projektliste As New Collection
        Dim first As Boolean = True

        Call projektTafelInit()

        enableOnUpdate = False



        Try
            'awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try


        ' jetzt die Aktion durchführen für alle selektierten 

        If Not awinSelection Is Nothing Then

            ' jetzt die Aktion durchführen ...

            For Each singleShp In awinSelection
                Dim shapeArt As Integer
                shapeArt = kindOfShape(singleShp)

                With singleShp
                    If isProjectType(shapeArt) Then

                        Try
                            hproj = ShowProjekte.getProject(.Name)
                            pName = hproj.name
                            If istLaufendesProjekt(hproj) Then

                                Try
                                    projektliste.Add(pName, pName)
                                Catch ex1 As Exception

                                End Try

                            End If

                        Catch ex As Exception

                        End Try

                        If first Then
                            top = .Top + boxHeight + 5
                            left = .Left - 5
                            first = False
                        End If

                    End If
                End With
            Next

            height = 300
            width = 400

            If projektliste.Count > 0 Then
                ' Diagramm erstellen 

                appInstance.EnableEvents = False
                appInstance.ScreenUpdating = False
                reportObj = Nothing


                Dim tmpObj As Excel.ChartObject = Nothing
                Call awinCreateStatusDiagram1(projektliste, tmpObj, compareTyp, auswahl, qualifier, True, True, _
                                               top, left, width, height)

                reportObj = CType(tmpObj, Excel.ChartObject)
                appInstance.EnableEvents = True
                appInstance.ScreenUpdating = True


            End If
        Else
            Call MsgBox("es wurden keine laufenden Projekte selektert ...")
        End If


        ' Ende 


        enableOnUpdate = True


    End Sub



    Sub Tom2G2M5M2B5ShowMilestones(control As IRibbonControl)


        Dim farbTyp As Integer = 4
        Dim numberIt As Boolean = False
        Dim namelist As New Collection

        Call projektTafelInit()

        appInstance.EnableEvents = False
        enableOnUpdate = False

        Call awinZeichneMilestones(namelist, farbTyp, numberIt, False)

        enableOnUpdate = True
        appInstance.EnableEvents = True
        'appInstance.ScreenUpdating = formerSU


    End Sub

    ''' <summary>
    ''' zeigt bei den ausgewählten Projekten die gewählten  erst eine Liste, aus der man die Namen auswählen kann 
    ''' zeigt dann alle Meilensteine, die zu dieser Liste gehören 
    ''' wenn Projekte selektiert sind: zeige nur die Meilensteine dieser Projekte an 
    ''' wenn nichts selektiert ist: Fehler MEldung 
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub PTShowMilestonesByName(control As IRibbonControl)



        Dim listOfItems As New Collection
        Dim nameList As New Collection
        Dim title As String = "Meilensteine visualisieren"

        Dim repObj As Object = Nothing

        Dim singleShp As Excel.Shape
        Dim myCollection As New Collection
        Dim hproj As clsProjekt
        Dim awinSelection As Excel.ShapeRange
        Dim selektierteProjekte As New clsProjekte

        Call projektTafelInit()

        Try
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            ' jetzt die Aktion durchführen ...

            For Each singleShp In awinSelection

                Try
                    hproj = ShowProjekte.getProject(singleShp.Name)
                    selektierteProjekte.Add(hproj)
                Catch ex As Exception
                    Call MsgBox("Projekt " & singleShp.Name & " nicht gefunden ...")
                End Try

            Next

            nameList = selektierteProjekte.getMilestoneNames

            If nameList.Count > 0 Then

                For Each tmpName As String In nameList
                    listOfItems.Add(tmpName)
                Next

                ' jetzt stehen in der listOfItems die Namen der Meilensteine - alphabetisch sortiert 
                Dim auswahlFenster As New ListSelectionWindow(listOfItems, title, "andere löschen")


                With auswahlFenster

                    .chTyp = DiagrammTypen(5)

                End With
                auswahlFenster.Show()

            Else
                Call MsgBox("keine Meilensteine in den selektierten Projekten vorhanden ..")
            End If


        Else
            Call MsgBox("Bitte mindestens ein Projekt selektieren ... ")
            Exit Sub
        End If







    End Sub

    ''' <summary>
    ''' zeigt bei den ausgewählten Projekten die gewählten  erst eine Liste, aus der man die Namen auswählen kann 
    ''' zeigt dann alle Meilensteine, die zu dieser Liste gehören 
    ''' wenn Projekte selektiert sind: zeige nur die Meilensteine dieser Projekte an 
    ''' wenn nichts selektiert ist: zeige die Namen der Meilensteine aus allen Projekten  
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Public Sub PTShowAllMilestonesByName(Control As IRibbonControl)

        Dim listOfItems As New Collection
        Dim nameList As New Collection
        Dim title As String = "Meilensteine visualisieren"

        Dim repObj As Object = Nothing

        Call projektTafelInit()
        Call awinDeSelect()

        If ShowProjekte.Count > 0 Then
            If showRangeRight - showRangeLeft > 5 Then

                nameList = ShowProjekte.getMilestoneNames

                If nameList.Count > 0 Then

                    For Each tmpName As String In nameList
                        listOfItems.Add(tmpName)
                    Next

                    ' jetzt stehen in der listOfItems die Namen der Meilensteine - alphabetisch sortiert 
                    Dim auswahlFenster As New ListSelectionWindow(listOfItems, title, "andere löschen")


                    With auswahlFenster

                        .chTyp = DiagrammTypen(5)

                    End With
                    auswahlFenster.Show()

                Else
                    Call MsgBox("keine Meilensteine in den selektierten Projekten vorhanden ..")
                End If
            Else
                Call MsgBox("Bitte wählen Sie einen Zeitraum aus!")
            End If
        Else
            Call MsgBox("Es sind keine Projekte geladen!")
        End If



    End Sub


    ''' <summary>
    ''' zeigt zu dem ausgewählten Projekt die Meilenstein Trendanalyse an 
    ''' dazu wird erst ein Fenster aufgeschaltet, aus dem der oder die Namen des betreffenden Meilensteins ausgewählt werden können 
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub PTShowMilestoneTrend(control As IRibbonControl)

        Dim request As New Request(awinSettings.databaseName, dbUsername, dbPasswort)
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
                        hproj = ShowProjekte.getProject(singleShp.Name)
                        nameList = hproj.getMilestones

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

                        If nameList.Count > 0 Then


                            appInstance.EnableEvents = False
                            enableOnUpdate = False

                            repObj = Nothing



                            For Each kvp As KeyValuePair(Of Date, String) In nameList
                                listOfItems.Add(kvp.Value)
                            Next

                            With singleShp
                                top = .Top + boxHeight + 5
                                left = .Left - 5
                            End With

                            height = 2 * ((nameList.Count - 1) * 20 + 110)
                            width = System.Math.Max(hproj.anzahlRasterElemente * boxWidth + 10, 24 * boxWidth + 10)

                            'Try

                            '    Call createMsTrendAnalysisOfProject(hproj, repObj, listOfItems, top, left, height, width)

                            'Catch ex As Exception

                            '    Call MsgBox(ex.Message)

                            'End Try



                            ' jetzt stehen in der listOfItems die Namen der Meilensteine - alphabetisch sortiert 
                            Dim auswahlFenster As New ListSelectionWindow(listOfItems, title)


                            With auswahlFenster

                                .kennung = " "
                                .chTyp = DiagrammTypen(6)
                                .chTop = top
                                .chLeft = left
                                .chWidth = width
                                .chHeight = height

                            End With
                            auswahlFenster.Show()

                        Else
                            Call MsgBox("keine Meilensteine in den selektierten Projekten vorhanden ..")
                        End If

                    Catch ex As Exception
                        Call MsgBox("Projekt " & singleShp.Name & " nicht gefunden ...")
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

    Sub PT0ShowProjektStatus(control As IRibbonControl)

        Dim singleShp As Excel.Shape
        Dim myCollection As New Collection
        Dim hproj As clsProjekt
        Dim awinSelection As Excel.ShapeRange

        Call projektTafelInit()

        appInstance.EnableEvents = False
        enableOnUpdate = False

        Try
            'awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            ' jetzt die Aktion durchführen ...

            For Each singleShp In awinSelection

                Try
                    hproj = ShowProjekte.getProject(singleShp.Name)
                    Call zeichneStatusSymbolInPlantafel(hproj, 0)
                Catch ex As Exception
                    Call MsgBox("Projekt " & singleShp.Name & " nicht gefunden ...")
                End Try



            Next

        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If

        enableOnUpdate = True
        appInstance.EnableEvents = True





    End Sub

    ''' <summary>
    ''' zeigt die Abhängigkeiten der ausgewählten Projekte an ...
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub PT0ShowDependencies(control As IRibbonControl)

        Dim singleShp As Excel.Shape
        Dim myCollection As New Collection
        Dim hproj As clsProjekt
        Dim awinSelection As Excel.ShapeRange
        Dim atleastOne As Boolean = False

        Call projektTafelInit()

        appInstance.EnableEvents = False
        enableOnUpdate = False

        Try
            'awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            ' jetzt die Aktion durchführen ...

            ' erst noch alle Connectoren löschen ... 

            Call awinDeleteProjectChildShapes(4)

            For Each singleShp In awinSelection

                Try

                    hproj = ShowProjekte.getProject(singleShp.Name)
                    Call zeichneDependenciesOfProject(hproj, PTdpndncyType.inhalt, 0)
                    atleastOne = True

                Catch ex As Exception
                    'Call MsgBox("Projekt " & singleShp.Name & " hat keine Abhängigkeiten")
                End Try



            Next

        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If

        enableOnUpdate = True
        appInstance.EnableEvents = True



    End Sub

    Sub Tom2G2M5B3NoShowSymbols(control As IRibbonControl)
        Call projektTafelInit()
        Call awinDeleteProjectChildShapes(0)
    End Sub


    ''' <summary>
    ''' löscht alle angezeigten Milestones
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G2M5B3NoShowMilestones(control As IRibbonControl)

        Call projektTafelInit()
        Call awinDeleteProjectChildShapes(1)

    End Sub

    Sub PT0VisualizePhases(control As IRibbonControl)

        Dim i As Integer
        Dim von As Integer, bis As Integer

        Dim listOfItems As New Collection
        Dim existingNames As New Collection

        Dim title As String = "Phasen visualisieren"
        Dim phaseName As String
        Dim hproj As clsProjekt


        Dim awinSelection As Excel.ShapeRange
        Dim selektierteProjekte As New clsProjekte
        Dim singleshp As Excel.Shape

        Call projektTafelInit()

        appInstance.EnableEvents = False
        enableOnUpdate = False

        Try
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try


        Dim anzElem As Integer = selektierteProjekte.Count

        If Not awinSelection Is Nothing Then

            ' jetzt die Aktion durchführen ...

            For Each singleshp In awinSelection

                Try
                    hproj = ShowProjekte.getProject(singleshp.Name)
                    selektierteProjekte.Add(hproj)
                Catch ex As Exception
                    Call MsgBox("Projekt " & singleshp.Name & " nicht gefunden ...")
                End Try

            Next


            existingNames = selektierteProjekte.getPhaseNames

            If existingNames.Count > 0 Then

                ' jetzt werden die Namen in der Reihenfolge, wie sie in der Phasen-Definition stehen in der listofItems eingetragen ..

                For i = 1 To PhaseDefinitions.Count
                    phaseName = PhaseDefinitions.getPhaseDef(i).name

                    If existingNames.Contains(phaseName) Then
                        listOfItems.Add(PhaseDefinitions.getPhaseDef(i).name)
                    End If

                Next

                ' jetzt stehen in der listOfItems die Namen der Phasen 
                Dim auswahlFenster As New ListSelectionWindow(listOfItems, title, "andere löschen")

                von = showRangeLeft
                bis = showRangeRight
                With auswahlFenster
                    .chTop = 50.0
                    .chLeft = (showRangeRight - 1) * boxWidth + 4
                    .chWidth = 265 + (bis - von - 12 + 1) * boxWidth + (bis - von) * screen_correct
                    .chHeight = awinSettings.ChartHoehe1
                    .chTyp = DiagrammTypen(0)

                End With
                auswahlFenster.Show()

            Else
                Call MsgBox("keine Phasen vorhanden ...")

            End If



        Else

            Call MsgBox("bitte mindestens ein Projekt selektieren ...")

        End If

        enableOnUpdate = True
        appInstance.EnableEvents = True



    End Sub

    Sub PT0VisualizePhasesAll(control As IRibbonControl)

        Dim i As Integer
        Dim von As Integer, bis As Integer

        Dim listOfItems As New Collection
        Dim existingNames As New Collection

        Dim title As String = "Phasen visualisieren"
        Dim phaseName As String

        Call projektTafelInit()
        Call awinDeSelect()

        If ShowProjekte.Count > 0 Then

            If showRangeRight - showRangeLeft > 5 Then


                existingNames = ShowProjekte.getPhaseNames

                ' jetzt werden die Namen in der Reihenfolge, wie sie in der Phasen-Definition stehen in der listofItems eingetragen ..

                For i = 1 To PhaseDefinitions.Count
                    phaseName = PhaseDefinitions.getPhaseDef(i).name

                    If existingNames.Contains(phaseName) Then
                        listOfItems.Add(PhaseDefinitions.getPhaseDef(i).name)
                    End If

                Next

                ' jetzt stehen in der listOfItems die Namen der Phasen 
                Dim auswahlFenster As New ListSelectionWindow(listOfItems, title, "andere löschen")

                von = showRangeLeft
                bis = showRangeRight
                With auswahlFenster
                    .chTop = 50.0
                    .chLeft = (showRangeRight - 1) * boxWidth + 4
                    .chWidth = 265 + (bis - von - 12 + 1) * boxWidth + (bis - von) * screen_correct
                    .chHeight = awinSettings.ChartHoehe1
                    .chTyp = DiagrammTypen(0)

                End With
                auswahlFenster.Show()
            Else
                Call MsgBox("Bitte wählen Sie einen Zeitraum aus!")
            End If
        Else
            Call MsgBox("Es sind keine Projekte geladen!")
        End If
    End Sub


    Sub PT0ShowPortfolioPhasen(control As IRibbonControl)

        Dim i As Integer
        Dim von As Integer, bis As Integer
        'Dim myCollection As Collection
        Dim listOfItems As New Collection
        'Dim left As Double, top As Double, height As Double, width As Double

        Dim phaseName As String

        Call projektTafelInit()

        If ShowProjekte.Count > 0 Then

            If showRangeRight - showRangeLeft > 5 Then

                For i = 1 To PhaseDefinitions.Count
                    phaseName = PhaseDefinitions.getPhaseDef(i).name
                    Try
                        listOfItems.Add(phaseName, phaseName)
                    Catch ex As Exception

                    End Try

                Next

                ' jetzt stehen in der listOfItems die Namen der Rollen 
                Dim auswahlFenster As New ListSelectionWindow(listOfItems, "Phasen auswählen", "pro Item ein Chart")

                von = showRangeLeft
                bis = showRangeRight
                With auswahlFenster
                    .chTop = 50.0 + awinSettings.ChartHoehe1
                    .chLeft = (showRangeRight - 1) * boxWidth + 4
                    .chWidth = 265 + (bis - von - 12 + 1) * boxWidth + (bis - von) * screen_correct
                    .chHeight = awinSettings.ChartHoehe1
                    .chTyp = DiagrammTypen(0)
                End With
                auswahlFenster.Show()
            Else
                Call MsgBox("Bitte wählen Sie einen Zeitraum aus!")
            End If
        Else
            Call MsgBox("Es sind noch keine Projekte geladen!")
        End If
    End Sub

    Sub PTShowMilestoneSummen(control As IRibbonControl)

        Dim von As Integer, bis As Integer

        Dim listOfItems As New Collection

        Dim nameList As New Collection

        Call projektTafelInit()

        If ShowProjekte.Count > 0 Then
            If showRangeRight - showRangeLeft > 5 Then

                nameList = ShowProjekte.getMilestoneNames

                If nameList.Count > 0 Then

                    For Each tmpName As String In nameList
                        listOfItems.Add(tmpName)
                    Next

                    ' jetzt stehen in der listOfItems die Namen der Rollen 
                    Dim auswahlFenster As New ListSelectionWindow(listOfItems, "Meilensteine auswählen", "pro Item ein Chart")

                    von = showRangeLeft
                    bis = showRangeRight
                    With auswahlFenster
                        .kennung = "sum"
                        .chTop = 50.0 + awinSettings.ChartHoehe1
                        .chLeft = (showRangeRight - 1) * boxWidth + 4
                        .chWidth = 265 + (bis - von - 12 + 1) * boxWidth + (bis - von) * screen_correct
                        .chHeight = awinSettings.ChartHoehe1
                        .chTyp = DiagrammTypen(5)
                    End With
                    auswahlFenster.Show()


                Else
                    Call MsgBox("keine Meilensteine in den selektierten Projekten vorhanden ..")
                End If
            Else
                Call MsgBox("Bitte wählen Sie einen Zeitraum aus!")
            End If
        Else
            Call MsgBox("Es sind keine Projekte geladen!")
        End If

    End Sub


    Sub PT0ShowAuslastung(control As IRibbonControl)

        Dim selectionType As Integer = -1 ' Keine Einschränkung
        Dim top As Double, left As Double, width As Double, height As Double
        Dim obj As Excel.ChartObject = Nothing
        Dim myCollection As New Collection

        Call projektTafelInit()

        appInstance.ScreenUpdating = False
        appInstance.EnableEvents = False
        enableOnUpdate = False


        myCollection = ShowProjekte.withinTimeFrame(selectionType, showRangeLeft, showRangeRight)

        If myCollection.Count > 0 Then

            top = 180
            width = 340
            left = showRangeRight * boxWidth + 4
            If left < 0 Then
                left = 4
            End If
            height = awinSettings.ChartHoehe2

            Try
                Call awinCreateAuslastungsDiagramm(obj, top, left, width, height, False)

                top = top + height + 10
                Call createAuslastungsDetailPie(obj, 1, top, left, height, width, False)

                ' jetzt Unterauslastung
                top = top + height + 10
                Call createAuslastungsDetailPie(obj, 2, top, left, height, width, False)

            Catch ex As Exception
                Call MsgBox("keine Information vorhanden")
            End Try

        Else

            If ShowProjekte.Count = 0 Then
                Call MsgBox("es sind keine Projekte angezeigt")

            Else
                If showRangeRight - showRangeLeft < 6 Then
                    Call MsgBox(" Bitte wählen Sie zuerst einen Zeitraum aus !")
                Else
                    Call MsgBox("im angezeigten Zeitraum " & textZeitraum(showRangeLeft, showRangeRight) & vbLf & _
                                "gibt es keine Projekte ")
                End If
            End If

        End If



        appInstance.ScreenUpdating = True
        appInstance.EnableEvents = True
        enableOnUpdate = True

    End Sub

    Sub PTXShowEngpass(control As IRibbonControl)

        Dim i As Integer
        Dim von As Integer, bis As Integer
        Dim myCollection As New Collection
        Dim listOfItems As New Collection
        Dim left As Double, top As Double, height As Double, width As Double
        Dim roleName As String
        Dim engpass As String = ""
        Dim engpassValue As Double = -100000.0
        Dim curValue As Double

        Dim repObj As Excel.ChartObject = Nothing

        Call projektTafelInit()

        'appInstance.ScreenUpdating = False
        appInstance.EnableEvents = False
        enableOnUpdate = False

        If (showRangeRight - showRangeLeft) >= 6 Then

            If ShowProjekte.Count > 0 Then

                For i = 1 To RoleDefinitions.Count
                    roleName = RoleDefinitions.getRoledef(i).name
                    With ShowProjekte
                        curValue = .getAuslastungsValues(roleName, 1).Sum - .getAuslastungsValues(roleName, 2).Sum
                        If curValue > engpassValue Then
                            engpassValue = curValue
                            engpass = roleName
                        End If
                    End With
                Next

                If engpass <> "" Then
                    myCollection.Add(engpass, engpass)
                    von = showRangeLeft
                    bis = showRangeRight

                    height = awinSettings.ChartHoehe1
                    top = 180

                    If von > 1 Then
                        left = showRangeRight * boxWidth + 4
                    Else
                        left = 0
                    End If

                    width = 265 + (bis - von - 12 + 1) * boxWidth + (bis - von) * screen_correct

                    Call awinCreateprcCollectionDiagram(myCollection, repObj, top, left, width, height, False, DiagrammTypen(1), False)

                Else
                    Call MsgBox("kein Engpass gefunden")
                End If
            Else
                Call MsgBox("Es sind keine Projekte geladen! ")
            End If
        Else
            Call MsgBox("Bitte wählen Sie zuerst einen Zeitraum aus, der mindestens 6 Monate lang ist!")
        End If

        'appInstance.ScreenUpdating = True
        appInstance.EnableEvents = True
        enableOnUpdate = True

    End Sub

    Sub PT0ShowPersonalBedarfe(control As IRibbonControl)

        Dim i As Integer
        Dim von As Integer, bis As Integer
        'Dim myCollection As Collection
        Dim listOfItems As New Collection
        'Dim left As Double, top As Double, height As Double, width As Double

        Dim repObj As Object = Nothing
        Dim title As String = "Rollen auswählen"

        Call projektTafelInit()

        'appInstance.ScreenUpdating = False
        'appInstance.EnableEvents = False
        'enableOnUpdate = False

        If ShowProjekte.Count > 0 Then

            If showRangeRight - showRangeLeft > 5 Then


                For i = 1 To RoleDefinitions.Count
                    listOfItems.Add(RoleDefinitions.getRoledef(i).name)
                Next

                ' jetzt stehen in der listOfItems die Namen der Rollen 
                Dim auswahlFenster As New ListSelectionWindow(listOfItems, title, "pro Item ein Chart")

                von = showRangeLeft
                bis = showRangeRight
                With auswahlFenster
                    .chTop = 100.0 + awinSettings.ChartHoehe1
                    .chLeft = ((von - 1) / 3 - 1) * 3 * boxWidth + 32.8 + von * screen_correct
                    .chWidth = 265 + (bis - von - 12 + 1) * boxWidth + (bis - von) * screen_correct
                    .chHeight = awinSettings.ChartHoehe1
                    .chTyp = DiagrammTypen(1)
                End With

                auswahlFenster.Show()
            Else
                Call MsgBox("Bitte wählen Sie einen Zeitraum aus!")
            End If
        Else
            Call MsgBox("Es sind noch keine Projekte geladen!")
        End If

        'appInstance.ScreenUpdating = True
        'appInstance.EnableEvents = True
        'enableOnUpdate = True

    End Sub

    Sub PT0ShowKostenBedarfe(control As IRibbonControl)

        Dim i As Integer
        Dim von As Integer, bis As Integer
        'Dim myCollection As Collection
        Dim listOfItems As New Collection
        'Dim left As Double, top As Double, height As Double, width As Double
        Dim repObj As Object = Nothing
        Dim title As String = "Kostenarten auswählen"

        Call projektTafelInit()

        'appInstance.EnableEvents = False
        'enableOnUpdate = False
        If ShowProjekte.Count > 0 Then

            If showRangeRight - showRangeLeft > 5 Then

                For i = 1 To CostDefinitions.Count
                    listOfItems.Add(CostDefinitions.getCostdef(i).name)
                Next

                ' jetzt stehen in der listOfItems die Namen der Rollen 
                'Dim auswahlFenster As New ListSelectionWindow(listOfItems, title)
                Dim auswahlFenster As New ListSelectionWindow(listOfItems, title, "pro Item ein Chart")


                von = showRangeLeft
                bis = showRangeRight
                With auswahlFenster
                    .chTop = 50.0
                    .chLeft = (showRangeRight - 1) * boxWidth + 4
                    .chWidth = 265 + (bis - von - 12 + 1) * boxWidth + (bis - von) * screen_correct
                    .chHeight = awinSettings.ChartHoehe1
                    .chTyp = DiagrammTypen(2)

                End With

                auswahlFenster.Show()


            Else
                Call MsgBox("Bitte wählen Sie einen Zeitraum von mindestens 6 Monaten aus!")
            End If
        Else
            Call MsgBox("Es sind noch keine Projekte geladen!")
        End If

        'appInstance.EnableEvents = True
        'enableOnUpdate = True

    End Sub

    Sub PT0ShowZieleUebersicht(control As IRibbonControl)

        Dim chtObject As Excel.ChartObject = Nothing
        'Dim top As Double, left As Double, width As Double, height As Double
        Dim future As Integer = 0

        Dim myCollection As New Collection
        myCollection.Add("Ziele")

        Call projektTafelInit()

        appInstance.EnableEvents = False
        enableOnUpdate = False
        If ShowProjekte.Count > 0 Then
            If showRangeRight - showRangeLeft > 5 Then

                ' betrachte sowohl Vergangenheit als auch Gegenwart
                future = 0

                Dim wpfInput As New Dictionary(Of String, clsWPFPieValues)
                Dim valueItem As New clsWPFPieValues

                ' Nicht bewertet 
                With valueItem
                    .value = ShowProjekte.getColorsInMonth(0, future).Sum
                    .name = "nicht bewertet"
                    .color = CType(awinSettings.AmpelNichtBewertet, UInt32)
                End With
                wpfInput.Add(valueItem.name, valueItem)

                valueItem = New clsWPFPieValues
                ' Grün bewertet
                With valueItem
                    .value = ShowProjekte.getColorsInMonth(1, future).Sum
                    .name = "OK"
                    .color = CType(awinSettings.AmpelGruen, UInt32)
                End With
                wpfInput.Add(valueItem.name, valueItem)

                valueItem = New clsWPFPieValues
                ' Gelb bewertet
                With valueItem
                    .value = ShowProjekte.getColorsInMonth(2, future).Sum
                    .name = "nicht vollständig"
                    .color = CType(awinSettings.AmpelGelb, UInt32)
                End With
                wpfInput.Add(valueItem.name, valueItem)

                valueItem = New clsWPFPieValues
                ' Rot bewertet
                With valueItem
                    .value = ShowProjekte.getColorsInMonth(3, future).Sum
                    .name = "Zielverfehlung"
                    .color = CType(awinSettings.AmpelRot, UInt32)
                End With
                wpfInput.Add(valueItem.name, valueItem)


                Dim pieChartZieleV As New PieChartWindow(wpfInput)

                With pieChartZieleV
                    .Title = "Ziele-Erreichung " & textZeitraum(showRangeLeft, showRangeRight)
                    '.Top = frmCoord(PTfrm.ziele, PTpinfo.top)
                    '.Left = frmCoord(PTfrm.ziele, PTpinfo.left)
                End With

                pieChartZieleV.Show()
            Else
                Call MsgBox("Bitte wählen Sie einen Zeitraum aus!")
            End If

        Else
            Call MsgBox("Es sind keine Projekte geladen!")
        End If

        appInstance.EnableEvents = True
        enableOnUpdate = True

    End Sub



    Sub PT0ShowStrategieRisiko(control As IRibbonControl)

        Dim selectionType As Integer = -1 ' keine Einschränkung
        Dim myCollection As New Collection
        Dim top As Double, left As Double, width As Double, height As Double
        Dim sichtbarerBereich As Excel.Range

        Call projektTafelInit()

        appInstance.EnableEvents = False
        enableOnUpdate = False

        myCollection = ShowProjekte.withinTimeFrame(selectionType, showRangeLeft, showRangeRight)

        If myCollection.Count > 0 Then

            With appInstance.ActiveWindow
                sichtbarerBereich = .VisibleRange
                left = CDbl(sichtbarerBereich.Left) + (CDbl(sichtbarerBereich.Width) - 600) / 2
                If left < CDbl(sichtbarerBereich.Left) Then
                    left = CDbl(sichtbarerBereich.Left) + 2
                End If

                top = CDbl(sichtbarerBereich.Top) + (CDbl(sichtbarerBereich.Height) - 450) / 2
                If top < CDbl(sichtbarerBereich.Top) Then
                    top = CDbl(sichtbarerBereich.Top) + 2
                End If

            End With

            width = 600
            height = 450

            Dim obj As Excel.ChartObject = Nothing

            Try
                Call awinCreatePortfolioDiagrams(myCollection, obj, False, PTpfdk.FitRisiko, 0, False, True, True, top, left, width, height)
            Catch ex As Exception

            End Try

        Else

            If ShowProjekte.Count = 0 Then
                Call MsgBox("es sind keine Projekte angezeigt")

            Else
                Call MsgBox("im angezeigten Zeitraum " & textZeitraum(showRangeLeft, showRangeRight) & vbLf & _
                            "gibt es keine Projekte")
            End If


        End If



        appInstance.EnableEvents = True
        enableOnUpdate = True

    End Sub

    Sub PT0ShowStratRisikoVolume(control As IRibbonControl)

        Dim selectionType As Integer = -1 ' keine Einschränkung
        Dim myCollection As New Collection
        Dim top As Double, left As Double, width As Double, height As Double
        Dim sichtbarerBereich As Excel.Range

        Call projektTafelInit()

        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False
        enableOnUpdate = False

        myCollection = ShowProjekte.withinTimeFrame(selectionType, showRangeLeft, showRangeRight)

        If myCollection.Count > 0 Then

            With appInstance.ActiveWindow
                sichtbarerBereich = .VisibleRange
                left = CDbl(sichtbarerBereich.Left) + (CDbl(sichtbarerBereich.Width) - 600) / 2
                If left < CDbl(sichtbarerBereich.Left) Then
                    left = CDbl(sichtbarerBereich.Left) + 2
                End If

                top = CDbl(sichtbarerBereich.Top) + (CDbl(sichtbarerBereich.Height) - 450) / 2
                If top < CDbl(sichtbarerBereich.Top) Then
                    top = CDbl(sichtbarerBereich.Top) + 2
                End If

            End With

            width = 600
            height = 450

            Dim obj As Excel.ChartObject = Nothing

            Try
                Call awinCreatePortfolioDiagrams(myCollection, obj, False, PTpfdk.FitRisikoVol, 0, False, True, True, top, left, width, height)
                'Call awinCreateStratRiskVolumeDiagramm(myCollection, obj, False, False, True, True, top, left, width, height)
            Catch ex As Exception

            End Try

        Else

            If ShowProjekte.Count = 0 Then
                Call MsgBox("es sind keine Projekte angezeigt")

            Else
                Call MsgBox("im angezeigten Zeitraum " & textZeitraum(showRangeLeft, showRangeRight) & vbLf & _
                            "gibt es keine Projekte")
            End If

        End If


        appInstance.EnableEvents = True
        enableOnUpdate = True
        appInstance.ScreenUpdating = True

    End Sub

    Sub PT0ShowAbhaengigkeiten(control As IRibbonControl)

        Dim selectionType As Integer = -1 ' keine Einschränkung
        Dim myCollection As New Collection
        Dim top As Double, left As Double, width As Double, height As Double
        Dim sichtbarerBereich As Excel.Range
        Dim deleteList As New Collection
        Dim hproj As clsProjekt
        Dim pname As String

        Dim activeNumber As Integer             ' Kennzahl: auf wieviele Projekte strahlt es aus ?
        Dim passiveNumber As Integer            ' Kennzahl: von wievielen Projekten abhängig 

        Call projektTafelInit()

        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False
        enableOnUpdate = False

        myCollection = ShowProjekte.withinTimeFrame(selectionType, showRangeLeft, showRangeRight)

        If myCollection.Count > 0 Then


            Dim i As Integer
            For i = 1 To myCollection.Count
                pname = CStr(myCollection.Item(i))
                Try
                    hproj = ShowProjekte.getProject(pname)
                    activeNumber = allDependencies.activeNumber(pname, PTdpndncyType.inhalt)
                    passiveNumber = allDependencies.passiveNumber(pname, PTdpndncyType.inhalt)
                    If activeNumber = 0 And passiveNumber = 0 Then
                        deleteList.Add(pname)
                    End If
                Catch ex As Exception

                End Try
            Next

            ' jetzt müssen die Projekte rausgenommen werden, die keine Abhängigkeiten haben 
            For i = 1 To deleteList.Count
                pname = CStr(deleteList.Item(i))
                Try
                    myCollection.Remove(pname)
                Catch ex As Exception

                End Try
            Next


            With appInstance.ActiveWindow
                sichtbarerBereich = .VisibleRange
                left = CDbl(sichtbarerBereich.Left) + (CDbl(sichtbarerBereich.Width) - 600) / 2
                If left < CDbl(sichtbarerBereich.Left) Then
                    left = CDbl(sichtbarerBereich.Left) + 2
                End If

                top = CDbl(sichtbarerBereich.Top) + (CDbl(sichtbarerBereich.Height) - 450) / 2
                If top < CDbl(sichtbarerBereich.Top) Then
                    top = CDbl(sichtbarerBereich.Top) + 2
                End If

            End With

            width = 600
            height = 450

            Dim obj As Excel.ChartObject = Nothing

            Try
                If myCollection.Count > 0 Then
                    Call awinCreatePortfolioDiagrams(myCollection, obj, False, PTpfdk.Dependencies, 0, False, True, True, top, left, width, height)
                Else
                    Call MsgBox(" es gibt in diesem Zeitraum keine Projekte mit Abhängigkeiten")
                End If


            Catch ex As Exception

            End Try

        Else
            If ShowProjekte.Count = 0 Then
                Call MsgBox("es sind keine Projekte angezeigt")

            Else
                Call MsgBox("im angezeigten Zeitraum " & textZeitraum(showRangeLeft, showRangeRight) & vbLf & _
                            "gibt es keine Projekte")
            End If
        End If



        appInstance.EnableEvents = True
        enableOnUpdate = True
        appInstance.ScreenUpdating = True


    End Sub


    ''' <summary>
    ''' zeigt an , welche Projekte Management Attention verdienen/benötigen, weil sie besser/schlechter als der letzte Stand geplant laufen
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub PT0ShowAttentionL(control As IRibbonControl)

        Dim selectionType As Integer
        Dim myCollection As New Collection
        Dim top As Double, left As Double, width As Double, height As Double
        Dim sichtbarerBereich As Excel.Range
        Dim deleteList As New Collection

        Call projektTafelInit()

        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False
        enableOnUpdate = False

        ' hier muss noch geklärt werden, welche Projekte betrachtet werden; es mcht keinen Sinn, 
        'das nur an den TimeFrame zu koppeln, es geht im wesentlichen um aktuell laufende und vergangene Projekte 
        ' Frage : was ist mit bereits beauftragten Projekten, die noch gar nicht begonnen haben, deren Planung aber bereits schlechter als beauftragt ist ? 

        selectionType = PTpsel.lfundab
        myCollection = ShowProjekte.withinTimeFrame(selectionType, showRangeLeft, showRangeRight)

        If myCollection.Count > 0 Then

            With appInstance.ActiveWindow

                sichtbarerBereich = .VisibleRange
                left = CDbl(sichtbarerBereich.Left) + (CDbl(sichtbarerBereich.Width) - 600) / 2
                If left < CDbl(sichtbarerBereich.Left) Then
                    left = CDbl(sichtbarerBereich.Left) + 2

                End If

                top = CDbl(sichtbarerBereich.Top) + (CDbl(sichtbarerBereich.Height) - 450) / 2
                If top < CDbl(sichtbarerBereich.Top) Then
                    top = CDbl(sichtbarerBereich.Top) + 2
                End If

            End With

            width = 600
            height = 450

            Dim obj As Excel.ChartObject = Nothing

            Try
                If myCollection.Count > 0 Then

                    Try
                        Call awinCreateBetterWorsePortfolio(ProjektListe:=myCollection, repChart:=obj, showAbsoluteDiff:=True, isTimeTimeVgl:=False, vglTyp:=1, _
                                                        charttype:=PTpfdk.betterWorseL, bubbleColor:=0, bubbleValueTyp:=PTbubble.strategicFit, showLabels:=True, chartBorderVisible:=True, _
                                                        top:=top, left:=left, width:=width, height:=height)
                    Catch ex As Exception
                        Call MsgBox(ex.Message)
                    End Try

                Else
                    Call MsgBox(" es gibt in diesem Zeitraum keine laufenden / abgeschlossenen Projekte")
                End If


            Catch ex As Exception

            End Try

        Else

            If ShowProjekte.Count = 0 Then
                Call MsgBox("es sind keine Projekte angezeigt")

            Else
                Call MsgBox("im angezeigten Zeitraum " & textZeitraum(showRangeLeft, showRangeRight) & vbLf & _
                            "gibt es keine Projekte")
            End If

        End If



        appInstance.EnableEvents = True
        enableOnUpdate = True
        appInstance.ScreenUpdating = True


    End Sub

    Sub PT0ShowAttentionB(control As IRibbonControl)

        Dim selectionType As Integer
        Dim myCollection As New Collection
        Dim top As Double, left As Double, width As Double, height As Double
        Dim sichtbarerBereich As Excel.Range
        Dim deleteList As New Collection

        Call projektTafelInit()

        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False
        enableOnUpdate = False

        ' hier muss noch geklärt werden, welche Projekte betrachtet werden; es mcht keinen Sinn, 
        'das nur an den TimeFrame zu koppeln, es geht im wesentlichen um aktuell laufende und vergangene Projekte 
        ' Frage : was ist mit bereits beauftragten Projekten, die noch gar nicht begonnen haben, deren Planung aber bereits schlechter als beauftragt ist ? 

        selectionType = PTpsel.lfundab
        myCollection = ShowProjekte.withinTimeFrame(selectionType, showRangeLeft, showRangeRight)

        If myCollection.Count > 0 Then

            With appInstance.ActiveWindow
                sichtbarerBereich = .VisibleRange
                left = CDbl(sichtbarerBereich.Left) + (CDbl(sichtbarerBereich.Width) - 600) / 2
                If left < CDbl(sichtbarerBereich.Left) Then
                    left = CDbl(sichtbarerBereich.Left) + 2
                End If

                top = CDbl(sichtbarerBereich.Top) + (CDbl(sichtbarerBereich.Height) - 450) / 2
                If top < CDbl(sichtbarerBereich.Top) Then
                    top = CDbl(sichtbarerBereich.Top) + 2
                End If

            End With

            width = 600
            height = 450

            Dim obj As Excel.ChartObject = Nothing

            Try
                If myCollection.Count > 0 Then

                    Try
                        Call awinCreateBetterWorsePortfolio(ProjektListe:=myCollection, repChart:=obj, showAbsoluteDiff:=True, isTimeTimeVgl:=False, vglTyp:=1, _
                                                        charttype:=PTpfdk.betterWorseB, bubbleColor:=0, bubbleValueTyp:=PTbubble.strategicFit, showLabels:=True, chartBorderVisible:=True, _
                                                        top:=top, left:=left, width:=width, height:=height)
                    Catch ex As Exception
                        Call MsgBox(ex.Message)
                    End Try

                Else
                    Call MsgBox(" es gibt in diesem Zeitraum keine laufenden bzw. abgeschlossenen Projekte")
                End If


            Catch ex As Exception

            End Try

        Else

            If ShowProjekte.Count = 0 Then
                Call MsgBox("es sind keine Projekte angezeigt")

            Else
                Call MsgBox("im angezeigten Zeitraum " & textZeitraum(showRangeLeft, showRangeRight) & vbLf & _
                            "gibt es keine Projekte")
            End If

        End If


        appInstance.EnableEvents = True
        enableOnUpdate = True
        appInstance.ScreenUpdating = True


    End Sub


    Sub PT0ShowComplexRisiko(control As IRibbonControl)

        Dim selectionType As Integer = -1 ' keine Einschränkung
        Dim myCollection As New Collection
        Dim top As Double, left As Double, width As Double, height As Double
        Dim sichtbarerBereich As Excel.Range

        Call projektTafelInit()

        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False
        enableOnUpdate = False

        myCollection = ShowProjekte.withinTimeFrame(selectionType, showRangeLeft, showRangeRight)


        If myCollection.Count > 0 Then

            With appInstance.ActiveWindow
                sichtbarerBereich = .VisibleRange
                left = CDbl(sichtbarerBereich.Left) + (CDbl(sichtbarerBereich.Width) - 600) / 2
                If left < CDbl(sichtbarerBereich.Left) Then
                    left = CDbl(sichtbarerBereich.Left) + 2
                End If

                top = CDbl(sichtbarerBereich.Top) + (CDbl(sichtbarerBereich.Height) - 450) / 2
                If top < CDbl(sichtbarerBereich.Top) Then
                    top = CDbl(sichtbarerBereich.Top) + 2
                End If

            End With

            width = 600
            height = 450


            Dim obj As Excel.ChartObject = Nothing

            Try
                Call awinCreatePortfolioDiagrams(myCollection, obj, False, PTpfdk.ComplexRisiko, 0, False, True, True, top, left, width, height)
            Catch ex As Exception

            End Try

        Else

            If ShowProjekte.Count = 0 Then
                Call MsgBox("es sind keine Projekte angezeigt")

            Else
                Call MsgBox("im angezeigten Zeitraum " & textZeitraum(showRangeLeft, showRangeRight) & vbLf & _
                            "gibt es keine Projekte")
            End If

        End If


        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True
        enableOnUpdate = True

        Call awinDeSelect()

    End Sub

    Sub PT0ShowZeitRisiko(control As IRibbonControl)

        Dim selectionType As Integer = -1 ' keine Einschränkung
        Dim myCollection As New Collection
        Dim top As Double, left As Double, width As Double, height As Double
        Dim sichtbarerBereich As Excel.Range

        Call projektTafelInit()

        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False
        enableOnUpdate = False

        myCollection = ShowProjekte.withinTimeFrame(selectionType, showRangeLeft, showRangeRight)

        If myCollection.Count > 0 Then

            With appInstance.ActiveWindow
                sichtbarerBereich = .VisibleRange
                left = CDbl(sichtbarerBereich.Left) + (CDbl(sichtbarerBereich.Width) - 600) / 2
                If left < CDbl(sichtbarerBereich.Left) Then
                    left = CDbl(sichtbarerBereich.Left) + 2
                End If

                top = CDbl(sichtbarerBereich.Top) + (CDbl(sichtbarerBereich.Height) - 450) / 2
                If top < CDbl(sichtbarerBereich.Top) Then
                    top = CDbl(sichtbarerBereich.Top) + 2
                End If

            End With

            width = 600
            height = 450


            Dim obj As Excel.ChartObject = Nothing

            Try
                Call awinCreatePortfolioDiagrams(myCollection, obj, False, PTpfdk.ZeitRisiko, 0, False, True, True, top, left, width, height)
            Catch ex As Exception

            End Try

        Else

            If ShowProjekte.Count = 0 Then
                Call MsgBox("es sind keine Projekte angezeigt")

            Else
                Call MsgBox("im angezeigten Zeitraum " & textZeitraum(showRangeLeft, showRangeRight) & vbLf & _
                            "gibt es keine Projekte")
            End If

        End If



        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True
        enableOnUpdate = True

        Call awinDeSelect()

    End Sub

    Sub PTOPTVariantenOptimieren(control As IRibbonControl)


        Dim optmierungsFenster As New frmOptimizeKPI
        Dim returnValue As DialogResult


        Call projektTafelInit()

        appInstance.EnableEvents = False
        enableOnUpdate = False

        returnValue = optmierungsFenster.ShowDialog
        'optmierungsFenster.Show()

        appInstance.EnableEvents = True
        enableOnUpdate = True

    End Sub


    Sub PT0ShowPortfolioBudgetCost(control As IRibbonControl)
        Dim selectionType As Integer = -1 ' keine Einschränkung
        Dim top As Double, left As Double, width As Double, height As Double
        Dim myCollection As New Collection

        Call projektTafelInit()

        appInstance.EnableEvents = False
        enableOnUpdate = False

        myCollection = ShowProjekte.withinTimeFrame(selectionType, showRangeLeft, showRangeRight)

        If myCollection.Count > 0 Then

            Dim sichtbarerBereich As Excel.Range

            height = awinSettings.ChartHoehe2
            width = 450

            With appInstance.ActiveWindow
                sichtbarerBereich = .VisibleRange
                left = CDbl(sichtbarerBereich.Left) + (CDbl(sichtbarerBereich.Width) - width) / 2
                If left < CDbl(sichtbarerBereich.Left) Then
                    left = CDbl(sichtbarerBereich.Left) + 2
                End If

                top = CDbl(sichtbarerBereich.Top) + (CDbl(sichtbarerBereich.Height) - height) / 2
                If top < CDbl(sichtbarerBereich.Top) Then
                    top = CDbl(sichtbarerBereich.Top) + 2
                End If

            End With

            Dim obj As Excel.ChartObject = Nothing
            Call awinCreateBudgetErgebnisDiagramm(obj, top, left, width, height, False, False)

        Else

            If ShowProjekte.Count = 0 Then
                Call MsgBox("es sind keine Projekte angezeigt")

            Else
                If showRangeRight - showRangeLeft < 6 Then
                    Call MsgBox(" Bitte wählen Sie zuerst einen Zeitraum aus !")
                Else
                    Call MsgBox("im angezeigten Zeitraum " & textZeitraum(showRangeLeft, showRangeRight) & vbLf & _
                                "gibt es keine Projekte ")
                End If
            End If

        End If


        appInstance.EnableEvents = True
        enableOnUpdate = True
    End Sub


    Sub PT0ShowPortfolioErgebnis(control As IRibbonControl)
        Dim selectionType As Integer = -1 ' keine Einschränkung
        Dim top As Double, left As Double, width As Double, height As Double
        Dim myCollection As New Collection

        Call projektTafelInit()

        appInstance.EnableEvents = False
        enableOnUpdate = False

        myCollection = ShowProjekte.withinTimeFrame(selectionType, showRangeLeft, showRangeRight)

        If myCollection.Count > 0 Then

            Dim sichtbarerBereich As Excel.Range

            height = awinSettings.ChartHoehe2
            width = 450

            With appInstance.ActiveWindow
                sichtbarerBereich = .VisibleRange
                left = CDbl(sichtbarerBereich.Left) + (CDbl(sichtbarerBereich.Width) - width) / 2
                If left < CDbl(sichtbarerBereich.Left) Then
                    left = CDbl(sichtbarerBereich.Left) + 2
                End If

                top = CDbl(sichtbarerBereich.Top) + (CDbl(sichtbarerBereich.Height) - height) / 2
                If top < CDbl(sichtbarerBereich.Top) Then
                    top = CDbl(sichtbarerBereich.Top) + 2
                End If

            End With



            Dim obj As Excel.ChartObject = Nothing
            Call awinCreateErgebnisDiagramm(obj, top, left, width, height, False, False)


        Else

            If ShowProjekte.Count = 0 Then
                Call MsgBox("es sind keine Projekte angezeigt")

            Else
                If showRangeRight - showRangeLeft < 6 Then
                    Call MsgBox(" Bitte wählen Sie zuerst einen Zeitraum aus !")
                Else
                    Call MsgBox("im angezeigten Zeitraum " & textZeitraum(showRangeLeft, showRangeRight) & vbLf & _
                                "gibt es keine Projekte ")
                End If
            End If


        End If



        appInstance.EnableEvents = True
        enableOnUpdate = True
    End Sub



    Sub Tom2G2M5M1B3ShowStatus(control As IRibbonControl)

        Dim selectionType As Integer = -1 ' keine Einschränkung
        Dim heute As Integer = getColumnOfDate(Date.Now)
        Dim myCollection As New Collection

        Call projektTafelInit()

        appInstance.EnableEvents = False
        enableOnUpdate = False

        myCollection = ShowProjekte.withinTimeFrame(selectionType, heute, heute)

        If myCollection.Count > 0 Then

            Dim nummerieren As Boolean = False
            Call awinZeichneStatus(nummerieren)

        Else

            Call MsgBox("es gibt keine aktuell laufenden Projekte")

        End If



        appInstance.EnableEvents = True
        enableOnUpdate = True

    End Sub


    ''' <summary>
    ''' Charakteristik Projekt-Ergebnis
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G2M1B7Ergebnis(control As IRibbonControl)

        Dim singleShp As Excel.Shape
        'Dim SID As String

        Dim awinSelection As Excel.ShapeRange

        Call projektTafelInit()

        enableOnUpdate = False
        appInstance.EnableEvents = False


        Try
            'awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            If awinSelection.Count = 1 Then
                ' jetzt die Aktion durchführen ...
                singleShp = awinSelection.Item(1)
                Dim formerSU As Boolean = appInstance.ScreenUpdating
                Dim formerEE As Boolean = appInstance.EnableEvents
                appInstance.ScreenUpdating = False
                appInstance.EnableEvents = False

                Dim dummyObj As Excel.ChartObject = Nothing
                Dim hproj As clsProjekt
                Try
                    hproj = ShowProjekte.getProject(singleShp.Name)
                    Call createProjektErgebnisCharakteristik2(hproj, dummyObj, PThis.current)
                Catch ex As Exception
                    Call MsgBox("Name nicht gefunden : " & singleShp.Name)
                End Try

                appInstance.ScreenUpdating = formerSU
                appInstance.EnableEvents = formerEE
            Else
                Call MsgBox("bitte nur ein Projekt selektieren")
                'For Each singleShp In awinSelection
                '    With singleShp
                '        If .AutoShapeType = MsoAutoShapeType.msoShapeRoundedRectangle Then
                '            nrSelPshp = nrSelPshp + 1
                '            SID = .ID.ToString
                '        End If
                '    End With
                'Next
            End If
        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If

        enableOnUpdate = True
        appInstance.EnableEvents = True



    End Sub

    ''' <summary>
    ''' Vergleichen mit Beauftragung / Freigabe
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G2M2B1Auftrag(control As IRibbonControl)

        Dim singleShp As Excel.Shape
        'Dim SID As String

        Dim awinSelection As Excel.ShapeRange

        Call projektTafelInit()

        enableOnUpdate = False
        appInstance.EnableEvents = False

        Try
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            If awinSelection.Count = 1 Then
                ' jetzt die Aktion durchführen ...
                singleShp = awinSelection.Item(1)
                Dim hproj As clsProjekt

                Try
                    hproj = ShowProjekte.getProject(singleShp.Name)
                    Dim cproj As New clsProjekt
                    Dim top As Double = singleShp.Top + boxHeight + 2
                    Dim left As Double = singleShp.Left - boxWidth
                    If left <= 0 Then
                        left = 1
                    End If
                    Call awinCompareProject(hproj, cproj, 0, top, left)

                Catch ex As Exception
                    Call MsgBox("Fehler bei Beauftragung " & vbLf & ex.Message)
                End Try


                'Call awinCompareProject(pname1:=singleShp.Name, pname2:=" ", compareType:=0)

            Else
                Call MsgBox("bitte nur ein Projekt selektieren")
                'For Each singleShp In awinSelection
                '    With singleShp
                '        If .AutoShapeType = MsoAutoShapeType.msoShapeRoundedRectangle Then
                '            nrSelPshp = nrSelPshp + 1
                '            SID = .ID.ToString
                '        End If
                '    End With
                'Next
            End If
        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If

        enableOnUpdate = True
        appInstance.EnableEvents = True

    End Sub

    ''' <summary>
    ''' die Phasen zweier Projekte vergleichen  - Darstellung in einem Chart
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G3M1B1PhasenVgl(control As IRibbonControl)

        Dim singleShp1 As Excel.Shape, singleShp2 As Excel.Shape
        'Dim SID As String
        Dim hproj As clsProjekt, cproj As clsProjekt
        Dim top As Double, left As Double, width As Double, height As Double
        Dim scale As Double

        Dim awinSelection As Excel.ShapeRange

        Call projektTafelInit()

        enableOnUpdate = False

        Try
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            If awinSelection.Count = 2 Then
                ' jetzt die Aktion durchführen ...
                singleShp1 = awinSelection.Item(1)
                singleShp2 = awinSelection.Item(2)

                Try
                    hproj = ShowProjekte.getProject(singleShp1.Name)
                    cproj = ShowProjekte.getProject(singleShp2.Name)
                Catch ex As Exception
                    Call MsgBox("Projekt nicht gefunden ...")
                    enableOnUpdate = True
                    Exit Sub
                End Try


                top = singleShp1.Top + boxHeight + 2
                left = singleShp1.Left - 5
                If left <= 0 Then
                    left = 1
                End If

                height = 380

                width = System.Math.Max(hproj.anzahlRasterElemente * boxWidth + 7, cproj.anzahlRasterElemente * boxWidth + 7)
                scale = System.Math.Max(hproj.dauerInDays, cproj.dauerInDays)
                'width = hproj1.Dauer * boxWidth + 7
                'scale = hproj1.Dauer

                Dim repObj As Excel.ChartObject
                appInstance.EnableEvents = False
                appInstance.ScreenUpdating = False

                repObj = Nothing
                Dim htitel As String = hproj.name
                Dim ctitel As String = cproj.name
                Call awinCompareProjectPhases(hproj, htitel, cproj, ctitel, 3, repObj)


                appInstance.ScreenUpdating = True

            Else
                Call MsgBox("bitte zwei Projekte selektieren")
                'For Each singleShp In awinSelection
                '    With singleShp
                '        If .AutoShapeType = MsoAutoShapeType.msoShapeRoundedRectangle Then
                '            nrSelPshp = nrSelPshp + 1
                '            SID = .ID.ToString
                '        End If
                '    End With
                'Next
            End If
        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If

        enableOnUpdate = True
        appInstance.EnableEvents = True

    End Sub

    ''' <summary>
    ''' die Phasen zweier Projekte vergleichen  - Darstellung in zwei Charts
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G3M1B2PhasenVgl(control As IRibbonControl)

        Dim singleShp1 As Excel.Shape, singleShp2 As Excel.Shape
        Dim hproj As New clsProjekt, cproj As clsProjekt
        Dim top As Double, left As Double, width As Double, height As Double
        Dim scale As Double
        Dim noColorCollection As New Collection

        Dim awinSelection As Excel.ShapeRange

        Call projektTafelInit()

        enableOnUpdate = False

        Try
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then


            If awinSelection.Count = 1 Then

                Dim vproj As clsProjektvorlage
                ' jetzt die Aktion durchführen ...
                singleShp1 = awinSelection.Item(1)

                Try
                    hproj = ShowProjekte.getProject(singleShp1.Name)
                    vproj = Projektvorlagen.getProject(hproj.VorlagenName)
                    If IsNothing(vproj) Then
                        Call MsgBox("Vorlage" & hproj.VorlagenName & " nicht gefunden ...")
                        enableOnUpdate = True
                        Exit Sub
                    End If
                    cproj = New clsProjekt
                    vproj.copyTo(cproj)
                    cproj.startDate = hproj.startDate

                Catch ex As Exception
                    Call MsgBox("Vorlage" & hproj.VorlagenName & " nicht gefunden ...")
                    enableOnUpdate = True
                    Exit Sub
                End Try


                top = singleShp1.Top + boxHeight + 2
                left = singleShp1.Left - 5
                If left <= 0 Then
                    left = 5
                End If

                height = 380
                width = System.Math.Max(hproj.dauerInDays / 365 * 12 * boxWidth + 7, cproj.dauerInDays / 365 * 12 * boxWidth + 7)
                scale = System.Math.Max(hproj.dauerInDays, cproj.dauerInDays)

                Dim repObj As Excel.ChartObject
                appInstance.EnableEvents = False
                appInstance.ScreenUpdating = False


                noColorCollection = getPhasenUnterschiede(hproj, cproj)

                repObj = Nothing
                Call createPhasesBalken(noColorCollection, hproj, repObj, scale, top, left, height, width, PThis.current)

                With repObj
                    top = .Top + .Height + 3
                End With


                repObj = Nothing
                Call createPhasesBalken(noColorCollection, cproj, repObj, scale, top, left, height, width, PThis.vorlage)
                appInstance.ScreenUpdating = True

            ElseIf awinSelection.Count = 2 Then
                ' jetzt die Aktion durchführen ...
                singleShp1 = awinSelection.Item(1)
                singleShp2 = awinSelection.Item(2)

                Try
                    hproj = ShowProjekte.getProject(singleShp1.Name)
                    cproj = ShowProjekte.getProject(singleShp2.Name)
                Catch ex As Exception
                    Call MsgBox("Projekt nicht gefunden ...")
                    enableOnUpdate = True
                    Exit Sub
                End Try


                top = singleShp1.Top + boxHeight + 2
                left = singleShp1.Left - 5
                If left <= 0 Then
                    left = 5
                End If

                height = 380
                width = System.Math.Max(hproj.dauerInDays / 365 * 12 * boxWidth + 7, cproj.dauerInDays / 365 * 12 * boxWidth + 7)
                scale = System.Math.Max(hproj.dauerInDays, cproj.dauerInDays)

                Dim repObj As Excel.ChartObject
                appInstance.EnableEvents = False
                appInstance.ScreenUpdating = False

                noColorCollection = getPhasenUnterschiede(hproj, cproj)

                repObj = Nothing
                Call createPhasesBalken(noColorCollection, hproj, repObj, scale, top, left, height, width, PThis.current)

                With repObj
                    top = .Top + .Height + 3
                End With


                repObj = Nothing
                Call createPhasesBalken(noColorCollection, cproj, repObj, scale, top, left, height, width, PThis.current)
                appInstance.ScreenUpdating = True
                'Call awinCompareProjectPhases(name1:=singleShp1.Name, _
                '                              name2:=singleShp2.Name, _
                '                              compareType:=3)
            Else
                Call MsgBox("bitte zwei Projekte selektieren")

            End If
        Else
            Call MsgBox("ein Projekt selektieren, um mit Vorlage zu vergleichen" & vbLf & _
                        " oder zwei Projekte für den Vergleich untereinander")
        End If

        enableOnUpdate = True
        appInstance.EnableEvents = True

    End Sub

    Sub PT3G1B2PhasenVgl(control As IRibbonControl)

        Dim request As New Request(awinSettings.databaseName, dbUsername, dbPasswort)
        Dim singleShp1 As Excel.Shape
        Dim hproj As clsProjekt, cproj As clsProjekt
        Dim top As Double, left As Double, width As Double, height As Double
        Dim scale As Double
        Dim noColorCollection As New Collection
        Dim vglName As String = " "
        Dim pName As String = "", variantName As String
        Dim awinSelection As Excel.ShapeRange

        Call projektTafelInit()

        enableOnUpdate = False

        Try
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try


        If request.pingMongoDb() Then

            If Not awinSelection Is Nothing Then

                If awinSelection.Count = 1 Then

                    Try
                        Dim lastElem As Integer

                        ' jetzt die Aktion durchführen ...
                        singleShp1 = awinSelection.Item(1)


                        Try
                            hproj = ShowProjekte.getProject(singleShp1.Name)
                        Catch ex As Exception
                            Call MsgBox("Projekt nicht gefunden ...")
                            enableOnUpdate = True
                            Exit Sub
                        End Try

                        ' jetzt ggf die Projekt-Historie aufbauen

                        If Not projekthistorie Is Nothing Then
                            If projekthistorie.Count > 0 Then
                                vglName = projekthistorie.First.getShapeText
                            End If
                        Else
                            projekthistorie = New clsProjektHistorie
                        End If

                        With hproj
                            pName = .name
                            variantName = .variantName
                        End With

                        If vglName <> hproj.getShapeText Then

                            ' projekthistorie muss nur dann neu bestimmt werden, wenn sie nicht bereits für dieses Projekt geholt wurde
                            projekthistorie.liste = request.retrieveProjectHistoryFromDB(projectname:=pName, variantName:=variantName, _
                                                                                storedEarliest:=StartofCalendar, storedLatest:=Date.Now)
                            projekthistorie.Add(Date.Now, hproj)
                            lastElem = projekthistorie.Count - 1


                        Else
                            ' der aktuelle Stand hproj muss hinzugefügt werden 
                            lastElem = projekthistorie.Count - 1
                            projekthistorie.RemoveAt(lastElem)
                            projekthistorie.Add(Date.Now, hproj)
                        End If


                        If projekthistorie.Count <= 1 Then

                            Call MsgBox(" es gibt zu diesem Projekt noch keine Historie")

                        Else

                            cproj = projekthistorie.ElementAt(lastElem - 1)

                            top = singleShp1.Top + boxHeight + 2
                            left = singleShp1.Left - 5
                            If left <= 0 Then
                                left = 5
                            End If

                            height = 380
                            width = System.Math.Max(hproj.dauerInDays / 365 * 12 * boxWidth + 7, cproj.dauerInDays / 365 * 12 * boxWidth + 7)
                            scale = System.Math.Max(hproj.dauerInDays, cproj.dauerInDays)

                            Dim repObj As Excel.ChartObject
                            appInstance.EnableEvents = False
                            appInstance.ScreenUpdating = False

                            noColorCollection = getPhasenUnterschiede(hproj, cproj)

                            repObj = Nothing
                            Call createPhasesBalken(noColorCollection, hproj, repObj, scale, top, left, height, width, PThis.current)

                            With repObj
                                top = .Top + .Height + 3
                            End With


                            repObj = Nothing
                            Call createPhasesBalken(noColorCollection, cproj, repObj, scale, top, left, height, width, PThis.letzterStand)

                            appInstance.ScreenUpdating = True

                        End If
                    Catch ex As Exception

                        Call MsgBox("es gibt keine Historie zu " & pName)

                    End Try


                Else
                    Call MsgBox("bitte nur ein Projekt selektieren")

                End If
            Else
                Call MsgBox("ein Projekt selektieren, um es mit seinem letzten Stand zu vergleichen")
            End If
        Else
            Call MsgBox("Datenbank-Verbindung ist unterbrochen!")
            projekthistorie.clear()
        End If

        enableOnUpdate = True
        appInstance.EnableEvents = True

    End Sub

    ''' <summary>
    ''' vergleicht die Phasen Termine des aktuellen Projektes mit der Beauftragung
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub PT3G1B3PhasenVgl(control As IRibbonControl)

        Dim request As New Request(awinSettings.databaseName, dbUsername, dbPasswort)
        Dim singleShp1 As Excel.Shape
        Dim hproj As clsProjekt, cproj As clsProjekt
        Dim top As Double, left As Double, width As Double, height As Double
        Dim scale As Double
        Dim noColorCollection As New Collection
        Dim vglName As String = " "
        Dim pName As String, variantName As String

        Dim awinSelection As Excel.ShapeRange

        Call projektTafelInit()

        enableOnUpdate = False

        Try
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If request.pingMongoDb() Then

            If Not awinSelection Is Nothing Then

                If awinSelection.Count = 1 Then

                    Dim lastElem As Integer

                    ' jetzt die Aktion durchführen ...
                    singleShp1 = awinSelection.Item(1)


                    Try
                        hproj = ShowProjekte.getProject(singleShp1.Name)
                    Catch ex As Exception
                        Call MsgBox("Projekt nicht gefunden ...")
                        enableOnUpdate = True
                        Exit Sub
                    End Try

                    ' jetzt ggf die Projekt-Historie aufbauen

                    If Not projekthistorie Is Nothing Then
                        If projekthistorie.Count > 0 Then
                            vglName = projekthistorie.First.getShapeText
                        End If
                    Else
                        projekthistorie = New clsProjektHistorie
                    End If

                    With hproj
                        pName = .name
                        variantName = .variantName
                    End With

                    If vglName <> hproj.getShapeText Then

                        ' projekthistorie muss nur dann neu bestimmt werden, wenn sie nicht bereits für dieses Projekt geholt wurde
                        projekthistorie.liste = request.retrieveProjectHistoryFromDB(projectname:=pName, variantName:=variantName, _
                                                                            storedEarliest:=StartofCalendar, storedLatest:=Date.Now)
                        projekthistorie.Add(Date.Now, hproj)
                        lastElem = projekthistorie.Count - 1


                    Else
                        ' der aktuelle Stand hproj muss hinzugefügt werden 
                        lastElem = projekthistorie.Count - 1
                        projekthistorie.RemoveAt(lastElem)
                        projekthistorie.Add(Date.Now, hproj)
                    End If


                    If projekthistorie.Count = 1 Then

                        Call MsgBox(" es gibt zu diesem Projekt noch keine Historie")

                    Else


                        Try
                            cproj = projekthistorie.beauftragung
                            top = singleShp1.Top + boxHeight + 2
                            left = singleShp1.Left - 5
                            If left <= 0 Then
                                left = 5
                            End If

                            height = 380
                            width = System.Math.Max(hproj.dauerInDays / 365 * 12 * boxWidth + 7, cproj.dauerInDays / 365 * 12 * boxWidth + 7)
                            scale = System.Math.Max(hproj.dauerInDays, cproj.dauerInDays)

                            Dim repObj As Excel.ChartObject
                            appInstance.EnableEvents = False
                            appInstance.ScreenUpdating = False

                            noColorCollection = getPhasenUnterschiede(hproj, cproj)

                            repObj = Nothing
                            Call createPhasesBalken(noColorCollection, hproj, repObj, scale, top, left, height, width, PThis.current)

                            With repObj
                                top = .Top + .Height + 3
                            End With


                            repObj = Nothing
                            Call createPhasesBalken(noColorCollection, cproj, repObj, scale, top, left, height, width, PThis.beauftragung)

                        Catch ex As Exception

                            Call MsgBox("es ist kein Beauftragungs-Stand vorhanden")

                        End Try


                    End If

                Else
                    Call MsgBox("bitte nur ein Projekt selektieren")

                End If
            Else
                Call MsgBox("ein Projekt selektieren, um es mit seiner Beauftragung zu vergleichen")
            End If

        Else
            Call MsgBox("Datenbank-Verbindung ist unterbrochen!")
        End If
        enableOnUpdate = True
        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True

    End Sub

    Sub Tom2G3M1B2ResourceVgl(control As IRibbonControl)

        Dim singleShp1 As Excel.Shape, singleShp2 As Excel.Shape
        'Dim SID As String

        Dim awinSelection As Excel.ShapeRange

        Call projektTafelInit()

        enableOnUpdate = False
        appInstance.EnableEvents = False

        Try
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            If awinSelection.Count = 2 Then
                ' jetzt die Aktion durchführen ...
                singleShp1 = awinSelection.Item(1)
                singleShp2 = awinSelection.Item(2)

                Dim hproj As clsProjekt
                Dim cproj As clsProjekt
                Try
                    hproj = ShowProjekte.getProject(singleShp1.Name)
                    cproj = ShowProjekte.getProject(singleShp2.Name)
                    Dim top As Double = singleShp1.Top + boxHeight + 2
                    Dim left As Double = singleShp1.Left - boxWidth
                    If left <= 0 Then
                        left = 1
                    End If
                    Call awinCompareProject(hproj, cproj, 3, top, left)
                Catch ex As Exception
                    Call MsgBox("Fehler bei Compare" & vbLf & ex.Message)
                End Try

            Else
                Call MsgBox("bitte zwei Projekte selektieren")
                'For Each singleShp In awinSelection
                '    With singleShp
                '        If .AutoShapeType = MsoAutoShapeType.msoShapeRoundedRectangle Then
                '            nrSelPshp = nrSelPshp + 1
                '            SID = .ID.ToString
                '        End If
                '    End With
                'Next
            End If
        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If


        enableOnUpdate = True
        appInstance.EnableEvents = True

    End Sub

    Sub awinShowTrendSR(control As IRibbonControl)

        Dim hproj As clsProjekt
        Dim pName As String, variantName As String
        Dim request As New Request(awinSettings.databaseName, dbUsername, dbPasswort)
        Dim singleShp As Excel.Shape
        Dim showCharacteristics As New frmShowProjCharacteristics
        'Dim returnValue As DialogResult
        Dim awinSelection As Excel.ShapeRange
        Dim top As Double, left As Double, height As Double, width As Double
        Dim vglName As String = " "

        Call projektTafelInit()

        enableOnUpdate = False
        appInstance.ScreenUpdating = False


        Try
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If request.pingMongoDb() Then

            If Not awinSelection Is Nothing Then

                If awinSelection.Count = 1 And isProjectType(kindOfShape(awinSelection.Item(1))) Then
                    ' jetzt die Aktion durchführen ...
                    singleShp = awinSelection.Item(1)


                    hproj = ShowProjekte.getProject(singleShp.Name)
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

                    Dim nrSnapshots As Integer = projekthistorie.Count

                    If nrSnapshots > 0 Then
                        With singleShp
                            top = .Top + boxHeight + 2
                            left = .Left - 3
                        End With
                        width = System.Math.Max(nrSnapshots * boxWidth * 0.65, 300)

                        height = 16 * boxHeight
                        Dim repObj As Excel.ChartObject = Nothing
                        Call createTrendSfit(repObj, top, left, height, width)

                    Else
                        Call MsgBox("es gibt noch keine Projekt-Historie zu " & pName)
                    End If




                Else
                    Call MsgBox("bitte nur ein Projekt selektieren")
                    'For Each singleShp In awinSelection
                    '    With singleShp
                    '        If .AutoShapeType = MsoAutoShapeType.msoShapeRoundedRectangle Then
                    '            nrSelPshp = nrSelPshp + 1
                    '            SID = .ID.ToString
                    '        End If
                    '    End With
                    'Next
                End If
            Else
                Call MsgBox("vorher Projekt selektieren ...")
            End If
        Else
            Call MsgBox("Datenbank-Verbindung ist unterbrochen!")
        End If

        enableOnUpdate = True
        appInstance.ScreenUpdating = True




    End Sub


    Sub awinShowTrendKPI(control As IRibbonControl)
        Dim hproj As clsProjekt
        Dim pName As String, variantName As String
        Dim request As New Request(awinSettings.databaseName, dbUsername, dbPasswort)
        Dim singleShp As Excel.Shape
        Dim showCharacteristics As New frmShowProjCharacteristics
        'Dim returnValue As DialogResult
        Dim awinSelection As Excel.ShapeRange
        Dim top As Double, left As Double, height As Double, width As Double
        Dim vglName As String = " "


        Call projektTafelInit()

        enableOnUpdate = False
        appInstance.ScreenUpdating = False


        Try
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            If awinSelection.Count = 1 And isProjectType(kindOfShape(awinSelection.Item(1))) Then
                ' jetzt die Aktion durchführen ...
                singleShp = awinSelection.Item(1)


                hproj = ShowProjekte.getProject(singleShp.Name)
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
                    If request.pingMongoDb() Then
                        ' projekthistorie muss nur dann neu bestimmt werden, wenn sie nicht bereits für dieses Projekt geholt wurde
                        projekthistorie.liste = request.retrieveProjectHistoryFromDB(projectname:=pName, variantName:=variantName, _
                                                                            storedEarliest:=StartofCalendar, storedLatest:=Date.Now)
                        projekthistorie.Add(Date.Now, hproj)
                    Else
                        Call MsgBox("Datenbank-Verbindung ist unterbrochen !")
                        projekthistorie.clear()
                    End If

                Else
                    ' der aktuelle Stand hproj muss hinzugefügt werden 
                    Dim lastElem As Integer = projekthistorie.Count - 1
                    projekthistorie.RemoveAt(lastElem)
                    projekthistorie.Add(Date.Now, hproj)
                End If

                Dim nrSnapshots As Integer = projekthistorie.Count

                If nrSnapshots > 0 Then
                    With singleShp
                        top = .Top + boxHeight + 2
                        left = .Left - 3
                    End With
                    width = System.Math.Max(nrSnapshots * boxWidth * 0.65, 300)

                    height = 16 * boxHeight
                    Dim repObj As Excel.ChartObject = Nothing
                    Call createTrendKPI(repObj, top, left, height, width)

                Else
                    Call MsgBox("es gibt noch keine Projekt-Historie zu " & pName)
                End If

            Else
                Call MsgBox("bitte nur ein Projekt selektieren")
                'For Each singleShp In awinSelection
                '    With singleShp
                '        If .AutoShapeType = MsoAutoShapeType.msoShapeRoundedRectangle Then
                '            nrSelPshp = nrSelPshp + 1
                '            SID = .ID.ToString
                '        End If
                '    End With
                'Next
            End If
        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If

        enableOnUpdate = True
        appInstance.ScreenUpdating = True


    End Sub


    Sub awinShowTimeMachine(control As IRibbonControl)
        Dim hproj As clsProjekt
        Dim pName As String, variantName As String
        Dim vglName As String = " "
        Dim request As New Request(awinSettings.databaseName, dbUsername, dbPasswort)
        Dim singleShp As Excel.Shape
        Dim showCharacteristics As New frmShowProjCharacteristics
        'Dim returnValue As DialogResult
        Dim awinSelection As Excel.ShapeRange
        Dim grueneAmpel As String = awinPath & "gruen.gif"
        Dim gelbeAmpel As String = awinPath & "gelb.gif"
        Dim roteAmpel As String = awinPath & "rot.gif"
        Dim graueAmpel As String = awinPath & "grau.gif"

        If timeMachineIsOn Then
            Call MsgBox("bitte erst Time Machine beenden ...")
            Exit Sub
        End If

        Call projektTafelInit()

        enableOnUpdate = False
        appInstance.EnableEvents = True


        Try
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then


            If awinSelection.Count = 1 And isProjectType(kindOfShape(awinSelection.Item(1))) Then
                ' jetzt die Aktion durchführen ...
                singleShp = awinSelection.Item(1)
                hproj = ShowProjekte.getProject(singleShp.Name)
                With hproj
                    pName = .name
                    variantName = .variantName
                    'Try
                    '    variantName = .variantName.Trim
                    'Catch ex As Exception
                    '    variantName = ""
                    'End Try

                End With

                If Not projekthistorie Is Nothing Then
                    If projekthistorie.Count > 0 Then
                        vglName = projekthistorie.First.getShapeText
                    End If

                Else
                    projekthistorie = New clsProjektHistorie
                End If

                If vglName <> hproj.getShapeText Then

                    If request.pingMongoDb() Then
                        ' projekthistorie muss nur dann neu geladen werden, wenn sie nicht bereits für dieses Projekt geholt wurde
                        projekthistorie.liste = request.retrieveProjectHistoryFromDB(projectname:=pName, variantName:=variantName, _
                                                                            storedEarliest:=StartofCalendar, storedLatest:=Date.Now)
                        If projekthistorie.Count <> 0 Then

                            projekthistorie.Add(Date.Now, hproj)

                        End If

                    Else
                        Call MsgBox("Datenbank-Verbindung ist unterbrochen")
                        projekthistorie.clear()
                    End If

                Else
                    ' der aktuelle Stand hproj muss hinzugefügt werden 
                    Dim lastElem As Integer = projekthistorie.Count - 1
                    projekthistorie.RemoveAt(lastElem)
                    projekthistorie.Add(Date.Now, hproj)
                End If


                Dim nrSnapshots As Integer = projekthistorie.Count

                If nrSnapshots > 0 Then

                    With showCharacteristics

                        .Text = "Historie für Projekt " & pName.Trim & vbLf & _
                                "( " & projekthistorie.getZeitraum & " )"
                        .timeSlider.Minimum = 0
                        .timeSlider.Maximum = nrSnapshots - 1

                        '.ampelErlaeuterung.Text = kvp.Value.ampelErlaeuterung

                        'If kvp.Value.ampelStatus = 1 Then
                        '    .ampelPicture.LoadAsync(grueneAmpel)
                        'ElseIf kvp.Value.ampelStatus = 2 Then
                        '    .ampelPicture.LoadAsync(gelbeAmpel)
                        'ElseIf kvp.Value.ampelStatus = 3 Then
                        '    .ampelPicture.LoadAsync(roteAmpel)
                        'Else
                        '    .ampelPicture.LoadAsync(graueAmpel)
                        'End If

                        '.snapshotDate.Text = kvp.Value.timeStamp.ToString
                        ' das ist ja der aktuelle Stand ..
                        .snapshotDate.Text = "Aktueller Stand"
                        ' Designer 
                        'Dim zE As String = "(" & awinSettings.zeitEinheit & ")"
                        '.engpass1.Text = "Designer:          " & kvp.Value.getRessourcenBedarf(3).Sum.ToString("###.#") & zE
                        '.engpass2.Text = "Personalkosten: " & kvp.Value.getAllPersonalKosten.Sum.ToString("###.#") & " (T€)"
                        '.engpass3.Text = "Sonstige Kosten:   " & kvp.Value.getGesamtAndereKosten.Sum.ToString("###.#") & " (T€)"


                    End With


                    ' jetzt wird das Form aufgerufen ... 

                    'returnValue = showCharacteristics.ShowDialog
                    showCharacteristics.Show()

                Else
                    Call MsgBox("es gibt noch keine Planungs-Historie")
                End If

            Else
                Call MsgBox("bitte nur ein Projekt selektieren")
                'For Each singleShp In awinSelection
                '    With singleShp
                '        If .AutoShapeType = MsoAutoShapeType.msoShapeRoundedRectangle Then
                '            nrSelPshp = nrSelPshp + 1
                '            SID = .ID.ToString
                '        End If
                '    End With
                'Next
            End If
        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If

        enableOnUpdate = True


    End Sub



    ' 

    ''' <summary>
    ''' aktuelle Konstellation wird dokumentiert
    ''' Report-Vorlage wird im Formular 'Auswählen der Report-Vorlage' ausgewählt
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub awinAllprojectsReport(control As IRibbonControl)

        Dim getReportVorlage As New frmSelectPPTTempl
        Dim returnValue As DialogResult

        getReportVorlage.calledfrom = "Portfolio1"

        Call projektTafelInit()

        enableOnUpdate = False
        appInstance.ScreenUpdating = False
        If showRangeRight - showRangeLeft > 6 Then

            If ShowProjekte.Count > 0 Then

                ' Formular zum Auswählen der Report-Vorlage wird aufgerufen

                returnValue = getReportVorlage.ShowDialog

            Else
                Call MsgBox("Es sind keine Projekte geladen!")
            End If
        Else
            Call MsgBox("Bitte wählen Sie den Zeitraum aus, für den der Report erstellt werden soll!")
        End If

        appInstance.ScreenUpdating = True
        enableOnUpdate = True


        ' das ist die alte Variante
        '
        'enableOnUpdate = False
        'appInstance.EnableEvents = False
        'appInstance.ScreenUpdating = False

        'Try
        '    Call createPPTSlidesFromConstellation()
        'Catch ex As Exception
        '    Call MsgBox(ex.Message)
        'End Try


        'appInstance.EnableEvents = True
        'appInstance.ScreenUpdating = True
        'enableOnUpdate = True

    End Sub
    Sub PTShowVersions(control As IRibbonControl)

        'Ermittlung der installierten Windows- und der Excelversion
        Call MsgBox("Betriebssystem: " & appInstance.OperatingSystem & Chr(10) & _
        "Excel-Version: " & appInstance.Version, vbInformation, "Info")
        'Call MsgBox("Betriebssystem: " & appInstance.OperatingSystem & Chr(10) & _
        '"Excel-Version: " & My.Settings.ExcelVersion, vbInformation, "Info")
    End Sub


    Sub PTTestFunktion1(control As IRibbonControl)

        Call MsgBox("Enable Events ist " & appInstance.EnableEvents.ToString)
        appInstance.EnableEvents = True


    End Sub

    Sub PTTestFunktion2(control As IRibbonControl)

        Dim hproj As clsProjekt
        Dim request As New Request(awinSettings.databaseName, dbUsername, dbPasswort)
        Dim singleShp As Excel.Shape
        Dim tstCollection As SortedList(Of Date, String)
        Dim anzElements As Integer

        Dim awinSelection As Excel.ShapeRange
        Dim projektHistorien As New clsProjektDBInfos
        Dim todoListe As New clsProjektDBInfos
        Dim i As Integer


        Dim schluessel As String = ""

        Call projektTafelInit()

        enableOnUpdate = False
        appInstance.EnableEvents = True



        Try
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            If awinSelection.Count > 1 Then
                anzElements = awinSelection.Count

                For i = 1 To anzElements

                    singleShp = awinSelection.Item(i)
                    hproj = ShowProjekte.getProject(singleShp.Name)

                    If i = 1 Then
                        schluessel = calcProjektKey(hproj)
                    End If

                    If request.pingMongoDb() Then
                        ' projekthistorie muss nur dann neu geladen werden, wenn sie nicht bereits für dieses Projekt geholt wurde
                        projekthistorie.liste = request.retrieveProjectHistoryFromDB(projectname:=hproj.name, variantName:=hproj.variantName, _
                                                                           storedEarliest:=StartofCalendar, storedLatest:=Date.Now)
                    Else
                        Call MsgBox("Datenbank-Verbindung ist unterbrochen")
                        projekthistorie.clear()
                    End If

                    If projekthistorie.Count > 0 Then
                        ' Aufbau der Listen 
                        projektHistorien.Add(projekthistorie)


                    End If

                Next
            End If
        End If

        Dim ts As Date


        tstCollection = projektHistorien.getTimeStamps(schluessel)
        anzElements = tstCollection.Count

        For i = 1 To anzElements
            ts = tstCollection.ElementAt(0).Key
            projektHistorien.Remove(schluessel, ts)
            todoListe.Add(schluessel, ts)
        Next


        enableOnUpdate = True


    End Sub

    ''' <summary>
    ''' testet, ob die Hierarchien in den geladenen Projekten alle stimmig sind 
    ''' das heißt, verweisen die Indices tatsächlich auf die richtigen Phasen bzw Meilensteine  
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub PTTestFunktion3(control As IRibbonControl)

        Dim hproj As clsProjekt
        Dim allesInOrdnung As Boolean = True
        Dim anzElements As Integer
        Dim curNode As clsHierarchyNode
        Dim parentNode As clsHierarchyNode
        Dim childNode As clsHierarchyNode
        Dim parentID As String
        Dim curID As String
        Dim childID As String
        Dim elemID As String
        Dim elemName As String
        Dim lfdNr As Integer
        Dim isMilestone As Boolean
        Dim cphase As clsPhase
        Dim cphase2 As clsPhase
        Dim cMilestone As clsMeilenstein
        Dim logMessage As String = ""
        Dim atleastOne As Boolean = False


        Call projektTafelInit()

        enableOnUpdate = False
        appInstance.EnableEvents = True

        ' Testreihe 1: ausgehend von der Hierarchie alle Projekte und Varianten 

        For Each kvp As KeyValuePair(Of String, clsProjekt) In AlleProjekte.liste

            hproj = kvp.Value


            ' zuerst wird ausgehend von der Hierarchie gecheckt 
            anzElements = hproj.hierarchy.count

            For ix As Integer = 1 To anzElements

                curID = hproj.hierarchy.getIDAtIndex(ix)
                elemName = elemNameOfElemID(curID)
                lfdNr = lfdNrOfElemID(curID)

                curNode = hproj.hierarchy.nodeItem(ix)

                If curID.StartsWith("1§") Then
                    isMilestone = True
                ElseIf curID.StartsWith("0§") Then
                    isMilestone = False
                Else
                    logMessage = logMessage & vbLf & kvp.Value.getShapeText & ": Node kann nicht identifiziert werden .." & curID
                    atleastOne = True
                End If

                If Not isMilestone Then
                    ' test 1: Zugriff über ID 
                    cphase = hproj.getPhaseByID(curID)

                    If cphase.nameID = curID Then
                        ' ok 
                    Else
                        logMessage = logMessage & vbLf & kvp.Value.getShapeText & "Node-Zugriff über ID nicht ok " & curID & ", " & cphase.nameID
                        atleastOne = True
                    End If

                    ' Test2: Zugriff über Name und lfd-Nr 
                    elemID = calcHryElemKey(elemName, isMilestone, lfdNr)

                    cphase = hproj.getPhaseByID(elemID)

                    If cphase.nameID = elemID Then
                        ' ok 
                    Else
                        logMessage = logMessage & vbLf & kvp.Value.getShapeText & "Node-Zugriff über Elem-Name, lfdNr nicht ok " & curID & ", " & cphase.nameID
                        atleastOne = True
                    End If

                Else
                    cMilestone = hproj.getMilestoneByID(curID)

                    If cMilestone.nameID = curID Then
                        ' ok 
                    Else
                        logMessage = logMessage & vbLf & kvp.Value.getShapeText & "Node-Zugriff über ID nicht ok " & curID & ", " & cMilestone.nameID
                        atleastOne = True
                    End If

                    ' Test2: Zugriff über Name und lfd-Nr 
                    elemID = calcHryElemKey(elemName, isMilestone, lfdNr)

                    cMilestone = hproj.getMilestoneByID(elemID)

                    If cMilestone.nameID = elemID Then
                        ' ok 
                    Else
                        logMessage = logMessage & vbLf & kvp.Value.getShapeText & "Node-Zugriff über Elem-Name, lfdNr nicht ok " & curID & ", " & cMilestone.nameID
                        atleastOne = True
                    End If


                End If

                ' jetzt wird gecheckt, ob das Element einen parent hat - wenn ja, ob es auch das Kind des Parents ist   
                ' wenn ja, wird gecheckt, ob der Parent-Knoten das aktuelle Element in der Liste der Child-Knoten hat 

                parentID = curNode.parentNodeKey
                If parentID <> "" Then
                    parentNode = hproj.hierarchy.nodeItem(parentID)

                    If Not IsNothing(parentNode) Then
                        Dim found As Boolean
                        For cx As Integer = 1 To parentNode.childCount
                            If curID = parentNode.getChild(cx) Then
                                found = True
                            End If
                        Next
                        If Not found Then
                            logMessage = logMessage & vbLf & kvp.Value.getShapeText & "Eltern-Knoten hat mich nicht als Kind" & parentID & ", Kind:  " & curID
                        End If
                    Else
                        logMessage = logMessage & vbLf & kvp.Value.getShapeText & "hat keinen Eltern-Knoten: curID: " & curID & ", parentID " & parentID
                    End If
                End If

                ' jetzt wird gecheckt, ob das Element Kinder hat -  
                ' wenn ja, ob jedes Kind das Element als parent hat    

                For cx As Integer = 1 To curNode.childCount

                    childID = curNode.getChild(cx)
                    childNode = hproj.hierarchy.nodeItem(childID)
                    If Not childNode.parentNodeKey = curID Then
                        logMessage = logMessage & vbLf & kvp.Value.getShapeText & "Kind hat mich nicht als Vater:  " & childID & ", CurID " & curID
                    End If

                Next


            Next

            ' jetzt wird ausgehend von den Phasen und den zugehörigen Milestones gecheckt 

            For ix As Integer = 1 To hproj.CountPhases

                cphase = hproj.getPhase(ix)
                curID = cphase.nameID

                ' check in der Hierarchie
                cphase2 = hproj.getPhaseByID(curID)
                If Not IsNothing(cphase2) Then
                    If Not cphase2.nameID = cphase.nameID Then
                        logMessage = logMessage & vbLf & kvp.Value.getShapeText & "Zugriff über ix: " & ix & ": " & cphase.nameID & " <> " & cphase.nameID
                        atleastOne = True
                    End If
                End If


                ' jetzt werden die Meilensteine gecheckt
                For mx As Integer = 1 To cphase.countMilestones
                    cMilestone = cphase.getMilestone(mx)
                    curID = cMilestone.nameID
                    curNode = hproj.hierarchy.nodeItem(curID)
                    If curNode.indexOfElem <> mx Then
                        logMessage = logMessage & vbLf & kvp.Value.getShapeText & "Meilenstein-Zugriff über mx: " & ix & vbLf & _
                                     curNode.indexOfElem & " <> " & mx
                    End If

                    parentID = curNode.parentNodeKey
                    parentNode = hproj.hierarchy.nodeItem(parentID)
                    If Not IsNothing(parentNode) Then
                        If parentNode.indexOfElem <> ix Then
                            logMessage = logMessage & vbLf & kvp.Value.getShapeText & "Phasen-Zugriff über ix: " & ix & vbLf & _
                                         parentNode.indexOfElem & " <> " & mx
                        End If
                    Else
                        If parentID <> "" Then
                            logMessage = logMessage & vbLf & kvp.Value.getShapeText & "Phasen-Zugriff über ix: " & ix & vbLf & _
                                         curID & " hat keinen Parent " & parentID
                        End If
                    End If

                Next


            Next


        Next

        If atleastOne Or logMessage.Length > 1 Then
            atleastOne = False
            Call MsgBox(logMessage)
            logMessage = ""
        End If

        If Not atleastOne Then
            Call MsgBox("done ..")
        End If

        enableOnUpdate = True


    End Sub

#End Region

#Region "Hilfsprogramme"

    Private Shared Function GetResourceText(ByVal resourceName As String) As String
        Dim asm As Reflection.Assembly = Reflection.Assembly.GetExecutingAssembly()
        Dim resourceNames() As String = asm.GetManifestResourceNames()
        For i As Integer = 0 To resourceNames.Length - 1
            If String.Compare(resourceName, resourceNames(i), StringComparison.OrdinalIgnoreCase) = 0 Then
                Using resourceReader As IO.StreamReader = New IO.StreamReader(asm.GetManifestResourceStream(resourceNames(i)))
                    If resourceReader IsNot Nothing Then
                        Return resourceReader.ReadToEnd()
                    End If
                End Using
            End If
        Next
        Return Nothing
    End Function

#End Region

End Class
