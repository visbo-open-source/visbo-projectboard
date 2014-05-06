Imports ProjectBoardDefinitions
Imports MongoDbAccess
Imports ClassLibrary1
Imports WpfWindow
Imports WPFPieChart
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Excel


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


    Sub awinNeueKonstellation(control As IRibbonControl)
        Dim storeConstellationFrm As New frmStoreConstellation
        Dim returnValue As DialogResult

        enableOnUpdate = False
        returnValue = storeConstellationFrm.ShowDialog
        enableOnUpdate = True

    End Sub

    Sub awinLadenKonstellation(control As IRibbonControl)
        Dim loadConstellationFrm As New frmLoadConstellation
        Dim constellationName As String


        Dim returnValue As DialogResult
        enableOnUpdate = False

        returnValue = loadConstellationFrm.ShowDialog
        If returnValue = DialogResult.OK Then
            constellationName = loadConstellationFrm.ListBox1.Text
            Call awinLoadConstellation(constellationName)

            appInstance.ScreenUpdating = False
            'Call diagramsVisible(False)
            Call awinClearPlanTafel()
            Call awinZeichnePlanTafel()
            Call awinNeuZeichnenDiagramme(2)
            'Call diagramsVisible(True)
            appInstance.ScreenUpdating = True
            Call MsgBox(constellationName & " wurde geladen ...")

            ' setzen der public variable, welche Konstellation denn jetzt gesetzt ist
            currentConstellation = constellationName

        End If
        enableOnUpdate = True

    End Sub


    Sub awinSetModusHistory(control As IRibbonControl)

        demoModusHistory = Not demoModusHistory
        historicDate = #2/27/2014#
        historicDate = historicDate.AddHours(16)
        If demoModusHistory Then
            Call MsgBox("Demo Modus History: Ein")
        Else
            Call MsgBox("Demo Modus History: Aus")
        End If

    End Sub

    Sub PT5StoreProjects(control As IRibbonControl)

        Dim zeitStempel As Date

        If AlleProjekte.Count > 0 Then

            Call StoreAllProjectsinDB()

            zeitStempel = AlleProjekte.First.Value.timeStamp

            Call MsgBox("ok, gespeichert!" & vbLf & zeitStempel.ToShortDateString & ", " & zeitStempel.ToShortTimeString)

            ' Änderung 18.6 - wenn gespeichert wird, soll die Projekthistorie zurückgesetzt werden 
            Try
                If projekthistorie.Count > 0 Then
                    projekthistorie.clear()
                End If
            Catch ex As Exception

            End Try
        Else
            Call MsgBox("keine Projekte zu speichern ...")
        End If


    End Sub

    Sub PT6DeleteCharts(control As IRibbonControl)

        Dim anzDiagrams As Integer
        Dim chtobj As Excel.ChartObject
        Dim i As Integer = 0

        With appInstance.Worksheets(arrWsNames(3))

            anzDiagrams = .ChartObjects.Count

            While i < anzDiagrams

                chtobj = .ChartObjects(1)
                Call awinDeleteChart(chtobj)
                i = i + 1

            End While


        End With


    End Sub

    ''' <summary>
    ''' wird aktuell verwendet , um eine Stelle für Testen bestimmter Funktionalitäten zu haben
    ''' ohne dass eine neue Ribbon Erweiterung gemacht werden muss
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub awinTestNewFunctions(control As IRibbonControl)
        'Call MsgBox("Anzahl Aufrufe: " & anzahlCalls)
        Dim ok As Boolean = True


        For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

            If Not kvp.Value.isConsistent Then
                Call MsgBox("inkonsistenz: " & kvp.Key)
                ok = False
            End If

        Next

        If ok Then
            Call MsgBox("keine Inkonsistenz gefunden ...")
        End If

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
        'Dim SID As String

        Dim tmpshapes As Excel.Shapes
        Dim oldKey As String, newKey As String

        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False
        Dim erg As String = ""
        Dim atleastOne As Boolean = False

        enableOnUpdate = False

        Try
            tmpshapes = CType(appInstance.ActiveSheet.shapes, Excel.Shapes)
        Catch ex As Exception
            tmpshapes = Nothing
        End Try

        If Not tmpshapes Is Nothing Then

            ' jetzt die Aktion durchführen ...

            For Each singleShp In tmpshapes
                With singleShp
                    If .AutoShapeType = MsoAutoShapeType.msoShapeRoundedRectangle Or
                        (.AutoShapeType = MsoAutoShapeType.msoShapeMixed And Not .HasChart _
                         And Not .Connector = Microsoft.Office.Core.MsoTriState.msoTrue) Then

                        If .Name <> .TextFrame2.TextRange.Text Then
                            ' das Shape wurde vom Nutzer umbenannt 
                            atleastOne = True
                            erg = erg & .Name & " -> " & .TextFrame2.TextRange.Text & vbLf

                            Dim oldName As String = .Name
                            Dim newName As String = .TextFrame2.TextRange.Text

                            Try
                                If ShowProjekte.Liste.ContainsKey(newName) Or Len(newName.Trim) = 0 Or IsNumeric(newName) Then
                                    ' ungültiger Name - alten Namen wiederherstellen 
                                    .TextFrame2.TextRange.Text = oldName
                                    erg = erg & oldName & "bleibt, " & newName & "ungültig" & vbLf
                                Else
                                    ' der neue Name ist gültig 
                                    .Name = newName

                                    Dim hproj As clsProjekt = ShowProjekte.getProject(oldName)
                                    oldKey = hproj.name & "#" & hproj.variantName
                                    newKey = newName & "#" & hproj.variantName
                                    With hproj
                                        .name = newName
                                    End With

                                    ShowProjekte.Remove(oldName)
                                    hproj.timeStamp = Date.Now
                                    ShowProjekte.Add(hproj)
                                    AlleProjekte.Remove(oldKey)
                                    AlleProjekte.Add(newKey, hproj)

                                End If
                            Catch ex As Exception
                                Call MsgBox(ex.Message)
                                .TextFrame2.TextRange.Text = oldName
                                erg = erg & oldName & "bleibt, " & newName & "ungültig" & vbLf
                            End Try




                        End If

                    End If
                End With
            Next

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

        enableOnUpdate = False

        returnValue = ProjektEingabe.ShowDialog

        If returnValue = DialogResult.Yes Then
            With ProjektEingabe


                Call TrageivProjektein(.projectName.Text, .vorlagenDropbox.Text, CDate(.calcProjektStart), _
                                       Date.MinValue, CType(.Erloes.Text, Double), zeile, _
                                       CType(.sFit.Text, Double), CType(.risiko.Text, Double), CDbl(.volume.Text))
            End With
        End If

        If returnValue = DialogResult.No Then
            With ProjektEingabe


                Call TrageivProjektein(.projectName.Text, .vorlagenDropbox.Text, CDate(.calcProjektStart), _
                                       CDate(.calcProjektEnde), CType(.Erloes.Text, Double), zeile, _
                                       CType(.sFit.Text, Double), CType(.risiko.Text, Double), CDbl(.volume.Text))
            End With
        End If
        enableOnUpdate = True

    End Sub

    Sub PT5changeTimeSpan(control As IRibbonControl)

        Dim mvTimeSpan As New frmMoveTimeSpan
        'Dim returnValue As DialogResult

        appInstance.EnableEvents = False

        'returnValue = mvTimeSpan.Showdialog
        ' in dieser auskommentierten Variante ist es sehr langsam ... deshalb als modales Fenster

        mvTimeSpan.Show()

        appInstance.EnableEvents = True


    End Sub

    Sub PTDefineDependencies(control As IRibbonControl)

        Dim defineDependencies As New frmDependencies
        Dim result As DialogResult

        enableOnUpdate = False

        result = defineDependencies.ShowDialog()

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
        'Dim shpElement As Excel.Shape
        'Dim tmpShapes As Excel.Shapes
        Dim hproj As clsProjekt

        ' es wird vbeim Betreten der Tabelle2 nochmal auf False gesetzt ... und insbesondere bei Activate Tabelle1 (!) auf true gesetzt, nicht vorher wieder
        enableOnUpdate = False

        ' damit man was sieht
        'appInstance.ActiveSheet.screenupdating = True



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
                                    Call awinCreateBudgetWerte(hproj)
                                Else
                                    Try
                                        Call awinUpdateBudgetWerte(hproj, CType(ProjektAendern.Erloes.Text, Double))
                                        .Erloes = CType(ProjektAendern.Erloes.Text, Double)
                                    Catch ex As Exception
                                        .Erloes = CType(ProjektAendern.Erloes.Text, Double)
                                        Call awinCreateBudgetWerte(hproj)
                                    End Try

                                End If
                            End If

                            .StrategicFit = CType(ProjektAendern.sFit.Text, Double)
                            .Risiko = CType(ProjektAendern.risiko.Text, Double)

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
    ''' Projekt ins Noshow stellen  
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G1NoShow(control As IRibbonControl)

        Dim singleShp As Excel.Shape
        'Dim SID As String

        Dim awinSelection As Excel.ShapeRange

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
                With singleShp
                    If .AutoShapeType = MsoAutoShapeType.msoShapeRoundedRectangle Or
                        (.AutoShapeType = MsoAutoShapeType.msoShapeMixed And Not .HasChart _
                         And Not .Connector = Microsoft.Office.Core.MsoTriState.msoTrue) Then

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
    ''' Projekt ins Show zurückholen 
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G1Show(control As IRibbonControl)

        Dim getBackToShow As New frmGetProjectbackFromNoshow

        Dim returnValue As DialogResult
        enableOnUpdate = False
        appInstance.ScreenUpdating = False

        returnValue = getBackToShow.ShowDialog

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
                With singleShp
                    If .AutoShapeType = MsoAutoShapeType.msoShapeRoundedRectangle Or
                        (.AutoShapeType = MsoAutoShapeType.msoShapeMixed And Not .HasChart _
                         And Not .Connector = Microsoft.Office.Core.MsoTriState.msoTrue) Then
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
                With singleShp
                    If .AutoShapeType = MsoAutoShapeType.msoShapeRoundedRectangle Or
                        (.AutoShapeType = MsoAutoShapeType.msoShapeMixed And Not .HasChart _
                         And Not .Connector = Microsoft.Office.Core.MsoTriState.msoTrue) Then
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
                With singleShp
                    If .AutoShapeType = MsoAutoShapeType.msoShapeRoundedRectangle Or
                        (.AutoShapeType = MsoAutoShapeType.msoShapeMixed And Not .HasChart _
                         And Not .Connector = Microsoft.Office.Core.MsoTriState.msoTrue) Then
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
            Dim firstCall As Boolean = True
            For Each singleShp In awinSelection


                With singleShp
                    If .AutoShapeType = MsoAutoShapeType.msoShapeRoundedRectangle Or _
                        (.AutoShapeType = MsoAutoShapeType.msoShapeMixed And Not .HasChart _
                         And Not .Connector = Microsoft.Office.Core.MsoTriState.msoTrue) Then

                        Try
                            Call awinDeleteChartorProject(vprojektname:=.Name, firstCall:=firstCall)
                            firstCall = False
                        Catch ex As Exception
                            Exit For
                        End Try

                    End If
                End With


            Next

        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If

        Call awinDeSelect()

        enableOnUpdate = True
        appInstance.EnableEvents = formerEE

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

            ' Formular zum Auswählen der Report-Vorlage wird aufgerufen

            getReportVorlage.calledfrom = "Projekt"
            returnValue = getReportVorlage.ShowDialog

            appInstance.EnableEvents = True
            appInstance.ScreenUpdating = True
            enableOnUpdate = True
        End If

    End Sub

    Public Sub Tom2G4B1InventurImport(control As IRibbonControl)

        'Dim projektInventurFile As String = "Projekt-Inventur.xlsx"
        Dim projektInventurFile As String = requirementsOrdner & "Projekt-Inventur.xlsx"
        'Dim projektInventurFile As String = requirementsOrdner & "RPLAN Projekte.xlsx"
        Dim dateiName As String
        Dim myCollection As New Collection
        Dim importDate As Date = Date.Now

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
            Call importProjekteEintragen(myCollection, importDate)
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
                Call bmwImportProjektInventur(myCollection)
                appInstance.ActiveWorkbook.Close(SaveChanges:=False)
                Call importProjekteEintragen(myCollection, importDate)

            Catch ex As Exception
                appInstance.ActiveWorkbook.Close(SaveChanges:=False)
                Call MsgBox("Fehler bei Import " & vbLf & dateiName & vbLf & ex.Message)
                Exit Sub
            End Try
        Else
            Call MsgBox(" Import RPLAN-Projekte wurde abgebrochen")
        End If



        enableOnUpdate = True
        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True

    End Sub

    Public Sub Tom2G4M1Import(control As IRibbonControl)

        Dim request As New Request(awinSettings.databaseName)
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



        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False
        enableOnUpdate = False

        Dim myCollection As New Collection




        dirName = awinPath & projektFilesOrdner
        listOfVorlagen = My.Computer.FileSystem.GetFiles(dirName, FileIO.SearchOption.SearchTopLevelOnly, "*.xlsx")

        ' alle Import Projekte erstmal löschen
        ImportProjekte.Clear()


        ' jetzt müssen die Projekte ausgelesen werden, die in dateiListe stehen 
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
                            ImportProjekte.Add(hproj)
                            myCollection.Add(hproj.name)
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
            Call importProjekteEintragen(myCollection, importDate)
        Catch ex As Exception
            Call MsgBox("Fehler bei Import : " & vbLf & ex.Message)
        End Try




        enableOnUpdate = True
        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True




    End Sub



    Public Sub Tom2G4M1Export(control As IRibbonControl)

        Dim request As New Request(awinSettings.databaseName)
        Dim singleShp As Excel.Shape
        Dim hproj As clsProjekt
        Dim vglName As String = " "
        Dim outputString As String = ""


        Dim awinSelection As Excel.ShapeRange

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

                    With singleShp
                        If .AutoShapeType = MsoAutoShapeType.msoShapeRoundedRectangle Or _
                        (.AutoShapeType = MsoAutoShapeType.msoShapeMixed And Not .HasChart _
                         And Not .Connector = Microsoft.Office.Core.MsoTriState.msoTrue) Then

                            Try
                                hproj = ShowProjekte.getProject(singleShp.Name)

                                ' jetzt wird dieses Projekt exportiert ... 
                                Try
                                    Call awinExportProject(hproj)
                                    outputString = outputString & hproj.name & " erfolgreich .." & vbLf
                                Catch ex As Exception
                                    outputString = outputString & hproj.name & " nicht erfolgreich .." & vbLf
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

        Dim fileName As String, altfileName As String
        Dim zeile As Integer = 2
        Dim anzRollen As Integer
        Dim i As Integer



        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False
        enableOnUpdate = False



        ' für jede Ressource eine eigene Datei machen
        anzRollen = RoleDefinitions.Count

        Dim ok As Boolean = True
        Dim roleName As String

        For i = 1 To anzRollen

            roleName = RoleDefinitions.getRoledef(i).name.Trim
            fileName = roleName & "-Zuordnung.xlsx"

            ' öffnen der Excel Datei 
            Try
                appInstance.Workbooks.Open(awinPath & projektRessOrdner & "\" & fileName)
                ok = True
            Catch ex As Exception

                altfileName = "Vorlage Zuordnung.xlsx"

                Try
                    appInstance.Workbooks.Open(awinPath & projektRessOrdner & "\" & altfileName)
                    Try
                        appInstance.ActiveWorkbook.SaveAs(awinPath & projektRessOrdner & "\" & fileName, _
                                      ConflictResolution:=XlSaveConflictResolution.xlLocalSessionChanges)
                        ok = True
                    Catch ex2 As Exception
                        ok = False
                    End Try



                Catch ex1 As Exception
                    Call MsgBox("File " & altfileName & " nicht gefunden ... Abbruch")
                    appInstance.EnableEvents = True
                    appInstance.ScreenUpdating = True
                    enableOnUpdate = True
                    Exit Sub
                End Try

            End Try


            If ok Then

                Call awinExportRessZuordnung(1, roleName)


                Try
                    appInstance.ActiveWorkbook.Save()
                    appInstance.ActiveWorkbook.Close()

                Catch ex As Exception
                    appInstance.ActiveWorkbook.Close()
                End Try

            Else

                Call MsgBox("Fehler bei Save as " & fileName)

            End If




        Next

        Call MsgBox("ok, Dateien erstellt ...")

        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True
        enableOnUpdate = True


    End Sub

    Public Sub PT5phasenZeichnenInit(control As IRibbonControl, ByRef pressed As Boolean)

        pressed = ProjectBoardDefinitions.My.Settings.drawPhases

    End Sub

    Public Sub PT5phasenZeichnen(control As IRibbonControl, pressed As Boolean)

        If pressed Then
            ' jetzt werden die Projekt-Symbole inkl Phasen Darstellung gezeichnet
            ProjectBoardDefinitions.My.Settings.drawPhases = True
            Call awinClearPlanTafel()
            Call awinZeichnePlanTafel()
        Else
            ' jetzt werden die Projekt-Symbole ohne Phasen Darstellung gezeichnet 
            ProjectBoardDefinitions.My.Settings.drawPhases = False
            Call awinLoadConstellation("Last")
            Call awinClearPlanTafel()
            Call awinZeichnePlanTafel()
        End If

    End Sub

    Public Sub PT5loadprojectsInit(control As IRibbonControl, ByRef pressed As Boolean)

        pressed = ProjectBoardDefinitions.My.Settings.loadProjectsOnChange

    End Sub

    Public Sub PT5loadProjectsOnChange(control As IRibbonControl, ByRef pressed As Boolean)

        If pressed Then
            ' jetzt sollen die Projekte gemäß Time Span geladen werden - auch bei Veränderung des TimeSpan 
            ProjectBoardDefinitions.My.Settings.loadProjectsOnChange = True
            ' noch zu tun 
            ' Call awinloadProjectsFromDB()
        Else

            ' jetzt werden bei TimeSpan Änderung keine Projekte nachgeladen 
            ProjectBoardDefinitions.My.Settings.loadProjectsOnChange = False


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
                Call createPhasesBalken(noColorCollection, hproj, repObj, scale, top, left, height, width, " ")

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

                ' bestimme die Anzahl Zeilen, die benötigt wird  
                Dim anzahlZeilen As Integer = getNeededSpace(hproj)

                Call moveShapesDown(hproj.tfZeile + 1, anzahlZeilen)
                'Call ZeichneProjektinPlanTafel2(pname, hproj.tfZeile)
                Call ZeichneProjektinPlanTafel(pname, hproj.tfZeile)


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

                Dim repObj As Object
                appInstance.EnableEvents = False
                appInstance.ScreenUpdating = False

                repObj = Nothing

                width = hproj.Dauer * boxWidth + 10

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



                width = hproj.Dauer * boxWidth + 10

                appInstance.EnableEvents = False
                appInstance.ScreenUpdating = False
                Dim repObj As Object = Nothing

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

                width = hproj.Dauer * boxWidth + 10
                appInstance.EnableEvents = False
                appInstance.ScreenUpdating = False
                Dim repObj As Object = Nothing

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

                width = hproj.Dauer * boxWidth + 10

                appInstance.EnableEvents = False
                appInstance.ScreenUpdating = False
                Dim repObj As Object = Nothing


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
                With singleShp
                    If .AutoShapeType = MsoAutoShapeType.msoShapeRoundedRectangle Or _
                        (.AutoShapeType = MsoAutoShapeType.msoShapeMixed And Not .HasChart _
                         And Not .Connector = Microsoft.Office.Core.MsoTriState.msoTrue) Then

                        myCollection.Add(.Name)
                        top = .Top + boxHeight + 2
                        left = .Left - 3
                        width = 12 * boxWidth
                        height = 8 * boxHeight

                    End If
                End With
            Next
            Dim obj As New Object
            Call awinCreatePortfolioDiagramms(myCollection, obj, True, PTpfdk.FitRisiko, 0, False, True, True, top, left, width, height)
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
                With singleShp
                    If .AutoShapeType = MsoAutoShapeType.msoShapeRoundedRectangle Or _
                        (.AutoShapeType = MsoAutoShapeType.msoShapeMixed And Not .HasChart _
                         And Not .Connector = Microsoft.Office.Core.MsoTriState.msoTrue) Then

                        myCollection.Add(.Name)
                        top = .Top + boxHeight + 2
                        left = .Left - 3
                        width = 12 * boxWidth
                        height = 8 * boxHeight

                    End If
                End With
            Next
            Dim obj As New Object

            Call awinCreatePortfolioDiagramms(myCollection, obj, True, PTpfdk.FitRisikoVol, 0, False, True, True, top, left, width, height)
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
                With singleShp
                    If .AutoShapeType = MsoAutoShapeType.msoShapeRoundedRectangle Or _
                        (.AutoShapeType = MsoAutoShapeType.msoShapeMixed And Not .HasChart _
                         And Not .Connector = Microsoft.Office.Core.MsoTriState.msoTrue) Then

                        myCollection.Add(.Name, .Name)
                        top = .Top + boxHeight + 2
                        left = .Left - 3
                        width = 12 * boxWidth
                        height = 8 * boxHeight

                    End If
                End With
            Next

            For i = 1 To myCollection.Count
                pname = myCollection.Item(i)
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
                pname = deleteList.Item(i)
                Try
                    myCollection.Remove(pname)
                Catch ex As Exception

                End Try
            Next

            If myCollection.Count > 0 Then
                Dim obj As New Object
                Call awinCreatePortfolioDiagramms(myCollection, obj, True, PTpfdk.Dependencies, 0, False, True, True, top, left, width, height)
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
                With singleShp
                    If .AutoShapeType = MsoAutoShapeType.msoShapeRoundedRectangle Or _
                        (.AutoShapeType = MsoAutoShapeType.msoShapeMixed And Not .HasChart _
                         And Not .Connector = Microsoft.Office.Core.MsoTriState.msoTrue) Then

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
                    left = sichtbarerBereich.Left + (sichtbarerBereich.Width - 500) / 2
                    top = sichtbarerBereich.Top + (sichtbarerBereich.Height - 450) / 2
                End With

                width = 500
                height = 450
            End If

            Dim obj As New Object

            Try
                Call awinCreatePortfolioDiagramms(myCollection, obj, True, PTpfdk.ComplexRisiko, 0, False, True, True, top, left, width, height)
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

        Call awinSollIstVergleich(auswahl, typ, vglBaseline)

    End Sub


    Sub Tom2G2M2M3B2SollIstGKosten(control As IRibbonControl)

        ' auswahl steuert , welche Kosten angezeigt werden
        Dim auswahl As Integer = 3
        Dim vglBaseline As Boolean = True

        ' typ steuert, ob Summenbetrachtung oder Curve angezeigt wird
        Dim typ As String = " "

        Call awinSollIstVergleich(auswahl, typ, vglBaseline)

    End Sub

    ''' <summary>
    ''' Fortschritts-Chart im Vergleich zur Beauftragung
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G2M4Fortschritt1(control As IRibbonControl)

        Call awinStatusAnzeige(1, 1, " ")

    End Sub

    ''' <summary>
    ''' Fortschritts-Chart im Vergleich zur letzten Planungs-Freigabe
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G2M4Fortschritt2(control As IRibbonControl)

        Call awinStatusAnzeige(2, 1, " ")

    End Sub



    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="auswahl"></param>
    ''' <param name="typ"></param>
    ''' <remarks></remarks>
    Private Sub awinSollIstVergleich(ByVal auswahl As Integer, ByVal typ As String, ByVal vglBaseline As Boolean)
        Dim request As New Request(awinSettings.databaseName)
        Dim singleShp As Excel.Shape
        Dim hproj As clsProjekt
        Dim awinSelection As Excel.ShapeRange
        Dim top As Double, left As Double, width As Double, height As Double
        Dim reportObj As Excel.ChartObject
        Dim heute As Date = Date.Now
        Dim vglName As String = " "
        Dim pName As String = ";"
        Dim variantName As String = ""

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
                        vglName = projekthistorie.First.name
                    End If
                End If

                With hproj
                    pName = .name
                    variantName = .variantName
                End With

                If vglName.Trim <> pName.Trim Then
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
        Dim request As New Request(awinSettings.databaseName)
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
                With singleShp
                    If .AutoShapeType = MsoAutoShapeType.msoShapeRoundedRectangle Or _
                        (.AutoShapeType = MsoAutoShapeType.msoShapeMixed And Not .HasChart _
                         And Not .Connector = Microsoft.Office.Core.MsoTriState.msoTrue) Then

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


                Dim tmpObj As Object = Nothing
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


    ''' <summary>
    ''' zeige die grünen Milestones für die ausgewählten Projekte an 
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G2M5M2B1ShowMilestones(control As IRibbonControl)


        Dim farbTyp As Integer = 1
        Dim numberIt As Boolean = False
        Dim namelist As New SortedList(Of String, String)


        appInstance.EnableEvents = False
        'appInstance.ScreenUpdating = False
        enableOnUpdate = False

        Call awinZeichneMilestones(namelist, farbTyp, numberIt)

        enableOnUpdate = True
        appInstance.EnableEvents = True
        'appInstance.ScreenUpdating = formerSU


    End Sub

    Sub Tom2G2M5M2B2ShowMilestones(control As IRibbonControl)


        Dim farbTyp As Integer = 2
        Dim numberIt As Boolean = False
        Dim namelist As New SortedList(Of String, String)

        Dim formerSU As Boolean = appInstance.ScreenUpdating
        Dim formerEE As Boolean = appInstance.EnableEvents


        appInstance.EnableEvents = False
        enableOnUpdate = False

        Call awinZeichneMilestones(namelist, farbTyp, numberIt)

        enableOnUpdate = True
        appInstance.EnableEvents = formerEE


    End Sub

    Sub Tom2G2M5M2B3ShowMilestones(control As IRibbonControl)


        Dim farbTyp As Integer = 3
        Dim numberIt As Boolean = False
        Dim namelist As New SortedList(Of String, String)


        appInstance.EnableEvents = False
        enableOnUpdate = False

        Call awinZeichneMilestones(namelist, farbTyp, numberIt)

        enableOnUpdate = True
        appInstance.EnableEvents = True




    End Sub

    Sub Tom2G2M5M2B4ShowMilestones(control As IRibbonControl)


        Dim farbTyp As Integer = 0
        Dim numberIt As Boolean = False
        Dim namelist As New SortedList(Of String, String)



        appInstance.EnableEvents = False
        enableOnUpdate = False

        Call awinZeichneMilestones(namelist, farbTyp, numberIt)

        enableOnUpdate = True
        appInstance.EnableEvents = True



    End Sub

    Sub Tom2G2M5M2B5ShowMilestones(control As IRibbonControl)


        Dim farbTyp As Integer = 4
        Dim numberIt As Boolean = False
        Dim namelist As New SortedList(Of String, String)


        appInstance.EnableEvents = False
        enableOnUpdate = False

        Call awinZeichneMilestones(namelist, farbTyp, numberIt)

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
        Dim nameList As New SortedList(Of String, String)
        Dim title As String = "Meilensteine visualisieren"

        Dim repObj As Object = Nothing

        Dim singleShp As Excel.Shape
        Dim myCollection As New Collection
        Dim hproj As clsProjekt
        Dim awinSelection As Excel.ShapeRange
        Dim selektierteProjekte As New clsProjekte



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

                For Each kvp As KeyValuePair(Of String, String) In nameList
                    listOfItems.Add(kvp.Key)
                Next

                ' jetzt stehen in der listOfItems die Namen der Meilensteine - alphabetisch sortiert 
                Dim auswahlFenster As New ListSelectionWindow(listOfItems, title)


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
        Dim nameList As New SortedList(Of String, String)
        Dim title As String = "Meilensteine visualisieren"

        Dim repObj As Object = Nothing


        Call awinDeSelect()

        nameList = ShowProjekte.getMilestoneNames

        If nameList.Count > 0 Then

            For Each kvp As KeyValuePair(Of String, String) In nameList
                listOfItems.Add(kvp.Key)
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




    End Sub


    ''' <summary>
    ''' zeigt zu dem ausgewählten Projekt die Meilenstein Trendanalyse an 
    ''' dazu wird erst ein Fenster aufgeschaltet, aus dem der oder die Namen des betreffenden Meilensteins ausgewählt werden können 
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub PTShowMilestoneTrend(control As IRibbonControl)

        Dim request As New Request(awinSettings.databaseName)
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


        Try
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

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
                            vglName = projekthistorie.First.name
                        End If
                    End If

                    If vglName.Trim <> pName.Trim Then
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
                        width = System.Math.Max(hproj.Dauer * boxWidth + 10, 24 * boxWidth + 10)

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

        enableOnUpdate = True
        appInstance.EnableEvents = True





    End Sub

    Sub PT0ShowProjektStatus(control As IRibbonControl)

        Dim singleShp As Excel.Shape
        Dim myCollection As New Collection
        Dim hproj As clsProjekt
        Dim awinSelection As Excel.ShapeRange


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

            Call awinDeleteMilestoneShapes(4)

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
        Call awinDeleteMilestoneShapes(0)
    End Sub


    ''' <summary>
    ''' löscht alle angezeigten Milestones
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G2M5B3NoShowMilestones(control As IRibbonControl)

        Call awinDeleteMilestoneShapes(1)

    End Sub

    Sub PT0VisualizePhases(control As IRibbonControl)

        Dim i As Integer
        Dim von As Integer, bis As Integer

        Dim listOfItems As New Collection
        Dim existingNames As New SortedList(Of String, String)

        Dim repObj As Object = Nothing
        Dim title As String = "Phasen visualisieren"
        Dim phaseName As String
        Dim hproj As clsProjekt


        Dim awinSelection As Excel.ShapeRange
        Dim selektierteProjekte As New clsProjekte

        appInstance.EnableEvents = False
        enableOnUpdate = False

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


            existingNames = selektierteProjekte.getPhaseNames

            If existingNames.Count > 0 Then

                ' jetzt werden die Namen in der Reihenfolge, wie sie in der Phasen-Definition stehen in der listofItems eingetragen ..

                For i = 1 To PhaseDefinitions.Count
                    phaseName = PhaseDefinitions.getPhaseDef(i).name

                    If existingNames.ContainsKey(phaseName) Then
                        listOfItems.Add(PhaseDefinitions.getPhaseDef(i).name)
                    End If

                Next

                ' jetzt stehen in der listOfItems die Namen der Phasen 
                Dim auswahlFenster As New ListSelectionWindow(listOfItems, title)

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
        Dim existingNames As New SortedList(Of String, String)

        Dim repObj As Object = Nothing
        Dim title As String = "Phasen visualisieren"
        Dim phaseName As String

        Call awinDeSelect()

        existingNames = ShowProjekte.getPhaseNames

        ' jetzt werden die Namen in der Reihenfolge, wie sie in der Phasen-Definition stehen in der listofItems eingetragen ..

        For i = 1 To PhaseDefinitions.Count
            phaseName = PhaseDefinitions.getPhaseDef(i).name

            If existingNames.ContainsKey(phaseName) Then
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

    End Sub


    Sub PT0ShowPortfolioPhasen(control As IRibbonControl)

        Dim i As Integer
        Dim von As Integer, bis As Integer
        'Dim myCollection As Collection
        Dim listOfItems As New Collection
        'Dim left As Double, top As Double, height As Double, width As Double

        Dim repObj As Object = Nothing
        Dim phaseName As String



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



    End Sub

    Sub PTShowMilestoneSummen(control As IRibbonControl)

        Dim von As Integer, bis As Integer

        Dim listOfItems As New Collection

        Dim repObj As Object = Nothing
        Dim nameList As New SortedList(Of String, String)


        nameList = ShowProjekte.getMilestoneNames

        If nameList.Count > 0 Then

            For Each kvp As KeyValuePair(Of String, String) In nameList
                listOfItems.Add(kvp.Key)
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


    End Sub


    Sub PT0ShowAuslastung(control As IRibbonControl)

        Dim top As Double, left As Double, width As Double, height As Double
        Dim obj As Object = Nothing

        appInstance.ScreenUpdating = False
        appInstance.EnableEvents = False
        enableOnUpdate = False



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

        Dim repObj As Object = Nothing


        'appInstance.ScreenUpdating = False
        appInstance.EnableEvents = False
        enableOnUpdate = False


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


        'appInstance.ScreenUpdating = False
        'appInstance.EnableEvents = False
        'enableOnUpdate = False


        For i = 1 To RoleDefinitions.Count
            listOfItems.Add(RoleDefinitions.getRoledef(i).name)
        Next

        ' jetzt stehen in der listOfItems die Namen der Rollen 
        Dim auswahlFenster As New ListSelectionWindow(listOfItems, title)

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


        'appInstance.EnableEvents = False
        'enableOnUpdate = False


        For i = 1 To CostDefinitions.Count - 1
            listOfItems.Add(CostDefinitions.getCostdef(i).name)
        Next

        ' jetzt stehen in der listOfItems die Namen der Rollen 
        Dim auswahlFenster As New ListSelectionWindow(listOfItems, title)


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

        'If auswahlFenster.ShowDialog() Then

        '    myCollection = auswahlFenster.selectedItems
        '    von = showRangeLeft
        '    bis = showRangeRight

        '    height = awinSettings.ChartHoehe1
        '    top = WertfuerTop()

        '    If von > 1 Then
        '        'left = ((von - 1) / 3 - 1) * 3 * boxWidth + 32.8 + von * screen_correct
        '        left = (showRangeRight - 1) * boxWidth + 4
        '    Else
        '        left = 0
        '    End If

        '    width = 265 + (bis - von - 12 + 1) * boxWidth + (bis - von) * screen_correct

        '    Call awinCreateprcCollectionDiagram(myCollection, repObj, top, left, width, height, False, DiagrammTypen(2), False)

        'End If


        'appInstance.EnableEvents = True
        'enableOnUpdate = True

    End Sub

    Sub PT0ShowZieleUebersicht(control As IRibbonControl)

        Dim chtObject As Excel.ChartObject = Nothing
        'Dim top As Double, left As Double, width As Double, height As Double
        Dim future As Integer = 0

        Dim myCollection As New Collection
        myCollection.Add("Ziele")

        appInstance.EnableEvents = False
        enableOnUpdate = False

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
            .color = awinSettings.AmpelGelb
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

        appInstance.EnableEvents = True
        enableOnUpdate = True

    End Sub



    Sub PT0ShowStrategieRisiko(control As IRibbonControl)

        Dim selectionType As Integer = -1 ' keine Einschränkung
        Dim myCollection As New Collection
        Dim top As Double, left As Double, width As Double, height As Double
        Dim sichtbarerBereich As Excel.Range

        appInstance.EnableEvents = False
        enableOnUpdate = False

        myCollection = ShowProjekte.withinTimeFrame(selectionType, showRangeLeft, showRangeRight)

        With appInstance.ActiveWindow
            sichtbarerBereich = .VisibleRange
            left = sichtbarerBereich.Left + (sichtbarerBereich.Width - 600) / 2
            If left < sichtbarerBereich.Left Then
                left = sichtbarerBereich.Left + 2
            End If

            top = sichtbarerBereich.Top + (sichtbarerBereich.Height - 450) / 2
            If top < sichtbarerBereich.Top Then
                top = sichtbarerBereich.Top + 2
            End If

        End With

        width = 600
        height = 450

        Dim obj As New Object

        Try
            Call awinCreatePortfolioDiagramms(myCollection, obj, False, PTpfdk.FitRisiko, 0, False, True, True, top, left, width, height)
        Catch ex As Exception

        End Try

        appInstance.EnableEvents = True
        enableOnUpdate = True

    End Sub

    Sub PT0ShowStratRisikoVolume(control As IRibbonControl)

        Dim selectionType As Integer = -1 ' keine Einschränkung
        Dim myCollection As New Collection
        Dim top As Double, left As Double, width As Double, height As Double
        Dim sichtbarerBereich As Excel.Range

        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False
        enableOnUpdate = False

        myCollection = ShowProjekte.withinTimeFrame(selectionType, showRangeLeft, showRangeRight)

        With appInstance.ActiveWindow
            sichtbarerBereich = .VisibleRange
            left = sichtbarerBereich.Left + (sichtbarerBereich.Width - 600) / 2
            If left < sichtbarerBereich.Left Then
                left = sichtbarerBereich.Left + 2
            End If

            top = sichtbarerBereich.Top + (sichtbarerBereich.Height - 450) / 2
            If top < sichtbarerBereich.Top Then
                top = sichtbarerBereich.Top + 2
            End If

        End With

        width = 600
        height = 450

        Dim obj As New Object

        Try
            Call awinCreatePortfolioDiagramms(myCollection, obj, False, PTpfdk.FitRisikoVol, 0, False, True, True, top, left, width, height)
            'Call awinCreateStratRiskVolumeDiagramm(myCollection, obj, False, False, True, True, top, left, width, height)
        Catch ex As Exception

        End Try

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

        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False
        enableOnUpdate = False

        myCollection = ShowProjekte.withinTimeFrame(selectionType, showRangeLeft, showRangeRight)



        For i = 1 To myCollection.Count
            pname = myCollection.Item(i)
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
            pname = deleteList.Item(i)
            Try
                myCollection.Remove(pname)
            Catch ex As Exception

            End Try
        Next


        With appInstance.ActiveWindow
            sichtbarerBereich = .VisibleRange
            left = sichtbarerBereich.Left + (sichtbarerBereich.Width - 600) / 2
            If left < sichtbarerBereich.Left Then
                left = sichtbarerBereich.Left + 2
            End If

            top = sichtbarerBereich.Top + (sichtbarerBereich.Height - 450) / 2
            If top < sichtbarerBereich.Top Then
                top = sichtbarerBereich.Top + 2
            End If

        End With

        width = 600
        height = 450

        Dim obj As New Object

        Try
            If myCollection.Count > 0 Then
                Call awinCreatePortfolioDiagramms(myCollection, obj, False, PTpfdk.Dependencies, 0, False, True, True, top, left, width, height)
            Else
                Call MsgBox(" es gibt in diesem Zeitraum keine Projekte mit Abhängigkeiten")
            End If


        Catch ex As Exception

        End Try

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



        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False
        enableOnUpdate = False

        ' hier muss noch geklärt werden, welche Projekte betrachtet werden; es mcht keinen Sinn, 
        'das nur an den TimeFrame zu koppeln, es geht im wesentlichen um aktuell laufende und vergangene Projekte 
        ' Frage : was ist mit bereits beauftragten Projekten, die noch gar nicht begonnen haben, deren Planung aber bereits schlechter als beauftragt ist ? 

        selectionType = PTpsel.lfundab
        myCollection = ShowProjekte.withinTimeFrame(selectionType, showRangeLeft, showRangeRight)


        With appInstance.ActiveWindow
            sichtbarerBereich = .VisibleRange
            left = sichtbarerBereich.Left + (sichtbarerBereich.Width - 600) / 2
            If left < sichtbarerBereich.Left Then
                left = sichtbarerBereich.Left + 2
            End If

            top = sichtbarerBereich.Top + (sichtbarerBereich.Height - 450) / 2
            If top < sichtbarerBereich.Top Then
                top = sichtbarerBereich.Top + 2
            End If

        End With

        width = 600
        height = 450

        Dim obj As New Object

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



        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False
        enableOnUpdate = False

        ' hier muss noch geklärt werden, welche Projekte betrachtet werden; es mcht keinen Sinn, 
        'das nur an den TimeFrame zu koppeln, es geht im wesentlichen um aktuell laufende und vergangene Projekte 
        ' Frage : was ist mit bereits beauftragten Projekten, die noch gar nicht begonnen haben, deren Planung aber bereits schlechter als beauftragt ist ? 

        selectionType = PTpsel.lfundab
        myCollection = ShowProjekte.withinTimeFrame(selectionType, showRangeLeft, showRangeRight)


        With appInstance.ActiveWindow
            sichtbarerBereich = .VisibleRange
            left = sichtbarerBereich.Left + (sichtbarerBereich.Width - 600) / 2
            If left < sichtbarerBereich.Left Then
                left = sichtbarerBereich.Left + 2
            End If

            top = sichtbarerBereich.Top + (sichtbarerBereich.Height - 450) / 2
            If top < sichtbarerBereich.Top Then
                top = sichtbarerBereich.Top + 2
            End If

        End With

        width = 600
        height = 450

        Dim obj As New Object

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

        appInstance.EnableEvents = True
        enableOnUpdate = True
        appInstance.ScreenUpdating = True


    End Sub


    Sub PT0ShowComplexRisiko(control As IRibbonControl)

        Dim selectionType As Integer = -1 ' keine Einschränkung
        Dim myCollection As New Collection
        Dim top As Double, left As Double, width As Double, height As Double
        Dim sichtbarerBereich As Excel.Range

        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False
        enableOnUpdate = False

        myCollection = ShowProjekte.withinTimeFrame(selectionType, showRangeLeft, showRangeRight)

        With appInstance.ActiveWindow
            sichtbarerBereich = .VisibleRange
            left = sichtbarerBereich.Left + (sichtbarerBereich.Width - 600) / 2
            If left < sichtbarerBereich.Left Then
                left = sichtbarerBereich.Left + 2
            End If

            top = sichtbarerBereich.Top + (sichtbarerBereich.Height - 450) / 2
            If top < sichtbarerBereich.Top Then
                top = sichtbarerBereich.Top + 2
            End If

        End With

        width = 600
        height = 450


        Dim obj As New Object

        Try
            Call awinCreatePortfolioDiagramms(myCollection, obj, False, PTpfdk.ComplexRisiko, 0, False, True, True, top, left, width, height)
        Catch ex As Exception

        End Try

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

        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False
        enableOnUpdate = False

        myCollection = ShowProjekte.withinTimeFrame(selectionType, showRangeLeft, showRangeRight)

        With appInstance.ActiveWindow
            sichtbarerBereich = .VisibleRange
            left = sichtbarerBereich.Left + (sichtbarerBereich.Width - 600) / 2
            If left < sichtbarerBereich.Left Then
                left = sichtbarerBereich.Left + 2
            End If

            top = sichtbarerBereich.Top + (sichtbarerBereich.Height - 450) / 2
            If top < sichtbarerBereich.Top Then
                top = sichtbarerBereich.Top + 2
            End If

        End With

        width = 600
        height = 450


        Dim obj As New Object

        Try
            Call awinCreatePortfolioDiagramms(myCollection, obj, False, PTpfdk.ZeitRisiko, 0, False, True, True, top, left, width, height)
        Catch ex As Exception

        End Try

        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True
        enableOnUpdate = True

        Call awinDeSelect()

    End Sub

    Sub PT0ShowPortfolioBudgetCost(control As IRibbonControl)
        Dim selectionType As Integer = -1 ' keine Einschränkung
        Dim top As Double, left As Double, width As Double, height As Double

        appInstance.EnableEvents = False
        enableOnUpdate = False

        Dim sichtbarerBereich As Excel.Range

        height = awinSettings.ChartHoehe2
        width = 450

        With appInstance.ActiveWindow
            sichtbarerBereich = .VisibleRange
            left = sichtbarerBereich.Left + (sichtbarerBereich.Width - width) / 2
            If left < sichtbarerBereich.Left Then
                left = sichtbarerBereich.Left + 2
            End If

            top = sichtbarerBereich.Top + (sichtbarerBereich.Height - height) / 2
            If top < sichtbarerBereich.Top Then
                top = sichtbarerBereich.Top + 2
            End If

        End With



        Dim obj As Object = Nothing
        Call awinCreateBudgetErgebnisDiagramm(obj, top, left, width, height, False, False)


        appInstance.EnableEvents = True
        enableOnUpdate = True
    End Sub


    Sub PT0ShowPortfolioErgebnis(control As IRibbonControl)
        Dim selectionType As Integer = -1 ' keine Einschränkung
        Dim top As Double, left As Double, width As Double, height As Double

        appInstance.EnableEvents = False
        enableOnUpdate = False

        Dim sichtbarerBereich As Excel.Range

        height = awinSettings.ChartHoehe2
        width = 450

        With appInstance.ActiveWindow
            sichtbarerBereich = .VisibleRange
            left = sichtbarerBereich.Left + (sichtbarerBereich.Width - width) / 2
            If left < sichtbarerBereich.Left Then
                left = sichtbarerBereich.Left + 2
            End If

            top = sichtbarerBereich.Top + (sichtbarerBereich.Height - height) / 2
            If top < sichtbarerBereich.Top Then
                top = sichtbarerBereich.Top + 2
            End If

        End With



        Dim obj As Object = Nothing
        Call awinCreateErgebnisDiagramm(obj, top, left, width, height, False, False)


        appInstance.EnableEvents = True
        enableOnUpdate = True
    End Sub



    Sub Tom2G2M5M1B3ShowStatus(control As IRibbonControl)

        appInstance.EnableEvents = False
        enableOnUpdate = False

        Dim nummerieren As Boolean = False
        Call awinZeichneStatus(nummerieren)

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

                Dim dummyObj As New Object
                Dim hproj As clsProjekt
                Try
                    hproj = ShowProjekte.getProject(singleShp.Name)
                    Call createProjektErgebnisCharakteristik2(hproj, dummyObj, 2)
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
                Dim hproj As clsProjekt = ShowProjekte.getProject(singleShp.Name)
                Dim cproj As New clsProjekt
                Dim top As Double = singleShp.Top + boxHeight + 2
                Dim left As Double = singleShp.Left - boxWidth
                If left <= 0 Then
                    left = 1
                End If
                Call awinCompareProject(hproj, cproj, 0, top, left)
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

                width = System.Math.Max(hproj.Dauer * boxWidth + 7, cproj.Dauer * boxWidth + 7)
                scale = System.Math.Max(hproj.dauerInDays, cproj.dauerInDays)
                'width = hproj1.Dauer * boxWidth + 7
                'scale = hproj1.Dauer

                Dim repObj As Object
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
        Dim hproj As clsProjekt, cproj As clsProjekt
        Dim top As Double, left As Double, width As Double, height As Double
        Dim scale As Double
        Dim noColorCollection As New Collection

        Dim awinSelection As Excel.ShapeRange

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
                    cproj = New clsProjekt
                    vproj.CopyTo(cproj)
                    cproj.startDate = hproj.startDate

                Catch ex As Exception
                    Call MsgBox("Vorlage / Projekt nicht gefunden ...")
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
                Call createPhasesBalken(noColorCollection, hproj, repObj, scale, top, left, height, width, " ")

                With repObj
                    top = .Top + .Height + 3
                End With


                repObj = Nothing
                Call createPhasesBalken(noColorCollection, cproj, repObj, scale, top, left, height, width, "Vorlage")
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
                Call createPhasesBalken(noColorCollection, hproj, repObj, scale, top, left, height, width, " ")

                With repObj
                    top = .Top + .Height + 3
                End With


                repObj = Nothing
                Call createPhasesBalken(noColorCollection, cproj, repObj, scale, top, left, height, width, " ")
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

        Dim request As New Request(awinSettings.databaseName)
        Dim singleShp1 As Excel.Shape
        Dim hproj As clsProjekt, cproj As clsProjekt
        Dim top As Double, left As Double, width As Double, height As Double
        Dim scale As Double
        Dim noColorCollection As New Collection
        Dim vglName As String = " "
        Dim pName As String, variantName As String

        Dim awinSelection As Excel.ShapeRange

        enableOnUpdate = False

        Try
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

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
                        vglName = projekthistorie.First.name
                    End If
                End If

                With hproj
                    pName = .name
                    variantName = .variantName
                End With

                If vglName.Trim <> pName.Trim Then
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
                    Call createPhasesBalken(noColorCollection, hproj, repObj, scale, top, left, height, width, " ")

                    With repObj
                        top = .Top + .Height + 3
                    End With


                    repObj = Nothing
                    Call createPhasesBalken(noColorCollection, cproj, repObj, scale, top, left, height, width, "letzter Stand")

                    appInstance.ScreenUpdating = True

                End If




            Else
                Call MsgBox("bitte nur ein Projekt selektieren")

            End If
        Else
            Call MsgBox("ein Projekt selektieren, um es mit seinem letzten Stand zu vergleichen")
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

        Dim request As New Request(awinSettings.databaseName)
        Dim singleShp1 As Excel.Shape
        Dim hproj As clsProjekt, cproj As clsProjekt
        Dim top As Double, left As Double, width As Double, height As Double
        Dim scale As Double
        Dim noColorCollection As New Collection
        Dim vglName As String = " "
        Dim pName As String, variantName As String

        Dim awinSelection As Excel.ShapeRange

        enableOnUpdate = False

        Try
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

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
                        vglName = projekthistorie.First.name
                    End If
                End If

                With hproj
                    pName = .name
                    variantName = .variantName
                End With

                If vglName.Trim <> pName.Trim Then
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
                        Call createPhasesBalken(noColorCollection, hproj, repObj, scale, top, left, height, width, " ")

                        With repObj
                            top = .Top + .Height + 3
                        End With


                        repObj = Nothing
                        Call createPhasesBalken(noColorCollection, cproj, repObj, scale, top, left, height, width, "Beauftragung")

                    Catch ex As Exception

                        Call MsgBox("es ist kein Beauftragungs-Stand vorhanden")

                    End Try


                End If

            Else
                Call MsgBox("bitte nur ein Projekt selektieren")

            End If
        Else
            Call MsgBox("ein Projekt selektieren, um es mit seinem letzten Stand zu vergleichen")
        End If

        enableOnUpdate = True
        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True

    End Sub

    Sub Tom2G3M1B2ResourceVgl(control As IRibbonControl)

        Dim singleShp1 As Excel.Shape, singleShp2 As Excel.Shape
        'Dim SID As String

        Dim awinSelection As Excel.ShapeRange

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
                Dim hproj As clsProjekt = ShowProjekte.getProject(singleShp1.Name)
                Dim cproj As clsProjekt = ShowProjekte.getProject(singleShp2.Name)
                Dim top As Double = singleShp1.Top + boxHeight + 2
                Dim left As Double = singleShp1.Left - boxWidth
                If left <= 0 Then
                    left = 1
                End If
                Call awinCompareProject(hproj, cproj, 3, top, left)
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
        Dim request As New Request(awinSettings.databaseName)
        Dim singleShp As Excel.Shape
        Dim showCharacteristics As New frmShowProjCharacteristics
        'Dim returnValue As DialogResult
        Dim awinSelection As Excel.ShapeRange
        Dim top As Double, left As Double, height As Double, width As Double
        Dim vglName As String = " "

        enableOnUpdate = False
        appInstance.ScreenUpdating = False


        Try
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            If awinSelection.Count = 1 Then
                ' jetzt die Aktion durchführen ...
                singleShp = awinSelection.Item(1)


                hproj = ShowProjekte.getProject(singleShp.Name)
                With hproj
                    pName = .name
                    variantName = .variantName
                End With

                If Not projekthistorie Is Nothing Then
                    If projekthistorie.Count > 0 Then
                        vglName = projekthistorie.First.name
                    End If
                End If

                If vglName.Trim <> pName.Trim Then
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
                    Dim repObj As Object = Nothing
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

        enableOnUpdate = True
        appInstance.ScreenUpdating = True




    End Sub


    Sub awinShowTrendKPI(control As IRibbonControl)
        Dim hproj As clsProjekt
        Dim pName As String, variantName As String
        Dim request As New Request(awinSettings.databaseName)
        Dim singleShp As Excel.Shape
        Dim showCharacteristics As New frmShowProjCharacteristics
        'Dim returnValue As DialogResult
        Dim awinSelection As Excel.ShapeRange
        Dim top As Double, left As Double, height As Double, width As Double
        Dim vglName As String = " "

        enableOnUpdate = False
        appInstance.ScreenUpdating = False


        Try
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            If awinSelection.Count = 1 Then
                ' jetzt die Aktion durchführen ...
                singleShp = awinSelection.Item(1)


                hproj = ShowProjekte.getProject(singleShp.Name)
                With hproj
                    pName = .name
                    variantName = .variantName
                End With

                If Not projekthistorie Is Nothing Then
                    If projekthistorie.Count > 0 Then
                        vglName = projekthistorie.First.name
                    End If
                End If

                If vglName.Trim <> pName.Trim Then
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
                    Dim repObj As Object = Nothing
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
        Dim request As New Request(awinSettings.databaseName)
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
                        vglName = projekthistorie.First.name
                    End If
                End If

                If vglName.Trim <> pName.Trim Then
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

                    With showCharacteristics

                        .Text = "Historie für Projekt " & pName.Trim
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

        getReportVorlage.calledfrom = "Portfolio"

        enableOnUpdate = False
        appInstance.ScreenUpdating = False

        ' Formular zum Auswählen der Report-Vorlage wird aufgerufen

        returnValue = getReportVorlage.ShowDialog

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
