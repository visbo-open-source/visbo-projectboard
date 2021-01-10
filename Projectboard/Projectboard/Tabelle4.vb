
Imports ProjectBoardDefinitions
Imports Microsoft.Office.Interop.Excel

Public Class Tabelle4


    Private columnDsc As Integer = 6
    Private oldColumn As Integer = 5
    Private oldRow As Integer = 2
    Private columnName As Integer = 2

    Private Sub Tabelle4_ActivateEvent() Handles Me.ActivateEvent
        ' in der Mass-Edit Termine sollen Header und Formular-Bar immer erhalten bleiben ...
        Application.DisplayFormulaBar = True

        'Dim filterRange As Excel.Range
        Dim formerEE As Boolean = Application.EnableEvents
        Application.EnableEvents = False

        Dim meWS As Excel.Worksheet =
            CType(CType(appInstance.Workbooks(myProjektTafel), Excel.Workbook) _
            .Worksheets(arrWsNames(ptTables.meAT)), Excel.Worksheet)


        ' jetzt den Schutz aufheben , falls einer definiert ist 
        If meWS.ProtectContents Then
            meWS.Unprotect(Password:="x")
        End If

        Try
            ' die Anzahl maximaler Zeilen bestimmen 
            With visboZustaende
                .meMaxZeile = CType(meWS, Excel.Worksheet).UsedRange.Rows.Count
                .meColRC = -1 ' keine Bedeutung
                .meColSD = -1 ' keine Bedeutung 
                .meColED = -1
                .meColpName = 2

            End With

        Catch ex As Exception
            Call MsgBox("Fehler in Laden des Sheets ...")
        End Try

        ' jetzt den AutoFilter setzen 
        Try

            ' jetzt die Autofilter aktivieren ... 
            If Not CType(meWS, Excel.Worksheet).AutoFilterMode = True Then
                'CType(meWS, Excel.Worksheet).Cells(1, 1).AutoFilter()
                CType(meWS, Excel.Worksheet).Rows(1).AutoFilter()


            End If

        Catch ex As Exception
            Call MsgBox("Fehler beim Filtersetzen und Speichern" & vbLf & ex.Message)
        End Try

        Try
            If awinSettings.meEnableSorting Then

                With CType(meWS, Excel.Worksheet)
                    ' braucht man nicht mehr - ist schon gemacht 
                    '.Unprotect("x")
                    .EnableSelection = XlEnableSelection.xlNoRestrictions
                End With
            Else
                With meWS
                    .Protect(Password:="x", UserInterfaceOnly:=True,
                             AllowFormattingCells:=True,
                             AllowFormattingColumns:=True,
                             AllowInsertingColumns:=False,
                             AllowInsertingRows:=False,
                             AllowDeletingColumns:=False,
                             AllowDeletingRows:=False,
                             AllowSorting:=True,
                             AllowFiltering:=True)
                    .EnableSelection = Excel.XlEnableSelection.xlNoRestrictions
                    .EnableAutoFilter = True
                End With
            End If


        Catch ex As Exception

        End Try

        ' jetzt die Gridline zeigen
        With appInstance.ActiveWindow
            If massColFontValues(2, 0) <> 0 Then
                .Zoom = massColFontValues(2, 0)
            End If

            .DisplayGridlines = True
            .GridlineColor = Excel.XlRgbColor.rgbBlack
        End With

        If Not IsNothing(appInstance.ActiveCell) Then
            visboZustaende.oldValue = CStr(CType(appInstance.ActiveCell, Excel.Range).Value)
        End If


        Application.EnableEvents = formerEE
        If Application.ScreenUpdating = False Then
            Application.ScreenUpdating = True
        End If


    End Sub

    Private Sub Tabelle4_Deactivate() Handles Me.Deactivate

        appInstance.ActiveWindow.SplitColumn = 0
        appInstance.ActiveWindow.SplitRow = 0
        appInstance.DisplayFormulaBar = False

    End Sub

    Private Sub Tabelle4_SelectionChange(Target As Range) Handles Me.SelectionChange

        appInstance.EnableEvents = False

        Dim meWS As Excel.Worksheet = CType(appInstance.ActiveSheet, Excel.Worksheet)
        Dim pname As String = ""
        Dim oldPname As String = ""
        Dim vName As String = ""


        Try
            ' wenn mehr wie eine Zelle selektiert wurde ...
            If Target.Cells.Count > 1 Then
                Target = CType(Target.Cells(1, 1), Excel.Range)
                Target.Select()
            End If

            ' jetzt die Werte merken 
            If Not IsNothing(CType(Target.Cells(1, 1), Excel.Range).Value) Then
                visboZustaende.oldValue = CStr(CType(Target.Cells(1, 1), Excel.Range).Value)
            Else
                visboZustaende.oldValue = ""
            End If

            ' alte Row merken 
            visboZustaende.oldRow = Target.Row

        Catch ex As Exception
            Call MsgBox("Fehler bei Selection Change, Massen-Edit Attribute" & vbLf & ex.Message)
            appInstance.EnableEvents = True

        End Try

        ' in oldRow muss jetzt der entsprechende Projekt-Name ausgelsen werden .. 
        ' folgende Bedingung muss gesichert sein: alle Projekte, die in MassEdit aufgeführt sind , 
        ' sind sowohl in Showprojekte als auch in dbCacheProjekte
        Dim pNameChanged As Boolean = False

        With visboZustaende
            pname = CStr(CType(meWS.Cells(Target.Row, visboZustaende.meColpName), Excel.Range).Value)
            If Not IsNothing(meWS.Cells(Target.Row, visboZustaende.meColpName + 1).value) Then
                vName = CStr(meWS.Cells(Target.Row, visboZustaende.meColpName + 1).value)
            End If


            If IsNothing(.currentProject) Then
                ' es wurde bisher kein lastProject geladen 
                If ShowProjekte.contains(pname) Then
                    .currentProject = ShowProjekte.getProject(pname)
                    .currentProjectinSession = sessionCacheProjekte.getProject(calcProjektKey(pname, .currentProject.variantName))
                    pNameChanged = True
                End If

            ElseIf pname <> .currentProject.name Then
                ' muss neu geholt werden 
                If ShowProjekte.contains(pname) Then
                    .currentProject = ShowProjekte.getProject(pname)
                    .currentProjectinSession = sessionCacheProjekte.getProject(calcProjektKey(pname, .currentProject.variantName))
                    pNameChanged = True
                End If
            End If

        End With



        appInstance.EnableEvents = True

    End Sub

    Private Sub Tabelle4_Change(Target As Range) Handles Me.Change
        ' damit nicht eine immerwährende Event Orgie durch Änderung in den Zellen abgeht ...
        appInstance.EnableEvents = False

        Dim currentCell As Excel.Range = Target

        Try

            Dim meWB As Excel.Workbook = CType(appInstance.Workbooks.Item(myProjektTafel), Excel.Workbook)
            Dim meWS As Excel.Worksheet = CType(appInstance.ActiveSheet, Excel.Worksheet)

            If Target.Cells.Count = 1 Then

                Dim zeile As Integer = Target.Row
                Dim spalte As Integer = Target.Column
                Dim eingabe As String = ""

                If Not IsNothing(Target.Value) Then
                    eingabe = CStr(Target.Value).Trim
                End If

                Dim pName As String = CStr(CType(meWS.Cells(Target.Row, visboZustaende.meColpName), Excel.Range).Value)
                Dim vName As String = ""
                If Not IsNothing(meWS.Cells(Target.Row, visboZustaende.meColpName + 1).value) Then
                    vName = CStr(meWS.Cells(Target.Row, visboZustaende.meColpName + 1).value)
                End If

                Dim hproj As clsProjekt = Nothing

                If ShowProjekte.contains(pName) Then
                    hproj = ShowProjekte.getProject(pName)
                End If

                If Not IsNothing(hproj) Then
                    ' je nachdem, welche User Rolle ...  
                    If myCustomUserRole.customUserRole = ptCustomUserRoles.OrgaAdmin Then
                        Select Case spalte

                            Case 1 ' Projekt-Nummer, String
                                hproj.kundenNummer = eingabe

                            Case 4 ' Varianten Beschreibung 
                                hproj.variantDescription = eingabe

                            Case 5 ' Business Unit
                                hproj.businessUnit = eingabe

                            Case 6 ' Ziele 
                                hproj.description = eingabe

                            Case 7 ' Budget, numerische Zahl >= 0 
                                Try
                                    If CDbl(eingabe) >= 0 Then
                                        hproj.Erloes = CDbl(eingabe)
                                    Else
                                        Dim errtxt As String = "Wert muss >= 0 sein"
                                        If awinSettings.englishLanguage Then
                                            errtxt = "value must not be < 0"
                                        End If
                                        Call MsgBox(errtxt)
                                        Target.Value = visboZustaende.oldValue
                                    End If
                                Catch ex As Exception
                                    Target.Value = visboZustaende.oldValue
                                End Try

                            Case 8 ' Verantwortlich, String
                                hproj.leadPerson = eingabe

                            Case 9 ' Strat Fit, Zahl > 0 , < 10.0
                                Try
                                    If CDbl(eingabe) > 0 And CDbl(eingabe) <= 10 Then
                                        hproj.StrategicFit = CDbl(eingabe)
                                    Else
                                        Dim errtxt As String = "Wert muss zwischen 0 und 10 liegen"
                                        If awinSettings.englishLanguage Then
                                            errtxt = "value need to be > 0 and <= 10"
                                        End If
                                        Call MsgBox(errtxt)
                                        Target.Value = visboZustaende.oldValue
                                    End If
                                Catch ex As Exception
                                    Target.Value = visboZustaende.oldValue
                                End Try

                            Case 10 ' Risk, Zahl > 0 , <10
                                Try
                                    If CDbl(eingabe) > 0 And CDbl(eingabe) <= 10 Then
                                        hproj.Risiko = CDbl(eingabe)
                                    Else
                                        Dim errtxt As String = "Wert muss zwischen 0 und 10 liegen"
                                        If awinSettings.englishLanguage Then
                                            errtxt = "value need to be > 0 and <= 10"
                                        End If
                                        Call MsgBox(errtxt)
                                        Target.Value = visboZustaende.oldValue
                                    End If
                                Catch ex As Exception
                                    Target.Value = visboZustaende.oldValue
                                End Try

                            Case Else ' es handelt sich um Custom Fields
                                Try
                                    Dim cfName As String = CStr(CType(meWS.Cells(1, spalte), Excel.Range).Value)
                                    Dim cfUid As Integer = customFieldDefinitions.getUid(cfName)
                                    If cfUid >= 0 Then
                                        hproj.addSetCustomSField(cfUid, eingabe)
                                    End If
                                Catch ex As Exception
                                    Target.Value = visboZustaende.oldValue
                                End Try


                        End Select
                    Else
                        ' es werden nur Ampel, Erläuterung und Risiko gezeigt 
                        Select Case spalte
                            Case 4 ' Ampel
                                Try
                                    If CInt(eingabe) >= 0 And CInt(eingabe) <= 3 Then
                                        hproj.ampelStatus = CInt(eingabe)

                                        Select Case hproj.ampelStatus
                                            Case 0
                                                CType(Target, Excel.Range).Interior.Color = visboFarbeNone
                                            Case 1
                                                CType(Target, Excel.Range).Interior.Color = visboFarbeGreen
                                            Case 2
                                                CType(Target, Excel.Range).Interior.Color = visboFarbeYellow
                                            Case 3
                                                CType(Target, Excel.Range).Interior.Color = visboFarbeRed
                                        End Select

                                    Else
                                        Dim errtxt As String = "Wert muss zwischen 0 und 3 liegen"
                                        If awinSettings.englishLanguage Then
                                            errtxt = "value need to be > 0 and <= 3"
                                        End If
                                        Call MsgBox(errtxt)
                                        Target.Value = visboZustaende.oldValue
                                    End If
                                Catch ex As Exception
                                    Target.Value = visboZustaende.oldValue
                                End Try

                            Case 5 ' Ampel-Beschreibung
                                hproj.ampelErlaeuterung = eingabe

                            Case 6 ' Risiko Beschreibung
                                Try
                                    Dim cfUid As Integer = customFieldDefinitions.getUid("Risiko")
                                    If cfUid >= 0 Then
                                        hproj.addSetCustomSField(cfUid, eingabe)
                                    Else
                                        Target.Value = visboZustaende.oldValue
                                    End If
                                Catch ex As Exception
                                    Target.Value = visboZustaende.oldValue
                                End Try


                            Case Else
                                Target.Value = visboZustaende.oldValue
                        End Select
                    End If

                Else
                    Call MsgBox("Fehler: Projekt nicht gefunden: " & pName)
                End If


            ElseIf Target.Rows.Count > 1 Then

                'appInstance.Undo()
                'Call MsgBox("bitte nur eine Zelle selektieren ...")
                appInstance.Undo()
                'Target.Cells(1, 1).value = visboZustaende.oldValue
            End If


        Catch ex As Exception
            Call MsgBox("Fehler bei Massen-Edit, Ändern : " & vbLf & ex.Message)
        End Try

        appInstance.EnableEvents = True
    End Sub
End Class
